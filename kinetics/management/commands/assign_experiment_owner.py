from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from kinetics.access import is_shared_example, shared_example_filter
from kinetics.models import Experiment


class Command(BaseCommand):
    help = (
        "Явно назначить владельца одной старой работе по ID эксперимента и пользователя. "
        "Работа должна быть без владельца и не должна быть общим демонстрационным примером."
    )

    def add_arguments(self, parser):
        parser.add_argument("experiment_id", type=int)
        parser.add_argument("user_id", type=int)

    def handle(self, *args, experiment_id, user_id, **options):
        with transaction.atomic():
            experiment = (
                Experiment.objects.select_for_update()
                .filter(pk=experiment_id)
                .first()
            )
            if experiment is None:
                raise CommandError(f"Эксперимент с ID {experiment_id} не найден.")
            if experiment.owner_id is not None:
                raise CommandError(
                    f"У эксперимента {experiment_id} уже есть владелец "
                    f"(user ID {experiment.owner_id}); назначение отменено."
                )
            if is_shared_example(experiment):
                raise CommandError(
                    f"Эксперимент {experiment_id} является общим демонстрационным "
                    "примером; назначение владельца запрещено."
                )

            user_model = get_user_model()
            try:
                user = user_model._default_manager.get(pk=user_id)
            except user_model.DoesNotExist as exc:
                raise CommandError(f"Пользователь с ID {user_id} не найден.") from exc

            # Keep the safety checks in the UPDATE too, so a concurrent change
            # cannot assign an already-owned or newly-marked example.
            updated = (
                Experiment.objects.filter(pk=experiment_id, owner__isnull=True)
                .exclude(shared_example_filter())
                .update(owner_id=user.pk)
            )
            if updated != 1:
                raise CommandError(
                    f"Эксперимент {experiment_id} изменился во время операции; "
                    "владелец не назначен."
                )

        self.stdout.write(self.style.SUCCESS(
            f"Эксперимент {experiment_id} назначен пользователю {user.pk} ({user})."
        ))

