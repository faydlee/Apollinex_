"""Experiment visibility and editing rules for app views and the admin."""

from django.conf import settings
from django.db.models import Q

from .examples import SYNTHETIC_EXAMPLES
from .models import Experiment


SHARED_EXAMPLE_SOURCE_FILES = frozenset(
    example["source_file"] for example in SYNTHETIC_EXAMPLES.values()
)
SHARED_EXAMPLE_NAMES = frozenset(
    example["initial"]["name"] for example in SYNTHETIC_EXAMPLES.values()
)
SHARED_EXAMPLE_DESCRIPTIONS = frozenset(
    example["initial"]["description"] for example in SYNTHETIC_EXAMPLES.values()
)


def shared_example_filter():
    """Match built-in examples, including legacy rows with older metadata."""

    return (
        Q(source_format=Experiment.SourceFormat.SYNTHETIC_RAW)
        | Q(source_file__in=SHARED_EXAMPLE_SOURCE_FILES)
        | Q(name__in=SHARED_EXAMPLE_NAMES)
        | Q(description__in=SHARED_EXAMPLE_DESCRIPTIONS)
    )


def is_shared_example(experiment):
    return (
        experiment.source_format == Experiment.SourceFormat.SYNTHETIC_RAW
        or experiment.source_file in SHARED_EXAMPLE_SOURCE_FILES
        or experiment.name in SHARED_EXAMPLE_NAMES
        or experiment.description in SHARED_EXAMPLE_DESCRIPTIONS
    )


def can_edit_experiment(user, experiment):
    if user.is_authenticated:
        return experiment.owner_id == user.pk
    return (
        settings.APOLLINEX_ALLOW_GUEST_UNOWNED_EXPERIMENTS
        and experiment.owner_id is None
    )


def visible_experiments_for(user):
    """Return only owned work and public examples, plus local/admin legacy rows."""

    shared = Q(owner__isnull=True) & shared_example_filter()
    if user.is_authenticated:
        visible = Q(owner_id=user.pk) | shared
        if settings.APOLLINEX_ALLOW_GUEST_UNOWNED_EXPERIMENTS or user.is_superuser:
            # Preserve the local single-user workflow and let admins inspect
            # unassigned legacy rows without exposing them to regular accounts.
            visible |= Q(owner__isnull=True)
    elif settings.APOLLINEX_ALLOW_GUEST_UNOWNED_EXPERIMENTS:
        # Local development historically allowed guests to work with unowned
        # data. Owned records remain hidden from guests.
        visible = Q(owner__isnull=True)
    else:
        visible = shared
    return Experiment.objects.filter(visible).distinct()


def editable_experiments_for(user):
    """Return mutable experiments without granting access to another owner."""

    if user.is_authenticated:
        return Experiment.objects.filter(owner_id=user.pk)
    if settings.APOLLINEX_ALLOW_GUEST_UNOWNED_EXPERIMENTS:
        return Experiment.objects.filter(owner__isnull=True)
    return Experiment.objects.none()
