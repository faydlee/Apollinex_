from django.contrib import admin

from .access import visible_experiments_for
from .models import Experiment, Measurement


# Административная панель
class MeasurementInline(admin.TabularInline):
    """Показывает измерения прямо на странице связанного эксперимента."""

    model = Measurement
    extra = 0


@admin.register(Experiment)
class ExperimentAdmin(admin.ModelAdmin):
    """Настраивает список экспериментов и встроенное редактирование измерений."""

    list_display = (
        "experiment_id",
        "name",
        "owner",
        "source_format",
        "source_file",
        "created_at",
        "updated_at",
    )
    list_filter = ("owner", "source_format", "input_mode")
    search_fields = ("name", "source_file")
    # Owner changes must go through the guarded management command.
    readonly_fields = ("owner",)
    inlines = [MeasurementInline]

    @admin.display(description="ID", ordering="pk")
    def experiment_id(self, obj):
        return obj.pk

    def get_queryset(self, request):
        return visible_experiments_for(request.user)

    def save_model(self, request, obj, form, change):
        if not change and obj.owner_id is None:
            obj.owner = request.user
        super().save_model(request, obj, form, change)


@admin.register(Measurement)
class MeasurementAdmin(admin.ModelAdmin):
    """Добавляет отдельный список измерений с фильтром по эксперименту."""

    list_display = ("experiment", "time", "release", "sd", "order")
    list_filter = ("experiment",)

    def get_queryset(self, request):
        queryset = super().get_queryset(request)
        return queryset.filter(
            experiment__in=visible_experiments_for(request.user)
        )
