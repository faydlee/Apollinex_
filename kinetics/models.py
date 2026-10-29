"""Модели базы данных для экспериментов и отдельных точек измерения."""

from django.conf import settings
from django.db import models

# Основные сущности эксперимента
class Experiment(models.Model):
    """Хранит сведения об одном эксперименте и исходном файле с данными."""

    class SourceFormat(models.TextChoices):
        LEGACY_UNKNOWN = "legacy-unknown", "Не определён (старая запись)"
        TIME_RELEASE_SD = "time-release-sd", "Таблица Time/Release/SD"
        MANUAL_RAW = "manual-raw", "Ручной ввод исходных измерений"
        SYNTHETIC_RAW = "synthetic-raw", "Синтетический учебный набор"
        CSHARP_RAW = "csharp-raw", "Исходная книга C#"
        CSHARP_ALLDATA = "csharp-alldata", "Книга AllData C#"

    class InputMode(models.TextChoices):
        PRECALCULATED = "precalculated", "Готовые Time/Release/SD"
        RAW = "raw", "Исходные измерения A"

    name = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    source_file = models.CharField(max_length=255)
    owner = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="experiments",
    )
    input_mode = models.CharField(
        max_length=16,
        choices=InputMode.choices,
        default=InputMode.PRECALCULATED,
    )
    source_format = models.CharField(
        max_length=24,
        choices=SourceFormat.choices,
        default=SourceFormat.LEGACY_UNKNOWN,
    )
    drug_loading = models.FloatField(null=True, blank=True)
    entrapment_efficiency = models.FloatField(null=True, blank=True)
    reactor_volume_ml = models.FloatField(null=True, blank=True)
    probe_volume_ml = models.FloatField(null=True, blank=True)
    calibration_coefficient = models.FloatField(null=True, blank=True)
    calculation_version = models.CharField(max_length=32, blank=True, default="")
    source_results_stale = models.BooleanField(default=False, editable=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    upload_token = models.UUIDField(
        null=True,
        blank=True,
        unique=True,
        editable=False,
    )

    def __str__(self):
        return self.name


class Measurement(models.Model):
    """Одна точка измерения, связанная с экспериментом отношением многие-к-одному."""

    # При удалении эксперимента Django удалит и его измерения; related_name
    # позволяет получать их через experiment.measurements.
    experiment = models.ForeignKey(
        Experiment,
        on_delete=models.CASCADE,
        related_name="measurements",
    )
    time = models.FloatField()
    release = models.FloatField()
    order = models.PositiveIntegerField(default=0)
    sd = models.FloatField(null=True, blank=True, default=None)


class RawReplicate(models.Model):
    """Исходный повтор опыта и масса образца, введённая для этого повтора."""

    experiment = models.ForeignKey(
        Experiment,
        on_delete=models.CASCADE,
        related_name="raw_replicates",
    )
    replicate_number = models.PositiveIntegerField()
    sample_mass_g = models.FloatField()

    class Meta:
        ordering = ("replicate_number", "pk")
        constraints = [
            models.UniqueConstraint(
                fields=("experiment", "replicate_number"),
                name="unique_raw_replicate_per_experiment",
            ),
        ]

    def __str__(self):
        return f"{self.experiment}: повтор {self.replicate_number}"


class RawMeasurement(models.Model):
    """Неизменённые время и сигнал A из одного повтора."""

    replicate = models.ForeignKey(
        RawReplicate,
        on_delete=models.CASCADE,
        related_name="measurements",
    )
    time = models.FloatField()
    signal_a = models.FloatField()
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ("replicate__replicate_number", "order", "pk")
        constraints = [
            models.UniqueConstraint(
                fields=("replicate", "order"),
                name="unique_raw_measurement_order",
            ),
        ]


class ReplicateCalculation(models.Model):
    """Рассчитанные значения Cabs, Mreal и Release для исходной точки."""

    measurement = models.OneToOneField(
        RawMeasurement,
        on_delete=models.CASCADE,
        related_name="calculation",
    )
    c_abs = models.FloatField()
    m_real = models.FloatField()
    release = models.FloatField()
    algorithm_version = models.CharField(max_length=32)


class ImportedReplicateCalculation(models.Model):
    """Per-row derived values preserved from a source AllData workbook."""

    measurement = models.OneToOneField(
        RawMeasurement,
        on_delete=models.CASCADE,
        related_name="source_calculation",
    )
    c_abs = models.FloatField(null=True, blank=True)
    m_real = models.FloatField(null=True, blank=True)
    release = models.FloatField(null=True, blank=True)
    source_format = models.CharField(max_length=24, default="csharp-alldata")


class ImportedModelPoint(models.Model):
    """Transformed coordinates preserved from a C# AllData workbook."""

    experiment = models.ForeignKey(
        Experiment,
        on_delete=models.CASCADE,
        related_name="imported_model_points",
    )
    model_key = models.CharField(max_length=32)
    order = models.PositiveIntegerField()
    x = models.FloatField(null=True, blank=True)
    y = models.FloatField(null=True, blank=True)
    sd = models.FloatField(null=True, blank=True)
    status = models.CharField(max_length=32, default="included")

    class Meta:
        ordering = ("model_key", "order", "pk")
        constraints = [
            models.UniqueConstraint(
                fields=("experiment", "model_key", "order"),
                name="unique_imported_model_point_order",
            ),
        ]
