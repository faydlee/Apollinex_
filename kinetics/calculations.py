"""Расчёты высвобождения, воспроизводящие формулы исходного C# приложения.

Raw values and per-replicate derived values are persisted separately. This
module contains the only implementation of the raw-signal release formula.
"""

from dataclasses import dataclass
import math
import struct
from typing import Iterable


CALCULATION_VERSION = "source-float32-1"

SOURCE_MODEL_DEFINITIONS = (
    {
        "key": "zero_order",
        "name": "Zero Order",
        "x_formula": "t",
        "y_formula": "Q",
        "rule": "все точки",
    },
    {
        "key": "first_order",
        "name": "First Order",
        "x_formula": "t",
        "y_formula": "ln(100 − Q)",
        "rule": "все точки; Q < 100 для конечного логарифма",
    },
    {
        "key": "higuchi",
        "name": "Higuchi",
        "x_formula": "√t",
        "y_formula": "Q",
        "rule": "все точки; t ≥ 0 для вещественного корня",
    },
    {
        "key": "hixson_crowell",
        "name": "Hixson–Crowell",
        "x_formula": "t",
        "y_formula": "∛100 − ∛(100 − Q)",
        "rule": "C# пропускает t = 0; Q ≤ 100",
    },
    {
        "key": "korsmeyer_peppas",
        "name": "Korsmeyer–Peppas",
        "x_formula": "ln(t)",
        "y_formula": "ln(Q)",
        "rule": "C# пропускает t = 0; t > 0 и Q > 0",
    },
)


class CalculationInputError(ValueError):
    """An input cannot be used by the source application's calculation."""


@dataclass(frozen=True)
class CalculationParameters:
    drug_loading: float
    entrapment_efficiency: float
    reactor_volume_ml: float
    probe_volume_ml: float
    calibration_coefficient: float


@dataclass(frozen=True)
class RawPoint:
    time: float
    signal_a: float
    source_c_abs: float | None = None
    source_m_real: float | None = None
    source_release: float | None = None


@dataclass(frozen=True)
class ReplicateInput:
    number: int
    sample_mass_g: float
    points: tuple[RawPoint, ...]


@dataclass(frozen=True)
class CalculatedPoint:
    time: float
    c_abs: float
    m_real: float
    release: float


@dataclass(frozen=True)
class AggregatePoint:
    time: float
    release: float
    sd: float


@dataclass(frozen=True)
class CalculationResult:
    replicate_points: tuple[tuple[CalculatedPoint, ...], ...]
    aggregate_points: tuple[AggregatePoint, ...]


def _finite(value, label):
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise CalculationInputError(f"{label}: требуется число.") from exc
    if not math.isfinite(number):
        raise CalculationInputError(f"{label}: требуется конечное число.")
    return number


def _float32(value):
    """Round one operation result as C#'s source-side float fields do."""

    try:
        return struct.unpack("!f", struct.pack("!f", float(value)))[0]
    except OverflowError:
        return math.copysign(math.inf, float(value))


def _validated_parameters(parameters):
    if not isinstance(parameters, CalculationParameters):
        raise CalculationInputError("Не заданы параметры расчёта.")

    drug_loading = _float32(_finite(parameters.drug_loading, "DL"))
    entrapment_efficiency = _float32(
        _finite(parameters.entrapment_efficiency, "EE")
    )
    reactor_volume = _float32(
        _finite(parameters.reactor_volume_ml, "Объём реактора")
    )
    probe_volume = _float32(
        _finite(parameters.probe_volume_ml, "Объём пробы")
    )
    calibration = _float32(
        _finite(parameters.calibration_coefficient, "Коэффициент калибровки")
    )
    if not all(math.isfinite(value) for value in (
        drug_loading,
        entrapment_efficiency,
        reactor_volume,
        probe_volume,
        calibration,
    )):
        raise CalculationInputError("Параметр выходит за диапазон чисел C# типа float.")

    if drug_loading < 0:
        raise CalculationInputError("DL не может быть отрицательным.")
    if entrapment_efficiency < 0:
        raise CalculationInputError("EE не может быть отрицательным.")
    if reactor_volume <= 0:
        raise CalculationInputError("Объём реактора должен быть больше нуля.")
    if probe_volume < 0:
        raise CalculationInputError("Объём пробы не может быть отрицательным.")
    if calibration < 0:
        raise CalculationInputError(
            "Коэффициент калибровки не может быть отрицательным."
        )

    return (
        drug_loading,
        entrapment_efficiency,
        reactor_volume,
        probe_volume,
        calibration,
    )


def calculate_replicate(replicate, parameters):
    """Calculate each raw point in its input order, using the C# equations.

    C# treats EE as a direct multiplier, although its UI labels EE with a
    percent sign. The input is therefore used as supplied, without silently
    dividing by 100.
    """

    if not isinstance(replicate, ReplicateInput):
        raise CalculationInputError("Не заданы исходные данные повтора.")
    if replicate.number < 1:
        raise CalculationInputError("Номер повтора должен начинаться с 1.")

    sample_mass = _float32(_finite(replicate.sample_mass_g, "Масса образца"))
    if not math.isfinite(sample_mass):
        raise CalculationInputError(
            "Масса образца выходит за диапазон чисел C# типа float."
        )
    if sample_mass <= 0:
        raise CalculationInputError("Масса образца должна быть больше нуля.")

    (
        drug_loading,
        entrapment_efficiency,
        reactor_volume,
        probe_volume,
        calibration,
    ) = _validated_parameters(parameters)

    if not replicate.points:
        raise CalculationInputError(
            f"В повторе {replicate.number} нет исходных измерений."
        )

    denominator = _float32(1 + drug_loading)
    drug_mass = _float32(sample_mass / denominator)
    drug_mass = _float32(drug_mass * drug_loading)
    drug_mass = _float32(drug_mass * 1000)
    drug_mass = _float32(drug_mass * entrapment_efficiency)
    if not math.isfinite(drug_mass):
        raise CalculationInputError("Расчётная масса препарата не является конечной.")

    calculated = []
    previous_c_abs_sum = 0.0
    for index, point in enumerate(replicate.points, start=1):
        time = _float32(_finite(
            point.time,
            f"Время, повтор {replicate.number}, точка {index}",
        ))
        signal = _float32(_finite(
            point.signal_a,
            f"Сигнал A, повтор {replicate.number}, точка {index}",
        ))
        if not math.isfinite(time) or not math.isfinite(signal):
            raise CalculationInputError(
                f"Исходное значение в повторе {replicate.number}, точке {index} "
                "выходит за диапазон чисел C# типа float."
            )
        if time < 0:
            raise CalculationInputError("Время не может быть отрицательным.")

        c_abs = _float32(signal * calibration)
        reactor_mass = _float32(c_abs * reactor_volume)
        probe_mass = _float32(previous_c_abs_sum * probe_volume)
        m_real = _float32(reactor_mass + probe_mass)
        if m_real == 0 and drug_mass == 0:
            release = 0.0
        elif drug_mass == 0:
            raise CalculationInputError(
                "Невозможно рассчитать Release: масса препарата равна нулю, "
                "а Mreal не равна нулю. Проверьте DL и EE."
            )
        else:
            release = _float32(m_real / drug_mass)
            release = _float32(release * 100)

        if not all(math.isfinite(value) for value in (c_abs, m_real, release)):
            raise CalculationInputError(
                f"Расчёт дал неконечное значение в повторе {replicate.number}, "
                f"точке {index}."
            )

        calculated.append(CalculatedPoint(time, c_abs, m_real, release))
        previous_c_abs_sum = _float32(previous_c_abs_sum + c_abs)

    return tuple(calculated)


def calculate_dataset(replicates: Iterable[ReplicateInput], parameters):
    """Calculate replicates and their mean/population SD by point index.

    The C# app uses the first replicate's time at each index and combines
    releases at that index. It assumes matching replicate lengths but does not
    reorder rows by time or compare timestamps, so this preserves that rule.
    """

    ordered_replicates = sorted(replicates, key=lambda item: item.number)
    if not ordered_replicates:
        raise CalculationInputError("Добавьте хотя бы один повтор опыта.")
    numbers = [replicate.number for replicate in ordered_replicates]
    if len(numbers) != len(set(numbers)):
        raise CalculationInputError("Номера повторов должны быть уникальными.")

    calculated_by_replicate = tuple(
        calculate_replicate(replicate, parameters)
        for replicate in ordered_replicates
    )
    point_count = len(calculated_by_replicate[0])
    for replicate, points in zip(ordered_replicates[1:], calculated_by_replicate[1:]):
        if len(points) != point_count:
            raise CalculationInputError(
                "Во всех повторах должно быть одинаковое количество точек; "
                f"в повторе {replicate.number} найдено {len(points)}, "
                f"ожидалось {point_count}."
            )

    aggregate = []
    for point_index in range(point_count):
        values = [points[point_index].release for points in calculated_by_replicate]
        # Enumerable.Average over float returns a float; the following SD
        # calculation in C# promotes each value and the rounded mean to double.
        mean_release = _float32(math.fsum(values) / len(values))
        population_sd = math.sqrt(
            sum((value - mean_release) ** 2 for value in values) / len(values)
        )
        aggregate.append(AggregatePoint(
            time=calculated_by_replicate[0][point_index].time,
            release=mean_release,
            sd=population_sd,
        ))

    return CalculationResult(
        replicate_points=calculated_by_replicate,
        aggregate_points=tuple(aggregate),
    )


def _transform_source_point(model_key, time, release):
    if model_key == "zero_order":
        return time, release
    if model_key == "first_order":
        remaining = 100 - release
        if remaining <= 0:
            return None
        return time, math.log(remaining)
    if model_key == "higuchi":
        if time < 0:
            return None
        return math.sqrt(time), release
    if model_key == "hixson_crowell":
        if time == 0 or release > 100:
            return None
        return time, math.pow(100, 1 / 3) - math.pow(100 - release, 1 / 3)
    if model_key == "korsmeyer_peppas":
        if time == 0 or time < 0 or release <= 0:
            return None
        return math.log(time), math.log(release)
    raise ValueError("Неизвестная модель преобразования.")


def transform_source_model(points, model_key, *, include_records=True):
    """Return source-C# transformed coordinates without fitting a regression.

    The original C# project constructs these coordinates for chart series but
    does not define regression coefficients or R²/RMSE for them. Invalid
    domains and source-specific t=0 omissions are retained as explicit rows.
    """

    definition = next(
        (item for item in SOURCE_MODEL_DEFINITIONS if item["key"] == model_key),
        None,
    )
    if definition is None:
        raise ValueError("Неизвестная модель преобразования.")

    records = []
    included_count = 0
    source_omission_count = 0
    invalid_domain_count = 0
    input_count = 0
    for index, point in enumerate(points, start=1):
        input_count += 1
        try:
            time = float(point[0])
            release = float(point[1])
        except (TypeError, ValueError, OverflowError, IndexError):
            time = release = math.nan

        if not math.isfinite(time) or not math.isfinite(release):
            coordinates = None
            status = "invalid_input"
            invalid_domain_count += 1
        elif model_key in {"hixson_crowell", "korsmeyer_peppas"} and time == 0:
            coordinates = None
            status = "skipped_t0_source_rule"
            source_omission_count += 1
        else:
            try:
                coordinates = _transform_source_point(model_key, time, release)
            except (ValueError, OverflowError, ZeroDivisionError):
                coordinates = None
            if coordinates is None or not all(math.isfinite(value) for value in coordinates):
                coordinates = None
                status = "invalid_domain"
                invalid_domain_count += 1
            else:
                status = "included"
                included_count += 1

        if include_records:
            records.append({
                "order": index,
                "time": time,
                "release": release,
                "x": coordinates[0] if coordinates is not None else None,
                "y": coordinates[1] if coordinates is not None else None,
                "status": status,
            })

    return {
        **definition,
        "records": tuple(records),
        "input_count": input_count,
        "included_count": included_count,
        "source_omission_count": source_omission_count,
        "invalid_domain_count": invalid_domain_count,
    }


def summarize_source_model_transforms(points):
    """Report usable and omitted point counts for each source transform."""

    return tuple(
        {
            key: value
            for key, value in transform_source_model(
                points,
                definition["key"],
                include_records=False,
            ).items()
            if key != "records"
        }
        for definition in SOURCE_MODEL_DEFINITIONS
    )


def recalculate_experiment(experiment):
    """Recompute raw-derived outputs atomically without altering raw values."""

    from django.db import transaction

    from .models import Measurement, RawReplicate, ReplicateCalculation

    if experiment.input_mode != experiment.InputMode.RAW:
        raise CalculationInputError(
            "Пересчёт доступен только для экспериментов с исходными измерениями A."
        )

    parameters = CalculationParameters(
        drug_loading=experiment.drug_loading,
        entrapment_efficiency=experiment.entrapment_efficiency,
        reactor_volume_ml=experiment.reactor_volume_ml,
        probe_volume_ml=experiment.probe_volume_ml,
        calibration_coefficient=experiment.calibration_coefficient,
    )
    replicates = []
    for replicate in RawReplicate.objects.filter(experiment=experiment).prefetch_related(
        "measurements"
    ):
        raw_points = tuple(
            RawPoint(row.time, row.signal_a)
            for row in replicate.measurements.all().order_by("order", "pk")
        )
        replicates.append(ReplicateInput(
            number=replicate.replicate_number,
            sample_mass_g=replicate.sample_mass_g,
            points=raw_points,
        ))

    result = calculate_dataset(replicates, parameters)
    raw_rows = list(
        RawReplicate.objects.filter(experiment=experiment)
        .order_by("replicate_number", "pk")
        .prefetch_related("measurements")
    )
    calculation_rows = []
    for replicate, calculated_points in zip(raw_rows, result.replicate_points):
        source_rows = list(replicate.measurements.all().order_by("order", "pk"))
        calculation_rows.extend(
            ReplicateCalculation(
                measurement=source,
                c_abs=calculated.c_abs,
                m_real=calculated.m_real,
                release=calculated.release,
                algorithm_version=CALCULATION_VERSION,
            )
            for source, calculated in zip(source_rows, calculated_points)
        )

    with transaction.atomic():
        ReplicateCalculation.objects.filter(
            measurement__replicate__experiment=experiment
        ).delete()
        Measurement.objects.filter(experiment=experiment).delete()
        ReplicateCalculation.objects.bulk_create(calculation_rows)
        Measurement.objects.bulk_create([
            Measurement(
                experiment=experiment,
                time=point.time,
                release=point.release,
                sd=point.sd,
                order=index,
            )
            for index, point in enumerate(result.aggregate_points)
        ])
        experiment.calculation_version = CALCULATION_VERSION
        experiment.save(update_fields=("calculation_version", "updated_at"))

    return result
