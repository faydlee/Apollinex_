"""Страницы загрузки, истории и просмотра экспериментов."""

import base64
import csv
import hashlib
import io
import json
import logging
import math
import re
import uuid
from pathlib import PurePosixPath
from urllib.parse import urlencode
from zipfile import BadZipFile

import matplotlib
import numpy as np

matplotlib.use("Agg")

import matplotlib.pyplot as plt
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.cache import cache
from django.core.paginator import Paginator
from django.db import DatabaseError, IntegrityError, transaction
from django.db.models import Count, Max, Min
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils.formats import get_format
from django.views.decorators.http import require_POST, require_http_methods
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as SpreadsheetImage
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from .calculations import (
    CALCULATION_VERSION,
    SOURCE_MODEL_DEFINITIONS,
    CalculationInputError,
    fit_linear_trendline,
    recalculate_experiment,
    summarize_source_model_transforms,
    transform_source_model,
)
from .access import (
    can_edit_experiment,
    editable_experiments_for,
    visible_experiments_for,
)
from .excel_import import (
    RawWorkbookValidationError,
    read_csharp_all_data_workbook,
    read_csharp_raw_workbook,
)
from .examples import SYNTHETIC_EXAMPLES
from .forms import RawExperimentForm
from .models import (
    Experiment,
    ImportedReplicateCalculation,
    ImportedModelPoint,
    Measurement,
    RawMeasurement,
    RawReplicate,
)


logger = logging.getLogger(__name__)

CURRENT_EXPERIMENT_SESSION_KEY = "kinetics_current_experiment_id"
MAX_EXCEL_UPLOAD_SIZE = 25 * 1024 * 1024
MAX_EXCEL_ROWS = 100_000
MAX_EXCEL_COLUMNS = 64
MAX_PLOT_POINTS = 5_000
TABLE_PAGE_SIZE = 100
HISTORY_PAGE_SIZE = 25
PLOT_CACHE_SECONDS = 60 * 60

TIME_HEADERS = {
    "time", "время", "timeh", "timehours", "timehr", "timehrs",
    "timemin", "timeminute", "timeminutes", "timesec", "timesecond",
    "timeseconds", "времячас", "времячасы", "времямин", "времяминута",
    "времяминуты", "времясек", "времясекунда", "времясекунды",
}
RELEASE_HEADERS = {
    "release",
    "releasepercent",
    "drugrelease",
    "drugreleasepercent",
    "высвобождение",
    "высвобождениеpercent",
}
SD_HEADERS = {
    "stardartd", "standardd", "standartd", "standartdeviation",
    "standarddeviation", "standarddeviationvalue", "sd",
}


class WorkbookValidationError(ValueError):
    """Ошибка входных данных книги, которую можно показать пользователю."""


def _normalize_header(value):
    if value is None:
        return ""
    return re.sub(r"[\W_]+", "", str(value).strip().casefold())


def _parse_measurement(value, row_number, label):
    if value is None or isinstance(value, bool):
        raise WorkbookValidationError(
            f"Пустое или некорректное значение «{label}» в строке {row_number}."
        )

    if isinstance(value, str):
        value = value.strip()
        if not value:
            raise WorkbookValidationError(
                f"Пустое значение «{label}» в строке {row_number}."
            )

    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise WorkbookValidationError(
            f"Значение «{label}» в строке {row_number} должно быть числом."
        ) from exc

    if not math.isfinite(number):
        raise WorkbookValidationError(
            f"Значение «{label}» в строке {row_number} должно быть конечным числом."
        )
    return number


def _parse_optional_sd(value, row_number, *, label="StardartD (SD)"):
    """Читает необязательное стандартное отклонение, не подменяя его SEM."""

    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    deviation = _parse_measurement(value, row_number, label)
    if deviation < 0:
        raise WorkbookValidationError(
            f"Значение «{label}» в строке {row_number} не может быть отрицательным."
        )
    return deviation


def _read_measurements(excel_file):
    """Потоково проверяет .xlsx и возвращает (время, высвобождение, SD)."""

    workbook = None
    sheet = None
    try:
        excel_file.seek(0)
        workbook = load_workbook(excel_file, read_only=True, data_only=True)
        sheet = workbook.active
        if sheet is None or not hasattr(sheet, "iter_rows"):
            raise WorkbookValidationError(
                "В книге нет листа с табличными измерениями."
            )

        header = next(
            sheet.iter_rows(
                min_row=1,
                max_row=1,
                max_col=MAX_EXCEL_COLUMNS,
                values_only=True,
            ),
            None,
        )
        if not header or not any(
            value is not None and str(value).strip()
            for value in header
        ):
            raise WorkbookValidationError("Лист Excel пуст.")

        column_indexes = {}
        for index, value in enumerate(header):
            normalized = _normalize_header(value)
            if normalized in TIME_HEADERS:
                if "time" in column_indexes:
                    raise WorkbookValidationError(
                        f"В строке 1 найден повторный заголовок времени в столбце "
                        f"{get_column_letter(index + 1)}."
                    )
                column_indexes["time"] = index
            elif normalized in RELEASE_HEADERS:
                if "release" in column_indexes:
                    raise WorkbookValidationError(
                        f"В строке 1 найден повторный заголовок Release в столбце "
                        f"{get_column_letter(index + 1)}."
                    )
                column_indexes["release"] = index
            elif normalized in SD_HEADERS:
                if "sd" in column_indexes:
                    raise WorkbookValidationError(
                        f"В строке 1 найден повторный заголовок SD в столбце "
                        f"{get_column_letter(index + 1)}."
                    )
                column_indexes["sd"] = index

        if "time" not in column_indexes or "release" not in column_indexes:
            raise WorkbookValidationError(
                "В строке 1 нужны заголовки Time (Время) и Release (Высвобождение)."
            )

        last_required_column = max(column_indexes.values()) + 1
        points = []
        for row_number, row in enumerate(
            sheet.iter_rows(
                min_row=2,
                max_col=last_required_column,
                values_only=True,
            ),
            start=2,
        ):
            if row_number > MAX_EXCEL_ROWS + 1:
                raise WorkbookValidationError(
                    f"В файле допускается не более {MAX_EXCEL_ROWS:,} строк данных."
                )

            time_index = column_indexes["time"]
            release_index = column_indexes["release"]
            sd_index = column_indexes.get("sd")
            time_value = row[time_index] if time_index < len(row) else None
            release_value = row[release_index] if release_index < len(row) else None

            if all(value is None or (isinstance(value, str) and not value.strip()) for value in row):
                continue

            time_column = get_column_letter(time_index + 1)
            release_column = get_column_letter(release_index + 1)
            time = _parse_measurement(
                time_value, row_number, f"Время, столбец {time_column}"
            )
            release = _parse_measurement(
                release_value, row_number, f"Высвобождение, столбец {release_column}"
            )
            sd_value = row[sd_index] if sd_index is not None and sd_index < len(row) else None
            sd_label = (
                f"StardartD (SD), столбец {get_column_letter(sd_index + 1)}"
                if sd_index is not None else "StardartD (SD)"
            )
            sd = _parse_optional_sd(sd_value, row_number, label=sd_label)
            points.append((time, release, sd))

        if not points:
            raise WorkbookValidationError(
                "В таблице нет строк с измерениями после строки заголовков."
            )
        return points
    except WorkbookValidationError as exc:
        if sheet is not None and f"Лист «{sheet.title}»" not in str(exc):
            raise WorkbookValidationError(
                f"Лист «{sheet.title}»: {exc}"
            ) from exc
        raise
    except (
        BadZipFile,
        InvalidFileException,
        OSError,
        EOFError,
        KeyError,
        ValueError,
        TypeError,
        IndexError,
        AttributeError,
        OverflowError,
        SyntaxError,
    ) as exc:
        raise WorkbookValidationError(
            "Файл повреждён или не является корректной книгой .xlsx."
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except OSError:
                logger.warning(
                    "Не удалось закрыть книгу Excel после чтения",
                    exc_info=True,
                )


def _has_precomputed_headers(excel_file):
    """Recognize the supported one-table layout before trying C# layouts."""

    workbook = None
    try:
        excel_file.seek(0)
        workbook = load_workbook(excel_file, read_only=True, data_only=True)
        sheet = workbook.active
        if sheet is None or not hasattr(sheet, "iter_rows"):
            raise WorkbookValidationError(
                "В книге нет листа с табличными измерениями."
            )
        header = next(
            sheet.iter_rows(
                min_row=1,
                max_row=1,
                max_col=MAX_EXCEL_COLUMNS,
                values_only=True,
            ),
            (),
        )
        populated_header = [
            value for value in header
            if value is not None and str(value).strip()
        ]
        if not populated_header:
            # C# workbooks keep their measurement blocks below the first row;
            # an actually empty worksheet should get a useful table error.
            if (getattr(sheet, "max_row", None) or 1) <= 1:
                raise WorkbookValidationError("Лист Excel пуст.")
            return False
        normalized = {_normalize_header(value) for value in header}
        if normalized.intersection(TIME_HEADERS | RELEASE_HEADERS):
            return True
        # A populated, table-like header row is a Time/Release candidate even
        # when both names are wrong; let the regular reader report the missing
        # columns instead of burying that actionable error in C# diagnostics.
        return len(populated_header) >= 2 and all(
            isinstance(value, str) for value in populated_header
        )
    except WorkbookValidationError:
        raise
    except (BadZipFile, InvalidFileException, OSError, EOFError, ValueError) as exc:
        raise WorkbookValidationError(
            "Файл повреждён или не является корректной книгой .xlsx."
        ) from exc
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except OSError:
                logger.warning("Не удалось закрыть книгу Excel после определения формата", exc_info=True)


def _has_all_data_marker(excel_file):
    workbook = None
    try:
        excel_file.seek(0)
        workbook = load_workbook(excel_file, read_only=True, data_only=True)
        sheet = workbook.active
        return _normalize_header(sheet.cell(row=2, column=9).value) == "zeroorder"
    except (BadZipFile, InvalidFileException, OSError, EOFError, ValueError):
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except OSError:
                logger.warning("Не удалось закрыть книгу AllData", exc_info=True)


def _safe_source_filename(filename):
    filename = PurePosixPath(str(filename).replace("\\", "/")).name
    filename = "".join(
        character
        for character in filename
        if character >= " " and character != "\x7f"
    )
    return filename[:255] or "experiment.xlsx"


def _get_upload_token():
    """Создаёт отдельный ключ для каждой формы, чтобы параллельные вкладки не конфликтовали."""

    return str(uuid.uuid4())


def _selected_experiment(request):
    visible_experiments = visible_experiments_for(request.user)
    query_experiment_id = request.GET.get("experiment_id")
    if query_experiment_id:
        try:
            requested_id = int(query_experiment_id)
            if 0 < requested_id <= 9_223_372_036_854_775_807:
                requested_experiment = visible_experiments.filter(pk=requested_id).first()
            else:
                requested_experiment = None
        except (TypeError, ValueError, OverflowError):
            requested_experiment = None
        if requested_experiment is not None:
            request.session[CURRENT_EXPERIMENT_SESSION_KEY] = requested_experiment.pk
            return requested_experiment

    experiment_id = request.session.get(CURRENT_EXPERIMENT_SESSION_KEY)
    if experiment_id is not None:
        try:
            experiment = visible_experiments.filter(pk=experiment_id).first()
        except (TypeError, ValueError):
            experiment = None
        if experiment is not None:
            return experiment
        request.session.pop(CURRENT_EXPERIMENT_SESSION_KEY, None)
    return visible_experiments.order_by("-created_at", "-pk").first()


def _plot_cache_key(points, dataset_fingerprint):
    payload = json.dumps(
        {"fingerprint": dataset_fingerprint, "points": points},
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")
    return f"kinetics:plot:{hashlib.sha256(payload).hexdigest()}"


def _build_plot(points):
    times = [point[0] for point in points]
    releases = [point[1] for point in points]
    sd_points = [
        (point[0], point[1], point[2])
        for point in points
        if len(point) > 2 and point[2] is not None
    ]
    figure, axis = plt.subplots(figsize=(8.4, 5.2))
    image_buffer = io.BytesIO()
    try:
        axis.set_axisbelow(True)
        axis.grid(True, which="major", color="#dce4ee", linewidth=0.8)
        if sd_points:
            axis.errorbar(
                [point[0] for point in sd_points],
                [point[1] for point in sd_points],
                yerr=[point[2] for point in sd_points],
                fmt="none",
                ecolor="#879bb4",
                elinewidth=1.15,
                capsize=3,
                label="± SD",
                zorder=2,
            )
        axis.plot(
            times,
            releases,
            color="#9eb8d2",
            linewidth=1.1,
            label="_nolegend_",
            zorder=2,
        )
        axis.scatter(
            times,
            releases,
            marker="o",
            s=34,
            color="#1769aa",
            edgecolors="white",
            linewidths=0.7,
            label="Экспериментальные данные",
            zorder=3,
        )
        axis.set_xlabel("Время (единицы исходного файла)")
        axis.set_ylabel("Высвобождение (единицы исходного файла)")
        axis.set_title("Кинетика высвобождения")
        axis.margins(x=0.04, y=0.12)
        axis.legend(frameon=False)
        figure.tight_layout()
        figure.savefig(image_buffer, format="png", bbox_inches="tight", dpi=220)
        image_buffer.seek(0)
        return base64.b64encode(image_buffer.getvalue()).decode("ascii")
    finally:
        plt.close(figure)
        image_buffer.close()


def _get_or_build_plot(points, dataset_fingerprint):
    if not points:
        return None

    key = _plot_cache_key(points, dataset_fingerprint)
    plot = cache.get(key)
    if plot is not None:
        return plot

    try:
        plot = _build_plot(points)
    except Exception:
        logger.exception("Не удалось построить график эксперимента")
        return None

    cache.set(key, plot, timeout=PLOT_CACHE_SECONDS)
    return plot


def _plot_points(experiment, count=None):
    # A kinetic line must follow actual elapsed time. The measurement table
    # keeps its separate source-row ordering.
    measurements = experiment.measurements.order_by("time", "order", "pk")
    if count is None:
        count = measurements.count()
    if not count:
        return [], hashlib.sha256().hexdigest()

    stride = max(1, math.ceil(count / MAX_PLOT_POINTS))
    points = []
    last_point = None
    fingerprint = hashlib.sha256()
    for index, point in enumerate(
        measurements.values_list("time", "release", "sd").iterator(chunk_size=2_000)
    ):
        fingerprint.update(repr(point).encode("ascii", errors="backslashreplace"))
        fingerprint.update(b";")
        if point[0] is None or point[1] is None:
            continue
        try:
            time = float(point[0])
            release = float(point[1])
        except (TypeError, ValueError, OverflowError):
            continue
        if not math.isfinite(time) or not math.isfinite(release):
            continue
        try:
            sd = None if point[2] is None else float(point[2])
        except (TypeError, ValueError, OverflowError):
            sd = None
        if sd is not None and (not math.isfinite(sd) or sd < 0):
            sd = None
        normalized = (time, release, sd)
        last_point = normalized
        if index % stride == 0 and len(points) < MAX_PLOT_POINTS:
            points.append(normalized)

    if last_point is not None and (not points or points[-1] != last_point):
        if len(points) == MAX_PLOT_POINTS:
            points[-1] = last_point
        else:
            points.append(last_point)
    return points, fingerprint.hexdigest()


def _experiment_context(
    experiment,
    request,
    include_table=False,
    include_plot=True,
):
    context = {
        "experiment": experiment,
        "data": False,
        "filename": "",
        "row_count": 0,
        "plot": None,
        "plot_error": False,
        "plot_data_error": False,
        "measurements_page": None,
    }
    if experiment is None:
        return context

    queryset = experiment.measurements.order_by("order", "pk")
    summary = queryset.aggregate(
        time_min=Min("time"),
        time_max=Max("time"),
        max_release=Max("release"),
    )
    row_count = queryset.count()
    context.update({
        "data": row_count > 0,
        "filename": experiment.source_file,
        "row_count": row_count,
        "time_min": summary["time_min"],
        "time_max": summary["time_max"],
        "max_release": summary["max_release"],
    })
    if include_plot:
        plot_points, fingerprint = _plot_points(experiment, row_count)
        plot = _get_or_build_plot(plot_points, fingerprint)
        context.update({
            "plot": plot,
            "plot_error": bool(plot_points) and plot is None,
            "plot_data_error": row_count > 0 and not plot_points,
        })

    if include_table:
        context["measurements_page"] = Paginator(
            queryset,
            TABLE_PAGE_SIZE,
        ).get_page(request.GET.get("page"))
    return context


def _finite_points(point_rows):
    """Возвращает только конечные пары времени и измеренного высвобождения."""

    points = []
    for time_value, release_value in point_rows:
        try:
            point = (float(time_value), float(release_value))
        except (TypeError, ValueError, OverflowError):
            continue
        if all(math.isfinite(value) for value in point):
            points.append(point)
    return points


def _finite_plot_points(point_rows):
    """Проверяет пары для графика и сохраняет отсутствующее SD как None."""

    points = []
    for time_value, release_value, sd_value in point_rows:
        try:
            time = float(time_value)
            release = float(release_value)
        except (TypeError, ValueError, OverflowError):
            continue
        if not math.isfinite(time) or not math.isfinite(release):
            continue
        try:
            sd = None if sd_value is None else float(sd_value)
        except (TypeError, ValueError, OverflowError):
            sd = None
        if sd is not None and (not math.isfinite(sd) or sd < 0):
            sd = None
        points.append((time, release, sd))
    return points


def _format_model_number(value):
    formatted = format(value, ".8g")
    decimal_separator = get_format("DECIMAL_SEPARATOR")
    if decimal_separator != ".":
        formatted = formatted.replace(".", decimal_separator, 1)
    return formatted


def _first_order_model_fit(points):
    """Подгоняет эмпирическую модель выхода к плато R(t)=R∞(1-exp(-kt))."""

    if len(points) < 3:
        return None, "Для оценки параметров нужны как минимум три измерения."
    if any(
        not math.isfinite(time) or not math.isfinite(release)
        for time, release in points
    ):
        return None, "В данных есть некорректные числовые значения."
    if any(time < 0 or release < 0 for time, release in points):
        return None, "Эта модель применима только к неотрицательным времени и высвобождению."
    times = [time for time, _ in points]
    releases = [release for _, release in points]
    if len(set(times)) < 2 or max(times) <= 0:
        return None, "Для подгонки нужны как минимум два разных значения времени, одно из них больше нуля."
    if max(releases) <= 0:
        return None, "В данных нет положительного высвобождения для оценки плато."

    time_scale = max(times)
    normalized_times = np.asarray([time / time_scale for time in times], dtype=float)
    release_scale = max(releases)
    normalized_releases = np.asarray(
        [release / release_scale for release in releases],
        dtype=float,
    )
    min_log_rate, max_log_rate = -20.0, 20.0
    grid_size = 321
    step = (max_log_rate - min_log_rate) / (grid_size - 1)

    def evaluate(log_scaled_rate):
        scaled_rate = math.exp(log_scaled_rate)
        fractions = -np.expm1(-scaled_rate * normalized_times)
        denominator = float(np.dot(fractions, fractions))
        if denominator <= 0 or not math.isfinite(denominator):
            return math.inf, 0.0, scaled_rate
        normalized_plateau = max(
            0.0,
            float(np.dot(normalized_releases, fractions)) / denominator,
        )
        residuals = normalized_releases - normalized_plateau * fractions
        residual_sum = float(np.dot(residuals, residuals))
        return residual_sum, normalized_plateau, scaled_rate

    candidates = [
        (evaluate(min_log_rate + index * step)[0], index)
        for index in range(grid_size)
    ]
    _, best_index = min(candidates)
    if best_index in (0, grid_size - 1):
        return None, "Оптимум оказался на границе диапазона поиска; по этим данным параметр скорости нельзя надёжно оценить."

    left = min_log_rate + (best_index - 1) * step
    right = min_log_rate + (best_index + 1) * step
    golden_ratio = (math.sqrt(5.0) - 1.0) / 2.0
    x1 = right - golden_ratio * (right - left)
    x2 = left + golden_ratio * (right - left)
    f1 = evaluate(x1)[0]
    f2 = evaluate(x2)[0]
    for _ in range(64):
        if f1 <= f2:
            right, x2, f2 = x2, x1, f1
            x1 = right - golden_ratio * (right - left)
            f1 = evaluate(x1)[0]
        else:
            left, x1, f1 = x1, x2, f2
            x2 = left + golden_ratio * (right - left)
            f2 = evaluate(x2)[0]

    log_rate = (left + right) / 2.0
    residual_sum, normalized_plateau, scaled_rate = evaluate(log_rate)
    try:
        plateau = normalized_plateau * release_scale
        rate = scaled_rate / time_scale
        rmse = math.sqrt(residual_sum / len(points)) * release_scale
    except OverflowError:
        return None, "Числовой диапазон данных не позволяет надёжно вычислить параметры модели."
    if not all(math.isfinite(value) for value in (plateau, rate, rmse)) or rate <= 0:
        return None, "Числовой диапазон данных не позволяет надёжно вычислить параметры модели."
    mean_release = float(np.mean(normalized_releases))
    centered = normalized_releases - mean_release
    total_sum = float(np.dot(centered, centered))
    fit_quality = None if total_sum == 0 else 1.0 - residual_sum / total_sum
    if fit_quality is not None and not math.isfinite(fit_quality):
        return None, "Числовой диапазон данных не позволяет надёжно оценить качество подгонки."
    return {
        "plateau": plateau,
        "rate": rate,
        "rmse": rmse,
        "r_squared": fit_quality,
        "point_count": len(points),
    }, ""


def _model_metrics(points, predictions):
    """Считает RMSE и R² на исходной шкале, масштабируя сумму квадратов."""

    if len(points) != len(predictions) or not points:
        return None, "Для расчёта качества подгонки недостаточно данных."
    releases = [release for _, release in points]
    if any(not math.isfinite(value) for value in releases + predictions):
        return None, "Модель дала неконечные предсказания."
    scale = max((abs(value) for value in releases), default=0.0) or 1.0
    normalized = [value / scale for value in releases]
    normalized_predictions = [value / scale for value in predictions]
    mean_release = math.fsum(normalized) / len(normalized)
    residual_sum = math.fsum(
        (actual - predicted) ** 2
        for actual, predicted in zip(normalized, normalized_predictions)
    )
    total_sum = math.fsum((actual - mean_release) ** 2 for actual in normalized)
    r_squared = None if total_sum == 0 else 1.0 - residual_sum / total_sum
    try:
        rmse = math.sqrt(residual_sum / len(points)) * scale
    except OverflowError:
        return None, "Числовой диапазон данных не позволяет вычислить RMSE."
    if not math.isfinite(rmse) or (r_squared is not None and not math.isfinite(r_squared)):
        return None, "Числовой диапазон данных не позволяет оценить качество подгонки."
    return {"rmse": rmse, "r_squared": r_squared}, ""


def _validate_fit_points(points, minimum=3):
    if len(points) < minimum:
        return "Для сравнения моделей нужны как минимум три измерения."
    if any(
        not math.isfinite(time) or not math.isfinite(release)
        for time, release in points
    ):
        return "В данных есть некорректные числовые значения."
    if any(time < 0 or release < 0 for time, release in points):
        return "Эта модель применима только к неотрицательным времени и высвобождению."
    if len({time for time, _ in points}) < 2:
        return "Для подгонки нужны как минимум два разных значения времени."
    if max(time for time, _ in points) <= 0:
        return "Для подгонки нужно хотя бы одно измерение при времени больше нуля."
    if max(release for _, release in points) <= 0:
        return "В данных нет положительного высвобождения для подгонки модели."
    return ""


def _fit_through_origin(points, model_key, name, formula, parameter_name, transform):
    error = _validate_fit_points(points)
    if error:
        return None, error
    times = [time for time, _ in points]
    releases = [release for _, release in points]
    release_scale = max(releases)
    transformed = [transform(time) for time in times]
    x_scale = max(transformed)
    if not math.isfinite(x_scale) or x_scale <= 0:
        return None, "Для подгонки нужно больше одного значения времени."
    x_values = [value / x_scale for value in transformed]
    y_values = [value / release_scale for value in releases]
    denominator = math.fsum(value * value for value in x_values)
    if denominator <= 0 or not math.isfinite(denominator):
        return None, "Для подгонки нужно больше одного значения времени."
    scaled_parameter = math.fsum(
        x * y for x, y in zip(x_values, y_values)
    ) / denominator
    try:
        parameter = scaled_parameter * release_scale / x_scale
        predictions = [scaled_parameter * x * release_scale for x in x_values]
    except OverflowError:
        return None, "Числовой диапазон данных не позволяет вычислить параметры модели."
    if not math.isfinite(parameter) or parameter < 0:
        return None, "Подгонка дала некорректный параметр модели."
    metrics, error = _model_metrics(points, predictions)
    if error:
        return None, error
    return {
        "model_key": model_key,
        "name": name,
        "formula": formula,
        "parameters": {parameter_name: parameter},
        "rmse": metrics["rmse"],
        "r_squared": metrics["r_squared"],
        "point_count": len(points),
        "parameter_count": 1,
        "note": "Подгонка с нулевым свободным членом.",
    }, ""


def _fit_hixson_crowell(points):
    """Подгоняет конечную ёмкость по кубическому закону убывания остатка."""

    error = _validate_fit_points(points)
    if error:
        return None, error
    time_scale = max(time for time, _ in points)
    release_scale = max(release for _, release in points)
    times = np.asarray([time / time_scale for time, _ in points], dtype=float)
    releases = np.asarray([release / release_scale for _, release in points], dtype=float)

    def evaluate(log_rate):
        scaled_rate = math.exp(log_rate)
        fractions = 1.0 - np.maximum(0.0, 1.0 - scaled_rate * times) ** 3
        denominator = float(np.dot(fractions, fractions))
        if denominator <= 0 or not math.isfinite(denominator):
            return math.inf, 0.0, scaled_rate
        plateau = max(
            0.0,
            float(np.dot(releases, fractions)) / denominator,
        )
        residuals = releases - plateau * fractions
        residual = float(np.dot(residuals, residuals))
        return residual, plateau, scaled_rate

    minimum, maximum = -20.0, 20.0
    grid_size = 321
    step = (maximum - minimum) / (grid_size - 1)
    candidates = [
        (evaluate(minimum + index * step)[0], index)
        for index in range(grid_size)
    ]
    _, best_index = min(candidates)
    if best_index in (0, grid_size - 1):
        return None, "Оптимум модели Hixson–Crowell оказался на границе диапазона поиска."

    left = minimum + (best_index - 1) * step
    right = minimum + (best_index + 1) * step
    ratio = (math.sqrt(5.0) - 1.0) / 2.0
    x1 = right - ratio * (right - left)
    x2 = left + ratio * (right - left)
    f1 = evaluate(x1)[0]
    f2 = evaluate(x2)[0]
    for _ in range(64):
        if f1 <= f2:
            right, x2, f2 = x2, x1, f1
            x1 = right - ratio * (right - left)
            f1 = evaluate(x1)[0]
        else:
            left, x1, f1 = x1, x2, f2
            x2 = left + ratio * (right - left)
            f2 = evaluate(x2)[0]
    residual, normalized_plateau, scaled_rate = evaluate((left + right) / 2.0)
    try:
        plateau = normalized_plateau * release_scale
        normalized_rate = scaled_rate / time_scale
        plateau_root = plateau ** (1.0 / 3.0)
        rate = normalized_rate * plateau_root
        predictions = [
            plateau - max(0.0, plateau_root - rate * time) ** 3
            for time, _ in points
        ]
    except OverflowError:
        return None, "Числовой диапазон данных не позволяет вычислить параметры модели."
    if not math.isfinite(plateau) or not math.isfinite(rate) or rate <= 0:
        return None, "Числовой диапазон данных не позволяет вычислить параметры модели."
    metrics, error = _model_metrics(points, predictions)
    if error:
        return None, error
    return {
        "model_key": "hixson_crowell",
        "name": "Hixson–Crowell",
        "formula": "R(t) = R∞ − max(0, R∞^(1/3) − kHC·t)³",
        "parameters": {"R∞": plateau, "kHC": rate},
        "rmse": metrics["rmse"],
        "r_squared": metrics["r_squared"],
        "point_count": len(points),
        "parameter_count": 2,
        "note": "Форма предполагает конечную исходную массу и накопленное высвобождение.",
    }, ""


def _fit_korsmeyer_peppas(points):
    error = _validate_fit_points(points)
    if error:
        return None, error
    times = [time for time, _ in points]
    releases = [release for _, release in points]
    positive_times = {time for time in times if time > 0}
    if len(positive_times) < 3:
        return None, "Для степенной подгонки нужны как минимум три разных положительных значения времени."
    time_scale = max(times)
    release_scale = max(releases)
    normalized_times = np.asarray([time / time_scale for time in times], dtype=float)
    normalized_releases = np.asarray(
        [release / release_scale for release in releases],
        dtype=float,
    )

    def evaluate(log_exponent):
        exponent = math.exp(log_exponent)
        fractions = np.power(normalized_times, exponent)
        denominator = float(np.dot(fractions, fractions))
        if denominator <= 0 or not math.isfinite(denominator):
            return math.inf, 0.0, exponent
        coefficient = max(
            0.0,
            float(np.dot(normalized_releases, fractions)) / denominator,
        )
        residuals = normalized_releases - coefficient * fractions
        residual = float(np.dot(residuals, residuals))
        return residual, coefficient, exponent

    min_log_exponent, max_log_exponent = -10.0, 3.0
    grid_size = 321
    step = (max_log_exponent - min_log_exponent) / (grid_size - 1)
    candidates = [
        (evaluate(min_log_exponent + index * step)[0], index)
        for index in range(grid_size)
    ]
    _, best_index = min(candidates)
    if best_index in (0, grid_size - 1):
        return None, "Оптимум модели Korsmeyer–Peppas оказался на границе диапазона поиска."
    left = min_log_exponent + (best_index - 1) * step
    right = min_log_exponent + (best_index + 1) * step
    ratio = (math.sqrt(5.0) - 1.0) / 2.0
    x1 = right - ratio * (right - left)
    x2 = left + ratio * (right - left)
    f1 = evaluate(x1)[0]
    f2 = evaluate(x2)[0]
    for _ in range(64):
        if f1 <= f2:
            right, x2, f2 = x2, x1, f1
            x1 = right - ratio * (right - left)
            f1 = evaluate(x1)[0]
        else:
            left, x1, f1 = x1, x2, f2
            x2 = left + ratio * (right - left)
            f2 = evaluate(x2)[0]
    _, normalized_coefficient, exponent = evaluate((left + right) / 2.0)
    if normalized_coefficient <= 0:
        return None, "Данные не позволяют оценить коэффициент степенной модели."
    try:
        log_coefficient = (
            math.log(normalized_coefficient)
            + math.log(release_scale)
            - exponent * math.log(time_scale)
        )
        coefficient = math.exp(log_coefficient)
        predictions = [
            normalized_coefficient * normalized_time ** exponent * release_scale
            for normalized_time in normalized_times
        ]
    except (OverflowError, ValueError):
        return None, "Числовой диапазон данных не позволяет вычислить степенную модель."
    if not math.isfinite(coefficient) or coefficient <= 0 or not math.isfinite(exponent):
        return None, "Подгонка дала некорректные параметры модели Korsmeyer–Peppas."
    metrics, error = _model_metrics(points, predictions)
    if error:
        return None, error
    return {
        "model_key": "korsmeyer_peppas",
        "name": "Korsmeyer–Peppas",
        "formula": "R(t) = k·tⁿ",
        "parameters": {"k": coefficient, "n": exponent},
        "rmse": metrics["rmse"],
        "r_squared": metrics["r_squared"],
        "point_count": len(points),
        "parameter_count": 2,
        "note": "Нелинейная подгонка по остаткам исходного Release без нормирования на полное количество; стандартная механистическая интерпретация часто требует долю высвобождения и начальную фазу.",
    }, ""


def _fit_kinetic_models(points):
    """Подгоняет поддерживаемые модели и выбирает максимум сопоставимого R²."""

    definitions = [
        (
            "zero_order",
            "Zero Order",
            "R(t) = k₀·t",
            lambda: _fit_through_origin(
                points, "zero_order", "Zero Order", "R(t) = k₀·t", "k₀", lambda time: time,
            ),
        ),
        (
            "first_order",
            "First Order",
            "R(t) = R∞·(1 − e⁻ᵏᵗ)",
            lambda: _first_order_model_fit(points),
        ),
        (
            "higuchi",
            "Higuchi",
            "R(t) = kH·√t",
            lambda: _fit_through_origin(
                points, "higuchi", "Higuchi", "R(t) = kH·√t", "kH", math.sqrt,
            ),
        ),
        (
            "hixson_crowell",
            "Hixson–Crowell",
            "R(t) = R∞ − max(0, R∞^(1/3) − kHC·t)³",
            lambda: _fit_hixson_crowell(points),
        ),
        (
            "korsmeyer_peppas",
            "Korsmeyer–Peppas",
            "R(t) = k·tⁿ",
            lambda: _fit_korsmeyer_peppas(points),
        ),
    ]
    results = []
    for key, name, formula, fit_function in definitions:
        fitted, error = fit_function()
        if fitted is None:
            results.append({
                "model_key": key,
                "name": name,
                "formula": formula,
                "success": False,
                "error": error,
                "parameters_display": [],
                "r_squared_display": None,
                "rmse_display": None,
            })
            continue
        if key == "first_order":
            parameters = {"R∞": fitted["plateau"], "k": fitted["rate"]}
            note = "Экспоненциальная эмпирическая модель выхода к плато."
            parameter_count = 2
        else:
            parameters = fitted["parameters"]
            note = fitted["note"]
            parameter_count = fitted["parameter_count"]
        model_result = {
            "model_key": key,
            "name": name,
            "formula": fitted.get("formula", formula),
            "parameters": parameters,
            "rmse": fitted.get("rmse"),
            "r_squared": fitted.get("r_squared"),
            "point_count": fitted.get("point_count", len(points)),
            "parameter_count": parameter_count,
            "note": note,
            "success": True,
            "error": "",
            "parameters_display": [
                {"name": parameter, "value": _format_model_number(value)}
                for parameter, value in parameters.items()
            ],
            "r_squared_display": (
                _format_model_number(fitted["r_squared"])
                if fitted.get("r_squared") is not None else None
            ),
            "rmse_display": (
                _format_model_number(fitted["rmse"])
                if fitted.get("rmse") is not None else None
            ),
        }
        if key == "first_order":
            model_result["fit"] = fitted
        results.append(model_result)

    ranked = [
        model for model in results
        if model["success"] and model["r_squared"] is not None
    ]
    best = max(ranked, key=lambda model: model["r_squared"]) if ranked else None
    if best is not None:
        best["best"] = True
    for model in results:
        model.setdefault("best", False)
    return results, best


def _predict_model(model, time):
    parameters = model["parameters"]
    key = model["model_key"]
    if key == "zero_order":
        return parameters["k₀"] * time
    if key == "first_order":
        return parameters["R∞"] * -math.expm1(-parameters["k"] * time)
    if key == "higuchi":
        return parameters["kH"] * math.sqrt(max(time, 0.0))
    if key == "hixson_crowell":
        plateau = parameters["R∞"]
        return plateau - max(
            0.0,
            plateau ** (1.0 / 3.0) - parameters["kHC"] * time,
        ) ** 3
    if key == "korsmeyer_peppas":
        return 0.0 if time <= 0 else parameters["k"] * time ** parameters["n"]
    raise ValueError(f"Неизвестная модель: {key}")


def _sample_points(points, limit=MAX_PLOT_POINTS):
    if len(points) <= limit:
        return points
    stride = math.ceil(len(points) / limit)
    sampled = points[::stride][:limit]
    if sampled[-1] != points[-1]:
        sampled[-1] = points[-1]
    return sampled


def _build_model_plot(points, fit):
    if isinstance(fit, dict) and "model_key" not in fit and "plateau" in fit:
        models = [{
            "model_key": "first_order",
            "name": "First Order",
            "parameters": {"R∞": fit["plateau"], "k": fit["rate"]},
            "r_squared": fit.get("r_squared"),
            "success": True,
        }]
    elif isinstance(fit, dict):
        models = [fit] if fit.get("success", True) else []
    else:
        models = [model for model in fit if model.get("success")]
    if not models:
        raise ValueError("Нет успешно рассчитанных моделей для графика.")

    figure, axis = plt.subplots(figsize=(9.4, 5.4))
    image_buffer = io.BytesIO()
    try:
        sorted_points = sorted(points, key=lambda point: point[0])
        sd_points = [
            point for point in sorted_points
            if len(point) > 2 and point[2] is not None
        ]
        axis.set_axisbelow(True)
        axis.grid(True, which="major", color="#dce4ee", linewidth=0.8)
        if sd_points:
            axis.errorbar(
                [point[0] for point in sd_points],
                [point[1] for point in sd_points],
                yerr=[point[2] for point in sd_points],
                fmt="none",
                ecolor="#879bb4",
                elinewidth=1.0,
                capsize=3,
                label="± SD",
                zorder=2,
            )
        axis.scatter(
            [point[0] for point in sorted_points],
            [point[1] for point in sorted_points],
            label="Измерения",
            color="#1769aa",
            edgecolors="white",
            linewidths=0.7,
            s=34,
            zorder=3,
        )
        max_time = max(point[0] for point in sorted_points)
        curve_times = [max_time * index / 240 for index in range(241)]
        colors = ("#d97732", "#3b8c6e", "#8656a3", "#c34e62", "#667a9b")
        for index, model in enumerate(models):
            curve_releases = [
                _predict_model(model, time) for time in curve_times
            ]
            label = model["name"]
            if model.get("r_squared") is not None:
                label += f" (R²={_format_model_number(model['r_squared'])})"
            axis.plot(
                curve_times,
                curve_releases,
                label=label,
                color=colors[index % len(colors)],
                linewidth=2.0,
            )
        axis.set_xlabel("Время (единицы исходного файла)")
        axis.set_ylabel("Высвобождение (единицы исходного файла)")
        axis.set_title("Экспериментальные данные и кинетические модели")
        axis.margins(x=0.04, y=0.12)
        axis.legend(frameon=False, ncol=2)
        figure.tight_layout()
        figure.savefig(image_buffer, format="png", bbox_inches="tight", dpi=220)
        image_buffer.seek(0)
        return base64.b64encode(image_buffer.getvalue()).decode("ascii")
    finally:
        plt.close(figure)
        image_buffer.close()


def _build_source_coordinate_plot(fit, definition):
    """Plot finite transformed coordinates and their unweighted linear fit."""

    if not fit["success"] or len(fit["points"]) < 2:
        raise ValueError("Недостаточно координат для линейного графика.")
    plotted_points = _sample_points(fit["points"])
    min_x = min(x_value for x_value, _ in fit["points"])
    max_x = max(x_value for x_value, _ in fit["points"])
    line_y = [
        fit["slope"] * x_value + fit["intercept"]
        for x_value in (min_x, max_x)
    ]
    if not all(math.isfinite(value) for value in line_y):
        raise ValueError("Линия подгонки выходит за числовой диапазон графика.")

    figure, axis = plt.subplots(figsize=(9.0, 5.0))
    image_buffer = io.BytesIO()
    try:
        axis.set_axisbelow(True)
        axis.grid(True, which="major", color="#dce4ee", linewidth=0.8)
        axis.scatter(
            [point[0] for point in plotted_points],
            [point[1] for point in plotted_points],
            label="Преобразованные точки",
            color="#1769aa",
            edgecolors="white",
            linewidths=0.7,
            s=34,
            zorder=3,
        )
        axis.plot(
            [min_x, max_x],
            line_y,
            label="Линейная подгонка Django",
            color="#d97732",
            linewidth=2.0,
            zorder=2,
        )
        axis.set_xlabel(f"X = {definition['x_formula']}")
        axis.set_ylabel(f"Y = {definition['y_formula']}")
        axis.set_title(f"{definition['name']}: преобразованные координаты")
        axis.legend(frameon=False)
        figure.tight_layout()
        figure.savefig(image_buffer, format="png", dpi=140)
        return base64.b64encode(image_buffer.getvalue()).decode("ascii")
    finally:
        plt.close(figure)
        image_buffer.close()


def _build_comparison_plot(datasets):
    figure, axis = plt.subplots(figsize=(9.4, 5.4))
    image_buffer = io.BytesIO()
    try:
        axis.set_axisbelow(True)
        axis.grid(True, which="major", color="#dce4ee", linewidth=0.8)
        colors = ("#1769aa", "#d97732", "#3b8c6e", "#8656a3", "#c34e62")
        for index, dataset in enumerate(datasets):
            points = sorted(dataset["points"], key=lambda point: point[0])
            color = colors[index % len(colors)]
            sd_points = [
                point for point in points
                if len(point) > 2 and point[2] is not None
            ]
            if sd_points:
                axis.errorbar(
                    [point[0] for point in sd_points],
                    [point[1] for point in sd_points],
                    yerr=[point[2] for point in sd_points],
                    fmt="none",
                    ecolor=color,
                    elinewidth=1.0,
                    capsize=3,
                    zorder=2,
                )
            axis.plot(
                [point[0] for point in points],
                [point[1] for point in points],
                marker="o",
                markersize=4,
                markeredgecolor="white",
                markeredgewidth=0.7,
                linewidth=1.7,
                label=(
                    f"{dataset['experiment'].name} — "
                    f"{dataset.get('method', 'метод не сохранён')} "
                    f"({dataset.get('time_unit', 'единицы не указаны')}; "
                    f"{dataset.get('release_unit', 'единицы не указаны')})"
                ),
                color=color,
                zorder=3,
            )
        time_units = {dataset.get("time_unit") for dataset in datasets}
        release_units = {dataset.get("release_unit") for dataset in datasets}
        axis.set_xlabel(
            f"Время ({next(iter(time_units))})"
            if len(time_units) == 1
            else "Время (единицы указаны отдельно для каждой серии)"
        )
        axis.set_ylabel(
            f"Высвобождение ({next(iter(release_units))})"
            if len(release_units) == 1
            else "Высвобождение (шкала указана отдельно для каждой серии)"
        )
        axis.set_title("Сравнение экспериментов по фактическому времени")
        axis.margins(x=0.04, y=0.12)
        axis.legend(frameon=False)
        figure.tight_layout()
        figure.savefig(image_buffer, format="png", bbox_inches="tight", dpi=220)
        image_buffer.seek(0)
        return base64.b64encode(image_buffer.getvalue()).decode("ascii")
    finally:
        plt.close(figure)
        image_buffer.close()


TTP_QTP_NO_EXPERIMENT_MESSAGE = "Выберите эксперимент, чтобы рассчитать TTP и QTP."


def _calculate_ttp_qtp(point_rows):
    """Рассчитывает TTP/QTP в порядке строк исходных измерений, как в C# версии."""

    rows = list(point_rows)
    unavailable = {
        "ttp": None,
        "qtp": None,
        "ttp_qtp_threshold": None,
        "ttp_display": None,
        "qtp_display": None,
        "ttp_qtp_threshold_display": None,
        "ttp_qtp_message": "",
    }
    if not rows:
        unavailable["ttp_qtp_message"] = "В эксперименте нет измерений для расчёта TTP и QTP."
        return unavailable

    normalized = []
    for row in rows:
        if not isinstance(row, (list, tuple)) or len(row) < 2:
            unavailable["ttp_qtp_message"] = "В измерениях отсутствуют время или высвобождение."
            return unavailable
        try:
            time = float(row[0])
            release = float(row[1])
        except (TypeError, ValueError, OverflowError):
            unavailable["ttp_qtp_message"] = "В измерениях есть некорректные числовые значения."
            return unavailable
        if not math.isfinite(time) or not math.isfinite(release):
            unavailable["ttp_qtp_message"] = "В измерениях есть некорректные числовые значения."
            return unavailable
        sd_value = row[2] if len(row) > 2 else None
        normalized.append((time, release, sd_value))

    last_sd_value = normalized[-1][2]
    if last_sd_value is None or (
        isinstance(last_sd_value, str) and not last_sd_value.strip()
    ):
        unavailable["ttp_qtp_message"] = (
            "TTP и QTP не рассчитаны: у последнего измерения отсутствует SD."
        )
        return unavailable
    try:
        last_sd = float(last_sd_value)
    except (TypeError, ValueError, OverflowError):
        unavailable["ttp_qtp_message"] = "SD последнего измерения должно быть конечным неотрицательным числом."
        return unavailable
    if not math.isfinite(last_sd) or last_sd < 0:
        unavailable["ttp_qtp_message"] = "SD последнего измерения должно быть конечным неотрицательным числом."
        return unavailable

    threshold = normalized[-1][1] - last_sd / 2
    if not math.isfinite(threshold):
        unavailable["ttp_qtp_message"] = "Порог TTP/QTP выходит за допустимый числовой диапазон."
        return unavailable

    for time, release, _ in normalized:
        if release > threshold:
            return {
                "ttp": time,
                "qtp": release,
                "ttp_qtp_threshold": threshold,
                "ttp_display": _format_model_number(time),
                "qtp_display": _format_model_number(release),
                "ttp_qtp_threshold_display": _format_model_number(threshold),
                "ttp_qtp_message": "",
            }

    unavailable["ttp_qtp_message"] = (
        "Измерения не пересекают порог Q_last − SD_last / 2 по условию Q_i > P."
    )
    return unavailable


def _kinetic_analysis_context(experiment):
    measurement_rows = list(
        experiment.measurements.order_by("order", "pk")
        .values_list("time", "release", "sd").iterator(chunk_size=2_000)
    )
    points = []
    fit_excluded_indices = []
    for index, row in enumerate(measurement_rows, start=1):
        try:
            time = float(row[0])
            release = float(row[1])
        except (TypeError, ValueError, OverflowError):
            fit_excluded_indices.append(index)
            continue
        if not math.isfinite(time) or not math.isfinite(release):
            fit_excluded_indices.append(index)
            continue
        points.append((time, release))
    ttp_qtp = _calculate_ttp_qtp(measurement_rows)
    model_fits, best_model = _fit_kinetic_models(points)
    source_point_rows = [(row[0], row[1]) for row in measurement_rows]
    source_model_transforms = summarize_source_model_transforms(source_point_rows)
    for model in model_fits:
        model["input_point_count"] = len(measurement_rows)
        model["excluded_input_indices"] = fit_excluded_indices[:25]
        model["excluded_input_count"] = len(fit_excluded_indices)
        model["excluded_input_more_count"] = max(0, len(fit_excluded_indices) - 25)
    first_order = next(
        (model.get("fit") for model in model_fits if model["model_key"] == "first_order"),
        None,
    )
    first_order_row = next(
        (model for model in model_fits if model["model_key"] == "first_order"),
        None,
    )
    model_plot = None
    model_plot_error = False
    if any(model["success"] for model in model_fits):
        plot_points, _ = _plot_points(experiment)
        try:
            model_plot = _build_model_plot(plot_points, model_fits)
        except Exception:
            logger.exception("Не удалось построить график сравнения моделей")
            model_plot_error = True
    display = None
    if first_order is not None:
        display = {
            key: _format_model_number(first_order[key])
            for key in ("plateau", "rate", "rmse", "r_squared")
            if first_order.get(key) is not None
        }
    return {
        "model_fits": model_fits,
        "best_model": best_model,
        "source_model_transforms": source_model_transforms,
        "source_transform_experiment_id": experiment.pk,
        "model_fit": first_order,
        "model_fit_display": display,
        "model_plot": model_plot,
        "model_plot_error": model_plot_error,
        "model_error": (
            "Параметры рассчитаны, но построить график моделей не удалось."
            if model_plot_error else (first_order_row["error"] if first_order_row else "")
        ),
        **ttp_qtp,
    }


def _source_coordinate_context(
    experiment,
    model_key,
    page_number="1",
    *,
    request=None,
    include_experiment_id=False,
):
    """Return a reviewable page of imported or source-rule coordinates."""

    definition = next(
        (item for item in SOURCE_MODEL_DEFINITIONS if item["key"] == model_key),
        SOURCE_MODEL_DEFINITIONS[0],
    )
    imported = ImportedModelPoint.objects.filter(
        experiment=experiment,
        model_key=definition["key"],
    ).order_by("order", "pk")
    source_points = list(
        experiment.measurements.order_by("order", "pk")
        .values_list("time", "release")
    )
    transform_counts = next(
        (
            summary for summary in summarize_source_model_transforms(source_points)
            if summary["key"] == definition["key"]
        ),
        None,
    )
    if imported.exists():
        rows = [
            {
                "order": row.order,
                "time": None,
                "release": None,
                "x": row.x,
                "y": row.y,
                "sd": row.sd,
                "status": row.status,
            }
            for row in imported
            if row.status != "empty_table"
        ]
        coordinate_source = "imported"
    else:
        rows = list(transform_source_model(
            source_points,
            definition["key"],
            include_records=True,
        )["records"])
        coordinate_source = "calculated"

    linear_fit = fit_linear_trendline(
        (row.get("x"), row.get("y"))
        for row in rows
    )
    linear_fit_display = None
    linear_fit_plot = None
    linear_fit_plot_error = False
    if linear_fit["success"]:
        linear_fit_display = {
            "slope": _format_model_number(linear_fit["slope"]),
            "intercept": _format_model_number(linear_fit["intercept"]),
            "r_squared": (
                _format_model_number(linear_fit["r_squared"])
                if linear_fit["r_squared"] is not None else None
            ),
        }
        try:
            linear_fit_plot = _build_source_coordinate_plot(
                linear_fit,
                definition,
            )
        except Exception:
            logger.exception(
                "Не удалось построить линейный график координат модели %s",
                definition["key"],
            )
            linear_fit_plot_error = True

    page = Paginator(rows, TABLE_PAGE_SIZE).get_page(page_number)
    page_query = {"source_model_key": definition["key"]}
    if include_experiment_id:
        page_query["experiment_id"] = experiment.pk
    return {
        "selected_source_model": definition,
        "source_model_definitions": SOURCE_MODEL_DEFINITIONS,
        "source_coordinate_source": coordinate_source,
        "source_coordinate_count": len(rows),
        "source_coordinate_linear_fit": linear_fit,
        "source_coordinate_linear_fit_display": linear_fit_display,
        "source_coordinate_linear_plot": linear_fit_plot,
        "source_coordinate_linear_plot_error": linear_fit_plot_error,
        "source_transform_counts": transform_counts,
        "source_coordinate_page": page,
        "source_coordinate_experiment_id": experiment.pk,
        "imported_model_keys": set(
            ImportedModelPoint.objects.filter(experiment=experiment)
            .values_list("model_key", flat=True)
            .distinct()
        ),
        "source_coordinate_action": request.path if request else "",
        "source_coordinate_include_experiment_id": include_experiment_id,
        "source_coordinate_page_base": (
            f"{request.path}?{urlencode(page_query)}" if request else ""
        ),
    }


def kinetic_models(request):
    """Сравнивает кинетические модели выбранного эксперимента."""

    experiments = visible_experiments_for(request.user).annotate(
        measurement_count=Count("measurements"),
    ).order_by("-created_at", "-pk")
    context = {
        "experiments": experiments,
        "selected_experiment": None,
        "selected_id": request.GET.get("experiment_id", ""),
        "model_fit": None,
        "model_fit_display": None,
        "model_plot": None,
        "model_error": "",
        "model_fits": [],
        "best_model": None,
        "model_plot_error": False,
        "ttp_qtp_message": TTP_QTP_NO_EXPERIMENT_MESSAGE,
        "ttp": None,
        "qtp": None,
        "ttp_qtp_threshold": None,
        "source_model_definitions": SOURCE_MODEL_DEFINITIONS,
        "source_model_key": request.GET.get("source_model_key", "zero_order"),
    }
    if not context["selected_id"]:
        return render(request, "kinetics/kinetic_models.html", context)

    try:
        experiment_id = int(context["selected_id"])
    except (TypeError, ValueError):
        context["model_error"] = "Выберите сохранённый эксперимент из списка."
        return render(request, "kinetics/kinetic_models.html", context)
    if not 0 < experiment_id <= 9_223_372_036_854_775_807:
        context["model_error"] = "Выбранный эксперимент не найден."
        return render(request, "kinetics/kinetic_models.html", context)
    experiment = visible_experiments_for(request.user).filter(pk=experiment_id).first()
    if experiment is None:
        context["model_error"] = "Выбранный эксперимент не найден."
        return render(request, "kinetics/kinetic_models.html", context)

    context.update(_kinetic_analysis_context(experiment))
    context["selected_experiment"] = experiment
    context.update(_source_coordinate_context(
        experiment,
        context["source_model_key"],
        request.GET.get("source_page", "1"),
        request=request,
        include_experiment_id=True,
    ))
    context["source_model_key"] = context["selected_source_model"]["key"]
    return render(request, "kinetics/kinetic_models.html", context)


@require_http_methods(["GET"])
def source_transform_csv(request, pk, model_key):
    """Download the exact C#-style transformed coordinates for inspection."""

    experiment = get_object_or_404(visible_experiments_for(request.user), pk=pk)
    input_points = experiment.measurements.order_by("order", "pk").values_list(
        "time", "release"
    )
    try:
        transformed = transform_source_model(input_points.iterator(), model_key)
    except ValueError as exc:
        from django.http import Http404

        raise Http404("Неизвестное преобразование модели.") from exc

    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response.write("\ufeff")
    writer = csv.writer(response, lineterminator="\n")
    writer.writerow((
        "source_order",
        "time",
        "release",
        "transformed_x",
        "transformed_y",
        "status",
    ))
    status_labels = {
        "included": "included",
        "skipped_t0_source_rule": "skipped_t0_by_source_rule",
        "invalid_domain": "invalid_transform_domain",
        "invalid_input": "invalid_source_value",
    }
    for record in transformed["records"]:
        writer.writerow((
            record["order"],
            record["time"],
            record["release"],
            "" if record["x"] is None else record["x"],
            "" if record["y"] is None else record["y"],
            status_labels[record["status"]],
        ))
    response["Content-Disposition"] = (
        f'attachment; filename="experiment-{experiment.pk}-{model_key}-coordinates.csv"'
    )
    return response


@require_http_methods(["GET"])
def imported_model_points_csv(request, pk, model_key):
    """Download model coordinates exactly as they appeared in AllData.xlsx."""

    experiment = get_object_or_404(visible_experiments_for(request.user), pk=pk)
    points = ImportedModelPoint.objects.filter(
        experiment=experiment,
        model_key=model_key,
    ).order_by("order", "pk")
    if not points.exists():
        from django.http import Http404

        raise Http404("Для этого эксперимента нет импортированных координат модели.")

    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response.write("\ufeff")
    writer = csv.writer(response, lineterminator="\n")
    writer.writerow(("source_order", "transformed_x", "transformed_y", "sd", "status"))
    for point in points.iterator(chunk_size=2_000):
        writer.writerow((
            point.order,
            "" if point.x is None else point.x,
            "" if point.y is None else point.y,
            "" if point.sd is None else point.sd,
            point.status,
        ))
    response["Content-Disposition"] = (
        f'attachment; filename="experiment-{pk}-{model_key}-alldata.csv"'
    )
    return response


def _comparison_model_rows(experiment):
    """Calculate Django empirical fits and expose their actual sample counts."""

    source_rows = list(
        experiment.measurements.order_by("order", "pk")
        .values_list("time", "release")
    )
    points = []
    for time_value, release_value in source_rows:
        try:
            point = (float(time_value), float(release_value))
        except (TypeError, ValueError, OverflowError):
            continue
        if all(math.isfinite(value) for value in point):
            points.append(point)
    invalid_count = len(source_rows) - len(points)
    fits, _best = _fit_kinetic_models(points)
    rows = []
    for model in fits:
        parameter_text = "; ".join(
            f"{item['name']} = {item['value']}"
            for item in model["parameters_display"]
        )
        rows.append({
            "experiment": experiment,
            "model": model,
            "parameter_text": parameter_text,
            "valid_point_count": len(points),
            "invalid_point_count": invalid_count,
            "fit_point_count": model.get("point_count"),
        })
    return rows


def _comparison_compatibility(experiments):
    """Describe whether selected values share declared units, method and time grid."""

    warnings = []
    experiment_ids = [experiment.pk for experiment in experiments]
    has_source_values = any(
        experiment.input_mode != Experiment.InputMode.RAW
        for experiment in experiments
    )
    if has_source_values:
        warnings.append(
            "У готовых таблиц единицы Time/Release не сохраняются и значения "
            "не нормализуются. Сверьте заголовки исходных файлов перед научным сравнением."
        )

    if any(
        experiment.source_format == Experiment.SourceFormat.CSHARP_ALLDATA
        for experiment in experiments
    ):
        warnings.append(
            "Сравнение показывает сохранённые результаты Django. Импортированные "
            "значения AllData C# и координаты моделей остаются отдельным снимком и в этот график не подставляются."
        )

    method_signatures = {
        (experiment.input_mode, experiment.calculation_version or "")
        for experiment in experiments
    }
    if len(method_signatures) > 1:
        warnings.append(
            "У выбранных опытов отличаются режим входных данных или версия расчёта. "
            "Показанные значения не преобразованы к одной методике."
        )

    if any(experiment.source_results_stale for experiment in experiments):
        warnings.append(
            "Как минимум один импортированный снимок C# помечен устаревшим; "
            "график и таблицы сравнения используют сохранённые результаты Django."
        )

    times_by_experiment = {experiment_id: [] for experiment_id in experiment_ids}
    for experiment_id, value in Measurement.objects.filter(
        experiment_id__in=experiment_ids
    ).values_list("experiment_id", "time").iterator(chunk_size=2_000):
        try:
            time_value = float(value)
        except (TypeError, ValueError, OverflowError):
            continue
        if math.isfinite(time_value):
            times_by_experiment[experiment_id].append(time_value)
    time_grids = {
        experiment_id: tuple(sorted(values))
        for experiment_id, values in times_by_experiment.items()
    }
    if len(set(time_grids.values())) > 1:
        warnings.append(
            "Наборы временных точек или число измерений различаются. "
            "График строится по фактическому времени без интерполяции; "
            "строки разных опытов не соответствуют друг другу по номеру."
        )

    for experiment in experiments:
        if experiment.input_mode == Experiment.InputMode.RAW:
            experiment.comparison_time_unit = "мин"
            experiment.comparison_release_unit = "%"
            experiment.comparison_method = (
                f"Django {experiment.calculation_version or 'версия не сохранена'}"
            )
        else:
            experiment.comparison_time_unit = "по файлу; не задано"
            experiment.comparison_release_unit = "по файлу; не задано"
            experiment.comparison_method = "Готовые значения; метод не сохранён"

    return warnings, time_grids


def comparison(request):
    """Сравнивает несколько экспериментов, сохраняя их исходные значения времени."""

    visible_experiments = visible_experiments_for(request.user)
    experiments = visible_experiments.annotate(
        measurement_count=Count("measurements"),
    ).order_by("-created_at", "-pk")
    selected_ids = []
    selection_error = ""
    raw_ids = request.GET.getlist("experiment_ids")
    if raw_ids:
        try:
            selected_ids = list(dict.fromkeys(int(value) for value in raw_ids))
            if any(
                value <= 0 or value > 9_223_372_036_854_775_807
                for value in selected_ids
            ):
                raise ValueError
        except (TypeError, ValueError):
            selected_ids = []
            selection_error = "Выберите эксперименты из списка."

    selected_experiments = []
    datasets = []
    comparison_plot = None
    measurement_page = None
    comparison_message = ""
    model_comparison_rows = []
    model_comparison_error = ""
    compatibility_warnings = []
    model_fits_requested = request.GET.get("compare_models") == "1"
    if selected_ids and not selection_error:
        selected_by_id = {
            experiment.pk: experiment
            for experiment in visible_experiments.filter(pk__in=selected_ids).annotate(
                measurement_count=Count("measurements"),
            )
        }
        if len(selected_by_id) != len(selected_ids):
            selection_error = "Один из выбранных экспериментов уже недоступен. Обновите выбор."
        elif len(selected_ids) < 2:
            comparison_message = "Для сравнения выберите как минимум два эксперимента."
        else:
            selected_experiments = [selected_by_id[pk] for pk in selected_ids]
            compatibility_warnings, time_grids = _comparison_compatibility(
                selected_experiments
            )
            aggregates = {
                summary["experiment_id"]: summary
                for summary in Measurement.objects.filter(
                    experiment_id__in=selected_ids,
                ).values("experiment_id").annotate(
                    measurement_count=Count("pk"),
                    time_min=Min("time"),
                    time_max=Max("time"),
                    release_min=Min("release"),
                    release_max=Max("release"),
                )
            }
            for experiment in selected_experiments:
                summary = aggregates.get(experiment.pk, {})
                count = summary.get("measurement_count", 0)
                experiment_summary = {
                    "experiment": experiment,
                    "measurement_count": count,
                    "time_min": summary.get("time_min"),
                    "time_max": summary.get("time_max"),
                    "release_min": summary.get("release_min"),
                    "release_max": summary.get("release_max"),
                    "time_grid_matches": len(set(time_grids.values())) <= 1,
                    "time_unit": experiment.comparison_time_unit,
                    "release_unit": experiment.comparison_release_unit,
                    "method": experiment.comparison_method,
                }
                if count:
                    points = _finite_plot_points(
                        experiment.measurements.order_by("time", "order", "pk")
                        .values_list("time", "release", "sd").iterator(chunk_size=2_000)
                    )
                    datasets.append({
                        **experiment_summary,
                        "points": _sample_points(points),
                        "valid_point_count": len(points),
                    })
                else:
                    datasets.append({
                        **experiment_summary,
                        "points": [],
                        "valid_point_count": 0,
                    })

            plottable = [
                dataset for dataset in datasets
                if dataset["valid_point_count"] >= 2
            ]
            if len(plottable) >= 2:
                try:
                    comparison_plot = _build_comparison_plot(plottable)
                except Exception:
                    logger.exception("Не удалось построить график сравнения")
                    comparison_message = "Не удалось построить совместный график; таблицы сравнения доступны."
            else:
                comparison_message = "Для совместного графика нужны измерения как минимум из двух выбранных экспериментов."

            measurement_queryset = Measurement.objects.filter(
                experiment_id__in=selected_ids,
            ).select_related("experiment").order_by(
                "time", "experiment_id", "order", "pk",
            )
            measurement_page = Paginator(measurement_queryset, TABLE_PAGE_SIZE).get_page(
                request.GET.get("page"),
            )
            for measurement in measurement_page.object_list:
                selected = selected_by_id[measurement.experiment_id]
                measurement.experiment.comparison_time_unit = selected.comparison_time_unit
                measurement.experiment.comparison_release_unit = selected.comparison_release_unit
                measurement.experiment.comparison_method = selected.comparison_method
            if model_fits_requested:
                for experiment in selected_experiments:
                    try:
                        model_comparison_rows.extend(
                            _comparison_model_rows(experiment)
                        )
                    except Exception:
                        logger.exception(
                            "Не удалось сравнить модели эксперимента %s",
                            experiment.pk,
                        )
                        model_comparison_error = (
                            "Не удалось рассчитать одну из подгонок моделей. "
                            "Проверьте данные экспериментов; остальные результаты доступны."
                        )

    return render(request, "kinetics/comparison.html", {
        "experiments": experiments,
        "selected_ids": {str(value) for value in selected_ids},
        "selection_query": "&".join(
            f"experiment_ids={value}" for value in selected_ids
        ),
        "selected_experiments": selected_experiments,
        "datasets": datasets,
        "comparison_plot": comparison_plot,
        "measurement_page": measurement_page,
        "selection_error": selection_error,
        "comparison_message": comparison_message,
        "model_fits_requested": model_fits_requested,
        "model_comparison_rows": model_comparison_rows,
        "model_comparison_error": model_comparison_error,
        "compatibility_warnings": compatibility_warnings,
    })


def comparison_export_xlsx(request):
    """Export the selected experiments' summaries and long-form observations."""

    try:
        selected_ids = list(dict.fromkeys(
            int(value) for value in request.GET.getlist("experiment_ids")
        ))
        if any(
            value <= 0 or value > 9_223_372_036_854_775_807
            for value in selected_ids
        ):
            raise ValueError
    except (TypeError, ValueError):
        messages.error(request, "Выберите эксперименты из списка для экспорта.")
        return redirect("comparison")

    if len(selected_ids) < 2:
        messages.error(request, "Для экспорта сравнения выберите как минимум два эксперимента.")
        return redirect("comparison")

    selected_by_id = {
        experiment.pk: experiment
        for experiment in visible_experiments_for(request.user).filter(pk__in=selected_ids)
    }
    if len(selected_by_id) != len(selected_ids):
        messages.error(request, "Один из выбранных экспериментов недоступен; обновите выбор.")
        return redirect("comparison")
    selected_experiments = [selected_by_id[pk] for pk in selected_ids]
    compatibility_warnings, time_grids = _comparison_compatibility(selected_experiments)

    aggregates = {
        row["experiment_id"]: row
        for row in Measurement.objects.filter(
            experiment_id__in=selected_ids
        ).values("experiment_id").annotate(
            measurement_count=Count("pk"),
            time_min=Min("time"),
            time_max=Max("time"),
            release_min=Min("release"),
            release_max=Max("release"),
        )
    }
    replicate_counts = {
        row["experiment_id"]: row["replicate_count"]
        for row in RawReplicate.objects.filter(
            experiment_id__in=selected_ids
        ).values("experiment_id").annotate(replicate_count=Count("pk"))
    }

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Experiments"
    summary.append((
        "Experiment ID", "Name", "Input mode", "Calculation method/version",
        "Drug loading (DL)", "Entrapment efficiency (EE), direct multiplier",
        "Replicates", "Measurements", "Minimum time", "Maximum time",
        "Minimum Release", "Maximum Release", "Time unit", "Release unit",
        "Time grid matches", "C# snapshot stale", "Source format",
    ))
    for experiment in selected_experiments:
        aggregate = aggregates.get(experiment.pk, {})
        summary.append((
            experiment.pk,
            _excel_text(experiment.name),
            _excel_text(experiment.get_input_mode_display()),
            _excel_text(experiment.calculation_version or "not saved"),
            experiment.drug_loading,
            experiment.entrapment_efficiency,
            replicate_counts.get(experiment.pk, 0),
            aggregate.get("measurement_count", 0),
            aggregate.get("time_min"),
            aggregate.get("time_max"),
            aggregate.get("release_min"),
            aggregate.get("release_max"),
            _excel_text(experiment.comparison_time_unit),
            _excel_text(experiment.comparison_release_unit),
            len(set(time_grids.values())) <= 1,
            "yes" if experiment.source_results_stale else "no",
            _excel_text(experiment.get_source_format_display()),
        ))
    _size_export_columns(summary)

    observations = workbook.create_sheet("Measurements by time")
    observations.append((
        "Experiment ID", "Experiment", "Method/version", "Time unit",
        "Time", "Release unit", "Release", "SD",
    ))
    for row in Measurement.objects.filter(
        experiment_id__in=selected_ids
    ).select_related("experiment").order_by(
        "time", "experiment_id", "order", "pk"
    ).iterator(chunk_size=2_000):
        comparison_experiment = selected_by_id[row.experiment_id]
        observations.append((
            row.experiment_id,
            _excel_text(row.experiment.name),
            _excel_text(comparison_experiment.comparison_method),
            _excel_text(comparison_experiment.comparison_time_unit),
            row.time,
            _excel_text(comparison_experiment.comparison_release_unit),
            row.release,
            row.sd,
        ))
    _size_export_columns(observations)

    notes = workbook.create_sheet("Comparison notes")
    notes.append(("Type", "Details"))
    notes.append((
        "Interpretation",
        "Values are exported without unit conversion, normalization, or interpolation. "
        "Review each experiment's method and units before drawing conclusions.",
    ))
    if compatibility_warnings:
        for warning in compatibility_warnings:
            notes.append(("Compatibility warning", _excel_text(warning)))
    else:
        notes.append((
            "Compatibility",
            "Selected experiments use the same declared Django method/version and units.",
        ))
    for experiment in selected_experiments:
        notes.append((
            "Experiment",
            _excel_text(
                f"{experiment.name}: {experiment.comparison_method}; "
                f"Time={experiment.comparison_time_unit}; "
                f"Release={experiment.comparison_release_unit}; "
                f"time grid points={len(time_grids[experiment.pk])}"
            ),
        ))
    _size_export_columns(notes)

    if request.GET.get("compare_models") == "1":
        model_fits_sheet = workbook.create_sheet("Django model fits")
        model_fits_sheet.append((
            "Experiment ID", "Experiment", "Method/version", "Time unit",
            "Release unit", "Model", "Equation", "Parameters",
            "R squared on Release", "RMSE on Release", "Points used",
            "Valid points", "Invalid inputs excluded", "Status",
        ))
        for experiment in selected_experiments:
            for row in _comparison_model_rows(experiment):
                model = row["model"]
                model_fits_sheet.append((
                    experiment.pk,
                    _excel_text(experiment.name),
                    _excel_text(experiment.comparison_method),
                    _excel_text(experiment.comparison_time_unit),
                    _excel_text(experiment.comparison_release_unit),
                    _excel_text(model["name"]),
                    _excel_text(model["formula"]),
                    _excel_text(row["parameter_text"]),
                    model.get("r_squared"),
                    model.get("rmse"),
                    row["fit_point_count"],
                    row["valid_point_count"],
                    row["invalid_point_count"],
                    _excel_text("success" if model["success"] else model["error"]),
                ))
        _size_export_columns(model_fits_sheet)

    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()
    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="apollinex-comparison.xlsx"'
    return response


def reports(request):
    """Формирует страницу отчёта выбранного эксперимента с печатью браузера."""

    visible_experiments = visible_experiments_for(request.user)
    experiments = visible_experiments.annotate(
        measurement_count=Count("measurements"),
    ).order_by("-created_at", "-pk")
    selected_id = request.GET.get("experiment_id", "")
    experiment = None
    report_error = ""
    report_measurements = []
    report_plot = None
    report_plot_error = False
    report_summary = None
    report_model_fits = []
    report_best_model = None
    report_source_model_transforms = []
    report_source_transform_experiment_id = None
    report_model_plot = None
    report_model_plot_error = False
    report_raw_replicates = []
    report_raw_measurements = []
    report_ttp = None
    report_qtp = None
    report_ttp_qtp_threshold = None
    report_ttp_display = None
    report_qtp_display = None
    report_ttp_qtp_threshold_display = None
    report_source_model_key = request.GET.get("source_model_key", "zero_order")
    source_coordinate_context = {}
    ttp_qtp_message = TTP_QTP_NO_EXPERIMENT_MESSAGE
    if selected_id:
        try:
            experiment_id = int(selected_id)
        except (TypeError, ValueError):
            experiment_id = None
            report_error = "Выберите эксперимент из списка."
        if experiment_id is not None and not 0 < experiment_id <= 9_223_372_036_854_775_807:
            experiment_id = None
            report_error = "Выбранный эксперимент не найден."
        if experiment_id is not None:
            experiment = visible_experiments.filter(pk=experiment_id).first()
            if experiment is None:
                report_error = "Выбранный эксперимент не найден."
            else:
                measurements = experiment.measurements.order_by("order", "pk")
                report_summary = measurements.aggregate(
                    measurement_count=Count("pk"),
                    time_min=Min("time"),
                    time_max=Max("time"),
                    release_min=Min("release"),
                    release_max=Max("release"),
                )
                report_measurements = list(measurements)
                if experiment.input_mode == Experiment.InputMode.RAW:
                    report_raw_replicates = list(
                        RawReplicate.objects.filter(experiment=experiment)
                        .order_by("replicate_number")
                    )
                    report_raw_measurements = list(
                        RawMeasurement.objects.filter(
                            replicate__experiment=experiment
                        ).select_related(
                            "replicate", "calculation", "source_calculation"
                        ).order_by(
                            "replicate__replicate_number", "order", "pk"
                        )
                    )
                points, fingerprint = _plot_points(
                    experiment,
                    report_summary["measurement_count"],
                )
                report_plot = _get_or_build_plot(points, fingerprint)
                report_plot_error = bool(points) and report_plot is None
                analysis = _kinetic_analysis_context(experiment)
                source_coordinate_context = _source_coordinate_context(
                    experiment,
                    report_source_model_key,
                    request.GET.get("source_page", "1"),
                    request=request,
                    include_experiment_id=True,
                )
                report_source_model_key = source_coordinate_context[
                    "selected_source_model"
                ]["key"]
                report_model_fits = analysis["model_fits"]
                report_best_model = analysis["best_model"]
                report_source_model_transforms = analysis["source_model_transforms"]
                report_source_transform_experiment_id = analysis[
                    "source_transform_experiment_id"
                ]
                report_model_plot = analysis["model_plot"]
                report_model_plot_error = analysis["model_plot_error"]
                report_ttp = analysis["ttp"]
                report_qtp = analysis["qtp"]
                report_ttp_qtp_threshold = analysis["ttp_qtp_threshold"]
                report_ttp_display = analysis["ttp_display"]
                report_qtp_display = analysis["qtp_display"]
                report_ttp_qtp_threshold_display = analysis[
                    "ttp_qtp_threshold_display"
                ]
                ttp_qtp_message = analysis["ttp_qtp_message"]

    return render(request, "kinetics/reports.html", {
        "experiments": experiments,
        "selected_id": selected_id,
        "experiment": experiment,
        "report_error": report_error,
        "report_measurements": report_measurements,
        "report_raw_replicates": report_raw_replicates,
        "report_raw_measurements": report_raw_measurements,
        "report_plot": report_plot,
        "report_plot_error": report_plot_error,
        "report_summary": report_summary,
        "report_model_fits": report_model_fits,
        "report_best_model": report_best_model,
        "report_source_model_transforms": report_source_model_transforms,
        "report_source_transform_experiment_id": report_source_transform_experiment_id,
        "report_model_plot": report_model_plot,
        "report_model_plot_error": report_model_plot_error,
        "ttp_qtp_message": ttp_qtp_message,
        "ttp": report_ttp,
        "qtp": report_qtp,
        "ttp_qtp_threshold": report_ttp_qtp_threshold,
        "ttp_display": report_ttp_display,
        "qtp_display": report_qtp_display,
        "ttp_qtp_threshold_display": report_ttp_qtp_threshold_display,
        "source_model_definitions": SOURCE_MODEL_DEFINITIONS,
        "source_model_key": report_source_model_key,
        **source_coordinate_context,
    })


def home(request):
    """Показывает сводку и график последнего выбранного эксперимента."""

    context = _experiment_context(_selected_experiment(request), request)
    context["upload_token"] = _get_upload_token()
    return render(request, "kinetics/home.html", context)


def method_documentation(request):
    """Show source-backed formulas, units, edge rules and comparison limits."""

    return render(request, "kinetics/method_documentation.html")


def data_page(request):
    """Показывает таблицу измерений выбранного или последнего эксперимента."""

    experiment = _selected_experiment(request)
    context = _experiment_context(
        experiment,
        request,
        include_table=True,
        include_plot=False,
    )
    return render(request, "kinetics/data.html", context)


def experiment_history(request):
    """Показывает сохранённые эксперименты с постраничной навигацией."""

    experiments = Paginator(
        visible_experiments_for(request.user).annotate(
            measurement_count=Count("measurements"),
        ).order_by("-created_at", "-pk"),
        HISTORY_PAGE_SIZE,
    ).get_page(request.GET.get("page"))
    return render(request, "kinetics/experiment_history.html", {
        "experiments_page": experiments,
    })


def experiment_detail(request, pk):
    """Загружает выбранный эксперимент и связанные точки непосредственно из БД."""

    experiment = get_object_or_404(visible_experiments_for(request.user), pk=pk)
    request.session[CURRENT_EXPERIMENT_SESSION_KEY] = experiment.pk
    context = _experiment_context(experiment, request, include_table=True)
    context["can_edit_experiment"] = can_edit_experiment(request.user, experiment)
    context.update(_kinetic_analysis_context(experiment))
    context.update(_source_coordinate_context(
        experiment,
        request.GET.get("source_model_key", "zero_order"),
        request.GET.get("source_page", "1"),
        request=request,
    ))
    if experiment.input_mode == Experiment.InputMode.RAW:
        raw_queryset = RawMeasurement.objects.filter(
            replicate__experiment=experiment
        ).select_related(
            "replicate", "calculation", "source_calculation"
        ).order_by(
            "replicate__replicate_number", "order", "pk"
        )
        context.update({
            "raw_replicates": RawReplicate.objects.filter(
                experiment=experiment
            ).order_by("replicate_number"),
            "raw_measurements_page": Paginator(
                raw_queryset,
                TABLE_PAGE_SIZE,
            ).get_page(request.GET.get("raw_page")),
            "raw_measurement_count": raw_queryset.count(),
        })
    return render(request, "kinetics/experiment_detail.html", context)


def _excel_text(value):
    """Keep untrusted text cells from being interpreted as spreadsheet formulas."""

    text = "" if value is None else str(value)
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def _size_export_columns(sheet):
    for column_cells in sheet.columns:
        letter = column_cells[0].column_letter
        width = min(max(
            len(str(cell.value or ""))
            for cell in column_cells
        ) + 2, 42)
        sheet.column_dimensions[letter].width = max(width, 12)
    sheet.freeze_panes = "A2"
    if sheet.max_row > 1:
        sheet.auto_filter.ref = sheet.dimensions


def export_experiment_xlsx(request, pk):
    """Download a reproducible workbook with source data, analysis and plots."""

    experiment = get_object_or_404(visible_experiments_for(request.user), pk=pk)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Experiment"
    summary.append(("Parameter", "Value"))
    summary_rows = [
        ("Name", experiment.name),
        ("Description", experiment.description),
        ("Source file", experiment.source_file),
        ("Source format", experiment.get_source_format_display()),
        ("Input mode", experiment.get_input_mode_display()),
        ("Calculation method/version", experiment.calculation_version or "not saved"),
        ("Imported C# snapshot stale", "yes" if experiment.source_results_stale else "no"),
        ("Time unit", "minutes" if experiment.input_mode == Experiment.InputMode.RAW else "as provided by source"),
        ("Release unit", "percent" if experiment.input_mode == Experiment.InputMode.RAW else "not recorded by source"),
        ("Synthetic example", "yes — educational data only" if experiment.source_format == Experiment.SourceFormat.SYNTHETIC_RAW else "no"),
    ]
    for key, value in summary_rows:
        summary.append((_excel_text(key), _excel_text(value) if isinstance(value, str) or value is None else value))
    _size_export_columns(summary)

    parameters = workbook.create_sheet("Input parameters")
    parameters.append(("Parameter", "Value", "Unit / scale", "Availability and interpretation"))
    parameter_rows = (
        ("Drug loading (DL)", experiment.drug_loading, "dimensionless ratio", "As entered; the source label and formula do not establish a percent conversion."),
        ("Entrapment efficiency (EE)", experiment.entrapment_efficiency, "unconfirmed input scale", "The C# formula uses the entered number directly and does not divide by 100."),
        ("Reactor volume", experiment.reactor_volume_ml, "mL", "Volume used in the Mreal calculation."),
        ("Probe volume", experiment.probe_volume_ml, "mL", "Volume used for the correction for previously withdrawn samples."),
        ("Calibration coefficient (k)", experiment.calibration_coefficient, "as calibrated", "Multiplies signal A to obtain Cabs; the source unit is not recorded."),
    )
    for label, value, unit, explanation in parameter_rows:
        parameters.append((
            label,
            value,
            unit,
            explanation if value is not None else "Not saved in this experiment.",
        ))
    parameters.append((
        "Number of replicates",
        RawReplicate.objects.filter(experiment=experiment).count(),
        "replicates",
        "Count of saved source replicates.",
    ))
    _size_export_columns(parameters)

    results = workbook.create_sheet("Mean Release SD")
    results.append(("Point", "Time", "Mean Release", "SD"))
    measurements = Measurement.objects.filter(
        experiment=experiment
    ).order_by("order", "pk")
    for index, measurement in enumerate(measurements, start=1):
        results.append((
            index,
            measurement.time,
            measurement.release,
            measurement.sd,
        ))
    _size_export_columns(results)

    raw_rows = RawMeasurement.objects.filter(
        replicate__experiment=experiment
    ).select_related("replicate", "calculation", "source_calculation").order_by(
        "replicate__replicate_number", "order", "pk"
    )
    raw = workbook.create_sheet("Replicate readings")
    raw.append((
        "Replicate", "Sample mass (g)", "Point", "Time (min)",
        "Signal A (instrument units)",
    ))
    raw_calculations = workbook.create_sheet("Django replicate calculations")
    raw_calculations.append((
        "Replicate", "Point", "Cabs", "Mreal (mg)", "Release (%)",
        "Algorithm version",
    ))
    imported_calculations_exist = ImportedReplicateCalculation.objects.filter(
        measurement__replicate__experiment=experiment
    ).exists()
    source_calculations = None
    if imported_calculations_exist:
        source_calculations = workbook.create_sheet("C# imported results")
        source_calculations.append((
            "Replicate", "Point", "Time (source unit)", "Signal A",
            "Cabs (source)", "Mreal (source)", "Release (source)", "Source format",
        ))

    for row in raw_rows:
        raw.append((
            row.replicate.replicate_number,
            row.replicate.sample_mass_g,
            row.order + 1,
            row.time,
            row.signal_a,
        ))
        calculated = getattr(row, "calculation", None)
        if calculated:
            raw_calculations.append((
                row.replicate.replicate_number,
                row.order + 1,
                calculated.c_abs,
                calculated.m_real,
                calculated.release,
                calculated.algorithm_version,
            ))
        if source_calculations is not None:
            source = getattr(row, "source_calculation", None)
            if source:
                source_calculations.append((
                    row.replicate.replicate_number,
                    row.order + 1,
                    row.time,
                    row.signal_a,
                    source.c_abs,
                    source.m_real,
                    source.release,
                    _excel_text(source.source_format),
                ))
    for sheet in (raw, raw_calculations):
        _size_export_columns(sheet)
    if source_calculations is not None:
        _size_export_columns(source_calculations)

    model_points = ImportedModelPoint.objects.filter(
        experiment=experiment
    ).order_by("model_key", "order", "pk")
    if model_points.exists():
        coordinates = workbook.create_sheet("CSharp coordinates")
        coordinates.append(("Model", "Point", "X", "Y", "SD", "Status"))
        for point in model_points:
            coordinates.append((
                point.model_key,
                point.order,
                point.x,
                point.y,
                point.sd,
                point.status,
            ))
        _size_export_columns(coordinates)

    analysis = _kinetic_analysis_context(experiment)
    ttp_sheet = workbook.create_sheet("TTP QTP")
    ttp_sheet.append(("Django result", "Value", "Unit / rule"))
    ttp_sheet.append(("Calculation method/version", _excel_text(experiment.calculation_version or "not saved"), "saved method"))
    ttp_sheet.append(("TTP", analysis["ttp"], "same unit as time" if experiment.input_mode == Experiment.InputMode.RAW else "source unit not recorded"))
    ttp_sheet.append(("QTP", analysis["qtp"], "same unit as time" if experiment.input_mode == Experiment.InputMode.RAW else "source unit not recorded"))
    ttp_sheet.append(("Threshold P = Q_last − SD_last / 2", analysis["ttp_qtp_threshold"], "Release scale; SD is never inferred"))
    ttp_sheet.append((
        "Status",
        _excel_text(
            "Strict crossing Q_i > P in source order"
            if analysis["ttp"] is not None else analysis["ttp_qtp_message"]
        ),
        "Django calculation from saved mean Release/SD",
    ))
    _size_export_columns(ttp_sheet)

    fits_sheet = workbook.create_sheet("Django model fits")
    fits_sheet.append((
        "Model", "Equation", "Parameters", "R² on Release (%)",
        "RMSE on Release (%)", "Points used", "Valid input points",
        "Excluded non-finite rows", "Fit method / metric scale", "Status",
    ))
    for model in analysis["model_fits"]:
        parameters = model.get("parameters", {})
        parameter_text = "; ".join(
            f"{name} = {_format_model_number(value)}"
            for name, value in parameters.items()
        )
        fits_sheet.append((
            _excel_text(model["name"]),
            _excel_text(model["formula"]),
            _excel_text(parameter_text),
            model.get("r_squared"),
            model.get("rmse"),
            model.get("point_count"),
            model.get("input_point_count", 0) - model.get("excluded_input_count", 0),
            ", ".join(str(index) for index in model.get("excluded_input_indices", [])),
            "Unweighted least-squares; metrics on original Release scale",
            _excel_text("Calculated" if model["success"] else model["error"]),
        ))
    for column, width in {
        "A": 36, "B": 54, "C": 36, "D": 18, "E": 18,
        "F": 16, "G": 20, "H": 24, "I": 48, "J": 58,
    }.items():
        fits_sheet.column_dimensions[column].width = width
    fits_sheet.freeze_panes = "A2"
    if fits_sheet.max_row > 1:
        fits_sheet.auto_filter.ref = fits_sheet.dimensions

    charts = workbook.create_sheet("Charts")
    charts.sheet_view.showGridLines = False
    charts["A1"] = "Графики по сохранённым данным"
    charts["A2"] = (
        "Первый график показывает сохранённые средние Release и SD. "
        "Кривые Django — отдельные эмпирические подгонки; C#-координаты приведены на другой вкладке."
    )
    charts["A4"] = "Среднее Release и SD"
    plot_points, plot_fingerprint = _plot_points(experiment)
    aggregate_plot = _get_or_build_plot(plot_points, plot_fingerprint)
    if aggregate_plot:
        charts["A5"] = "Расчёт Django по сохранённым средним значениям"
    else:
        charts["A5"] = "Нет корректных измерений для графика."
    charts["A34"] = "Эмпирические подгонки Django"
    if analysis["model_plot"]:
        charts["A35"] = "Метрики и уравнения приведены на вкладке Django model analysis."
    else:
        charts["A35"] = "График моделей недоступен: недостаточно точек или ни одна подгонка не рассчитана."

    image_buffers = []

    def add_chart_image(encoded_png, anchor):
        if not encoded_png:
            return False
        try:
            image_buffer = io.BytesIO(base64.b64decode(encoded_png, validate=True))
            image = SpreadsheetImage(image_buffer)
            original_width = image.width
            original_height = image.height
            image.width = 760
            image.height = round(original_height * image.width / original_width)
            charts.add_image(image, anchor)
            image_buffers.append(image_buffer)
            return True
        except Exception:
            logger.exception("Не удалось добавить график в Excel-отчёт")
            return False

    add_chart_image(aggregate_plot, "A6")
    add_chart_image(analysis["model_plot"], "A36")
    charts.column_dimensions["A"].width = 90

    buffer = io.BytesIO()
    try:
        workbook.save(buffer)
    finally:
        workbook.close()
        for image_buffer in image_buffers:
            image_buffer.close()
    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = (
        f'attachment; filename="apollinex-experiment-{experiment.pk}.xlsx"'
    )
    return response


def _persist_raw_experiment(
    *, name, description, source_file, parameters, replicates,
    source_format=Experiment.SourceFormat.MANUAL_RAW, upload_token=None,
    owner=None,
):
    """Persist source readings and all derived outputs as one transaction."""

    with transaction.atomic():
        experiment = Experiment.objects.create(
            name=name[:200],
            description=description,
            source_file=source_file,
            owner=owner,
            input_mode=Experiment.InputMode.RAW,
            source_format=source_format,
            upload_token=upload_token,
            **parameters,
        )
        for replicate_input in replicates:
            replicate = RawReplicate.objects.create(
                experiment=experiment,
                replicate_number=replicate_input.number,
                sample_mass_g=replicate_input.sample_mass_g,
            )
            raw_rows = [
                RawMeasurement(
                    replicate=replicate,
                    time=point.time,
                    signal_a=point.signal_a,
                    order=index,
                )
                for index, point in enumerate(replicate_input.points)
            ]
            RawMeasurement.objects.bulk_create(raw_rows)
            if any(
                point.source_c_abs is not None
                or point.source_m_real is not None
                or point.source_release is not None
                for point in replicate_input.points
            ):
                persisted_rows = list(
                    replicate.measurements.order_by("order", "pk")
                )
                ImportedReplicateCalculation.objects.bulk_create([
                    ImportedReplicateCalculation(
                        measurement=raw_row,
                        c_abs=point.source_c_abs,
                        m_real=point.source_m_real,
                        release=point.source_release,
                        source_format="csharp-alldata",
                    )
                    for raw_row, point in zip(persisted_rows, replicate_input.points)
                    if (
                        point.source_c_abs is not None
                        or point.source_m_real is not None
                        or point.source_release is not None
                    )
                ])
        recalculate_experiment(experiment)
    return experiment


def _persist_csharp_sheet(
    *, sheet, source_filename, upload_token, is_all_data=False, owner=None
):
    """Save source readings and, for AllData, its separate C# coordinate tables."""

    parameters = {
        "drug_loading": sheet.parameters.drug_loading,
        "entrapment_efficiency": sheet.parameters.entrapment_efficiency,
        "reactor_volume_ml": sheet.parameters.reactor_volume_ml,
        "probe_volume_ml": sheet.parameters.probe_volume_ml,
        "calibration_coefficient": sheet.parameters.calibration_coefficient,
    }
    with transaction.atomic():
        experiment = _persist_raw_experiment(
            name=sheet.name,
            description=(
                "Импортировано из книги AllData C#; исходные координаты моделей "
                "сохранены отдельно от расчёта Django."
                if is_all_data else "Импортировано из исходного Excel-листа C#."
            ),
            source_file=source_filename,
            owner=owner,
            parameters=parameters,
            replicates=sheet.replicates,
            source_format=(
                Experiment.SourceFormat.CSHARP_ALLDATA
                if is_all_data else Experiment.SourceFormat.CSHARP_RAW
            ),
            upload_token=upload_token,
        )
        if is_all_data and sheet.model_points:
            ImportedModelPoint.objects.bulk_create([
                ImportedModelPoint(
                    experiment=experiment,
                    model_key=point.model_key,
                    order=point.order,
                    x=point.x,
                    y=point.y,
                    sd=point.sd,
                    status=point.status,
                )
                for point in sheet.model_points
            ])
    return experiment


def _raw_experiment_initial(experiment):
    """Build a round-trippable form value from immutable stored readings."""

    rows = []
    replicates = RawReplicate.objects.filter(
        experiment=experiment
    ).prefetch_related("measurements").order_by("replicate_number", "pk")
    for replicate in replicates:
        for point in replicate.measurements.all().order_by("order", "pk"):
            rows.append(
                f"{replicate.replicate_number}\t{replicate.sample_mass_g!r}\t"
                f"{point.time!r}\t{point.signal_a!r}"
            )
    return {
        "name": experiment.name,
        "description": experiment.description,
        "drug_loading": experiment.drug_loading,
        "entrapment_efficiency": experiment.entrapment_efficiency,
        "reactor_volume_ml": experiment.reactor_volume_ml,
        "probe_volume_ml": experiment.probe_volume_ml,
        "calibration_coefficient": experiment.calibration_coefficient,
        "raw_rows": "\n".join(rows),
    }


def _has_imported_source_results(experiment):
    return (
        ImportedReplicateCalculation.objects.filter(
            measurement__replicate__experiment=experiment
        ).exists()
        or ImportedModelPoint.objects.filter(experiment=experiment).exists()
    )


def _raw_replicates_match(experiment, submitted_replicates):
    stored_replicates = list(
        RawReplicate.objects.filter(experiment=experiment)
        .prefetch_related("measurements")
        .order_by("replicate_number", "pk")
    )
    if len(stored_replicates) != len(submitted_replicates):
        return False
    for stored, submitted in zip(stored_replicates, submitted_replicates):
        if (
            stored.replicate_number != submitted.number
            or stored.sample_mass_g != submitted.sample_mass_g
        ):
            return False
        stored_points = [
            (point.time, point.signal_a)
            for point in stored.measurements.all().order_by("order", "pk")
        ]
        submitted_points = [
            (point.time, point.signal_a) for point in submitted.points
        ]
        if stored_points != submitted_points:
            return False
    return True


def _replace_raw_replicates(experiment, replicates):
    """Replace raw inputs when their values changed, within caller transaction."""

    RawReplicate.objects.filter(experiment=experiment).delete()
    for replicate_input in replicates:
        replicate = RawReplicate.objects.create(
            experiment=experiment,
            replicate_number=replicate_input.number,
            sample_mass_g=replicate_input.sample_mass_g,
        )
        RawMeasurement.objects.bulk_create([
            RawMeasurement(
                replicate=replicate,
                time=point.time,
                signal_a=point.signal_a,
                order=index,
            )
            for index, point in enumerate(replicate_input.points)
        ])


def _render_raw_experiment_form(request, form, **context):
    return render(
        request,
        "kinetics/raw_experiment_form.html",
        {"form": form, **context},
    )


@require_http_methods(["GET", "POST"])
def raw_experiment_new(request):
    """Create an experiment from original replicate-level A readings."""

    example_key = request.GET.get("example", "")
    synthetic_example = SYNTHETIC_EXAMPLES.get(example_key)
    initial = synthetic_example["initial"] if synthetic_example else None
    form = RawExperimentForm(
        request.POST if request.method == "POST" else None,
        initial=initial,
    )
    context = {
        "is_synthetic_example": synthetic_example is not None,
        "synthetic_example": synthetic_example,
        "synthetic_example_key": example_key,
    }
    if request.method == "POST" and form.is_valid():
        values = form.cleaned_data
        parameters = {
            "drug_loading": values["drug_loading"],
            "entrapment_efficiency": values["entrapment_efficiency"],
            "reactor_volume_ml": values["reactor_volume_ml"],
            "probe_volume_ml": values["probe_volume_ml"],
            "calibration_coefficient": values["calibration_coefficient"],
        }
        try:
            experiment = _persist_raw_experiment(
                name=values["name"],
                description=values["description"].strip(),
                source_file=(
                    synthetic_example["source_file"]
                    if synthetic_example else "Ручной ввод исходных данных"
                ),
                parameters=parameters,
                replicates=values["raw_rows"],
                source_format=(
                    Experiment.SourceFormat.SYNTHETIC_RAW
                    if synthetic_example else Experiment.SourceFormat.MANUAL_RAW
                ),
                owner=request.user if request.user.is_authenticated else None,
            )
        except CalculationInputError as exc:
            form.add_error("raw_rows", str(exc))
            return _render_raw_experiment_form(request, form, **context)
        except DatabaseError:
            logger.exception("Не удалось сохранить исходный расчёт")
            messages.error(
                request,
                "Не удалось сохранить эксперимент. База данных не изменилась.",
            )
            return _render_raw_experiment_form(request, form, **context)

        request.session[CURRENT_EXPERIMENT_SESSION_KEY] = experiment.pk
        if synthetic_example:
            messages.success(
                request,
                "Синтетический учебный пример сохранён; это не реальные научные данные.",
            )
        else:
            messages.success(
                request,
                "Исходные измерения сохранены отдельно; Release, среднее и SD рассчитаны.",
            )
        return redirect("experiment_detail", pk=experiment.pk)

    return _render_raw_experiment_form(request, form, **context)


@require_http_methods(["GET", "POST"])
def raw_experiment_edit(request, pk):
    """Edit a saved raw experiment and rebuild outputs from its source rows."""

    experiment = get_object_or_404(editable_experiments_for(request.user), pk=pk)
    if experiment.input_mode != Experiment.InputMode.RAW:
        messages.error(
            request,
            "У этой записи нет исходных повторов для редактирования."
        )
        return redirect("experiment_detail", pk=experiment.pk)

    source_results_locked = _has_imported_source_results(experiment)
    initial = _raw_experiment_initial(experiment)
    form = RawExperimentForm(
        request.POST if request.method == "POST" else None,
        initial=initial,
    )
    if source_results_locked:
        # AllData results belong to the imported raw rows; preserve the linkage
        # and let users edit parameters without silently deleting the C# data.
        form.fields["raw_rows"].disabled = True

    context = {
        "is_edit": True,
        "experiment": experiment,
        "source_results_locked": source_results_locked,
    }
    if request.method == "POST" and form.is_valid():
        values = form.cleaned_data
        parameters = {
            "drug_loading": values["drug_loading"],
            "entrapment_efficiency": values["entrapment_efficiency"],
            "reactor_volume_ml": values["reactor_volume_ml"],
            "probe_volume_ml": values["probe_volume_ml"],
            "calibration_coefficient": values["calibration_coefficient"],
        }
        previous_parameters = tuple(
            getattr(experiment, field)
            for field in parameters
        )
        updated_parameters = tuple(parameters.values())
        parameters_changed = previous_parameters != updated_parameters
        raw_inputs_changed = not _raw_replicates_match(
            experiment, values["raw_rows"]
        )

        try:
            with transaction.atomic():
                experiment.name = values["name"]
                experiment.description = values["description"].strip()
                for field, value in parameters.items():
                    setattr(experiment, field, value)
                if source_results_locked and parameters_changed:
                    experiment.source_results_stale = True
                experiment.save(update_fields=(
                    "name",
                    "description",
                    *parameters.keys(),
                    "source_results_stale",
                    "updated_at",
                ))
                if raw_inputs_changed:
                    if source_results_locked:
                        raise CalculationInputError(
                            "Исходные строки связаны с результатами C# AllData и не могут "
                            "быть заменены отдельно. Создайте новый эксперимент для другого набора данных."
                        )
                    _replace_raw_replicates(experiment, values["raw_rows"])
                recalculate_experiment(experiment)
        except CalculationInputError as exc:
            form.add_error("raw_rows", str(exc))
            return _render_raw_experiment_form(request, form, **context)
        except DatabaseError:
            logger.exception("Не удалось сохранить изменения эксперимента %s", pk)
            form.add_error(
                None,
                "Изменения не сохранены: база данных оставила прежние значения.",
            )
            return _render_raw_experiment_form(request, form, **context)

        request.session[CURRENT_EXPERIMENT_SESSION_KEY] = experiment.pk
        if source_results_locked and (parameters_changed or raw_inputs_changed):
            messages.warning(
                request,
                "Расчёт Django обновлён. Результаты и координаты C# сохранены как исходный снимок и помечены как устаревшие.",
            )
        else:
            messages.success(
                request,
                f"Изменения сохранены; расчёт выполнен заново по версии {CALCULATION_VERSION}.",
            )
        return redirect("experiment_detail", pk=experiment.pk)

    return _render_raw_experiment_form(request, form, **context)


@require_POST
def upload_csharp_workbook(request):
    """Import all raw experiment sheets using the C# source coordinates."""

    try:
        upload_token = uuid.UUID(request.POST.get("upload_token", ""))
    except (ValueError, TypeError, AttributeError):
        messages.error(request, "Форма импорта устарела. Обновите страницу.")
        return redirect("home")

    first_token = uuid.uuid5(upload_token, "worksheet-0")
    owner = request.user if request.user.is_authenticated else None
    existing = Experiment.objects.filter(
        upload_token=first_token,
        owner=owner,
    ).first()
    if existing is not None:
        request.session[CURRENT_EXPERIMENT_SESSION_KEY] = existing.pk
        return redirect("experiment_detail", pk=existing.pk)

    excel_file = request.FILES.get("file")
    if excel_file is None:
        messages.error(request, "Выберите исходную книгу Excel.")
        return redirect("home")
    source_filename = _safe_source_filename(excel_file.name)
    if not source_filename.lower().endswith(".xlsx"):
        messages.error(request, "Поддерживается только формат .xlsx.")
        return redirect("home")

    try:
        is_all_data = _has_all_data_marker(excel_file)
        workbook_sheets = (
            read_csharp_all_data_workbook(excel_file)
            if is_all_data else read_csharp_raw_workbook(excel_file)
        )
        with transaction.atomic():
            saved = []
            for index, sheet in enumerate(workbook_sheets):
                saved.append(_persist_csharp_sheet(
                    sheet=sheet,
                    source_filename=source_filename,
                    upload_token=uuid.uuid5(upload_token, f"worksheet-{index}"),
                    is_all_data=is_all_data,
                    owner=owner,
                ))
    except (RawWorkbookValidationError, CalculationInputError) as exc:
        messages.error(request, str(exc))
        return redirect("home")
    except DatabaseError:
        logger.exception("Не удалось импортировать исходную книгу C#")
        messages.error(request, "Импорт не выполнен; база данных не изменилась.")
        return redirect("home")

    if not saved:
        messages.error(request, "В книге не найдено листов исходных измерений.")
        return redirect("home")
    request.session[CURRENT_EXPERIMENT_SESSION_KEY] = saved[0].pk
    if is_all_data:
        messages.success(
            request,
            f"Импортировано листов AllData: {len(saved)}. Исходные повторы, "
            "результаты C# и координаты моделей сохранены отдельно.",
        )
    else:
        messages.success(
            request,
            f"Импортировано листов экспериментов: {len(saved)}. "
            "Исходные и расчётные значения сохранены отдельно.",
        )
    if len(saved) == 1:
        return redirect("experiment_detail", pk=saved[0].pk)
    return redirect("experiment_history")


@require_POST
def recalculate_raw_experiment(request, pk):
    """Recompute derived values from immutable stored raw readings."""

    experiment = get_object_or_404(editable_experiments_for(request.user), pk=pk)
    try:
        recalculate_experiment(experiment)
    except CalculationInputError as exc:
        messages.error(request, str(exc))
    except DatabaseError:
        logger.exception("Не удалось повторно рассчитать эксперимент %s", pk)
        messages.error(request, "Пересчёт не выполнен; прежние результаты сохранены.")
    else:
        messages.success(
            request,
            f"Расчёт повторён по версии {CALCULATION_VERSION}; исходные данные не изменены.",
        )
    return redirect("experiment_detail", pk=experiment.pk)


@login_required
@require_POST
def delete_experiment(request, pk):
    """Удаляет эксперимент и измерения каскадно только через POST."""

    experiment = get_object_or_404(Experiment, pk=pk, owner=request.user)
    experiment_name = experiment.name
    experiment_id = experiment.pk
    try:
        with transaction.atomic():
            experiment.delete()
    except DatabaseError:
        logger.exception("Не удалось удалить эксперимент")
        messages.error(request, "Не удалось удалить эксперимент. Изменения отменены.")
        return redirect("experiment_history")

    if request.session.get(CURRENT_EXPERIMENT_SESSION_KEY) == experiment_id:
        request.session.pop(CURRENT_EXPERIMENT_SESSION_KEY, None)

    messages.success(request, f"Эксперимент «{experiment_name}» удалён.")
    return redirect("experiment_history")


def _upload_error(request, message):
    messages.error(request, message)
    return redirect("home")


@require_POST
def upload(request):
    """Auto-detect supported .xlsx layouts and save a complete experiment."""

    try:
        upload_uuid = uuid.UUID(request.POST.get("upload_token", ""))
    except (ValueError, TypeError, AttributeError):
        return _upload_error(
            request,
            "Форма загрузки устарела. Обновите страницу и попробуйте ещё раз.",
        )

    upload_token = str(upload_uuid)
    raw_first_token = uuid.uuid5(upload_uuid, "worksheet-0")
    owner = request.user if request.user.is_authenticated else None
    existing = Experiment.objects.filter(
        upload_token__in=(upload_uuid, raw_first_token),
        owner=owner,
    ).first()
    if existing is not None:
        request.session[CURRENT_EXPERIMENT_SESSION_KEY] = existing.pk
        return redirect("experiment_detail", pk=existing.pk)

    excel_file = request.FILES.get("file")
    if excel_file is None:
        return _upload_error(request, "Выберите Excel-файл для загрузки.")

    source_filename = _safe_source_filename(excel_file.name)
    if not source_filename.lower().endswith(".xlsx"):
        return _upload_error(request, "Поддерживается только формат .xlsx.")
    if excel_file.size == 0:
        return _upload_error(request, "Файл пуст. Выберите непустую книгу .xlsx.")
    if excel_file.size > MAX_EXCEL_UPLOAD_SIZE:
        return _upload_error(
            request,
            "Размер файла превышает допустимый предел 25 МБ.",
        )

    try:
        is_precomputed = _has_precomputed_headers(excel_file)
    except WorkbookValidationError as exc:
        return _upload_error(request, str(exc))

    if is_precomputed:
        try:
            points = _read_measurements(excel_file)
        except WorkbookValidationError as exc:
            return _upload_error(request, str(exc))

        experiment_name = PurePosixPath(source_filename).stem.strip()[:200]
        if not experiment_name:
            experiment_name = "Эксперимент"

        try:
            with transaction.atomic():
                experiment = Experiment.objects.create(
                    name=experiment_name,
                    source_file=source_filename,
                    owner=owner,
                    source_format=Experiment.SourceFormat.TIME_RELEASE_SD,
                    upload_token=upload_uuid,
                )

                batch = []
                for order, (time, release, sd) in enumerate(points):
                    batch.append(Measurement(
                        experiment=experiment,
                        time=time,
                        release=release,
                        sd=sd,
                        order=order,
                    ))
                    if len(batch) == 1_000:
                        Measurement.objects.bulk_create(batch, batch_size=1_000)
                        batch.clear()
                if batch:
                    Measurement.objects.bulk_create(batch, batch_size=1_000)
                experiment.save(update_fields=("updated_at",))
        except IntegrityError:
            existing = Experiment.objects.filter(
                upload_token=upload_uuid,
                owner=owner,
            ).first()
            if existing is not None:
                request.session[CURRENT_EXPERIMENT_SESSION_KEY] = existing.pk
                return redirect("experiment_detail", pk=existing.pk)
            logger.exception("Нарушение целостности при сохранении эксперимента")
            return _upload_error(
                request,
                "Не удалось сохранить эксперимент. Проверьте файл и повторите попытку.",
            )
        except DatabaseError:
            logger.exception("Ошибка базы данных при сохранении эксперимента")
            return _upload_error(
                request,
                "Не удалось сохранить эксперимент. Данные не были записаны частично.",
            )
        except Exception:
            logger.exception("Непредвиденная ошибка при сохранении эксперимента")
            return _upload_error(
                request,
                "Не удалось сохранить эксперимент. Попробуйте позже.",
            )

        request.session[CURRENT_EXPERIMENT_SESSION_KEY] = experiment.pk
        messages.success(request, "Эксперимент и измерения сохранены.")
        return redirect("experiment_detail", pk=experiment.pk)

    is_all_data = _has_all_data_marker(excel_file)
    try:
        try:
            workbook_sheets = read_csharp_all_data_workbook(excel_file)
            is_all_data = True
        except RawWorkbookValidationError as all_data_error:
            if is_all_data:
                return _upload_error(
                    request,
                    f"Книга отмечена как AllData, но её таблицы не удалось прочитать: {all_data_error}",
                )
            try:
                workbook_sheets = read_csharp_raw_workbook(excel_file)
            except RawWorkbookValidationError as raw_error:
                return _upload_error(
                    request,
                    "Формат Excel не распознан. Поддерживаются таблица Time/Release/SD, "
                    "исходная книга C# и AllData.xlsx. "
                    f"Проверка AllData: {all_data_error} Проверка исходной книги: {raw_error}",
                )

        saved = []
        with transaction.atomic():
            for index, sheet in enumerate(workbook_sheets):
                saved.append(_persist_csharp_sheet(
                    sheet=sheet,
                    source_filename=source_filename,
                    upload_token=uuid.uuid5(upload_uuid, f"worksheet-{index}"),
                    is_all_data=is_all_data,
                    owner=owner,
                ))
    except (RawWorkbookValidationError, CalculationInputError) as exc:
        return _upload_error(request, str(exc))
    except IntegrityError:
        existing = Experiment.objects.filter(
            upload_token=raw_first_token,
            owner=owner,
        ).first()
        if existing is not None:
            request.session[CURRENT_EXPERIMENT_SESSION_KEY] = existing.pk
            return redirect("experiment_detail", pk=existing.pk)
        logger.exception("Нарушение целостности при импорте книги C#")
        return _upload_error(request, "Импорт не выполнен; база данных не изменилась.")
    except DatabaseError:
        logger.exception("Не удалось импортировать книгу C#")
        return _upload_error(request, "Импорт не выполнен; база данных не изменилась.")

    if not saved:
        return _upload_error(request, "В книге не найдено листов экспериментов.")
    request.session[CURRENT_EXPERIMENT_SESSION_KEY] = saved[0].pk
    if is_all_data:
        messages.success(
            request,
            f"Импортировано листов AllData: {len(saved)}. Исходные повторы, "
            "таблица Zero Order и преобразованные координаты C# сохранены.",
        )
    else:
        messages.success(
            request,
            f"Импортировано листов исходных измерений: {len(saved)}. "
            "Параметры, повторы и исходные сигналы сохранены.",
        )
    if len(saved) == 1:
        return redirect("experiment_detail", pk=saved[0].pk)
    return redirect("experiment_history")
