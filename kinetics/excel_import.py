"""Readers for the raw workbook layout used by the original C# importer."""

from dataclasses import dataclass
import math
import re

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

from .calculations import CalculationParameters, RawPoint, ReplicateInput


MAX_WORKBOOK_BYTES = 25 * 1024 * 1024
MAX_ROWS = 100_000
MAX_COLUMNS = 64


class RawWorkbookValidationError(ValueError):
    """The workbook does not match the source reader's raw-sheet structure."""


@dataclass(frozen=True)
class RawWorkbookSheet:
    name: str
    parameters: CalculationParameters
    replicates: tuple[ReplicateInput, ...]
    model_points: tuple["ImportedCoordinate", ...] = ()


@dataclass(frozen=True)
class ImportedCoordinate:
    model_key: str
    order: int
    x: float | None
    y: float | None
    sd: float | None = None
    status: str = "included"


_ALL_DATA_MODEL_TITLES = {
    "zeroorder": "zero_order",
    "firstorder": "first_order",
    "higuchi": "higuchi",
    "hixsoncrowel": "hixson_crowell",
    "hixsoncrowell": "hixson_crowell",
    "korsmeyerpeppas": "korsmeyer_peppas",
}


def _normalized_cell(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def _optional_coordinate(value, label):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        number = float(str(value).strip().replace(" ", "").replace(",", "."))
    except (TypeError, ValueError, OverflowError):
        return None
    return number if math.isfinite(number) else None


def _parse_all_data_coordinates(sheet):
    """Read the aggregate and transformed model tables written by the C# app."""

    if (sheet.max_row or 0) > MAX_ROWS or (sheet.max_column or 0) > MAX_COLUMNS:
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}» превышает допустимый размер книги AllData."
        )

    zero_header = tuple(
        _normalized_cell(sheet.cell(row=3, column=column).value)
        for column in (9, 10, 11)
    )
    if (
        _normalized_cell(sheet.cell(row=2, column=9).value) != "zeroorder"
        or zero_header[0] not in {"time", "timemin", "timeh", "timesec"}
        or zero_header[1] not in {
            "drugrelease", "drugreleasepercent", "release", "releasepercent"
        }
        or zero_header[2] not in {
            "standartdeviation", "standarddeviation", "sd", "stardartd"
        }
    ):
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}» не соответствует подтверждённой структуре AllData.xlsx "
            "(таблица Zero Order должна начинаться в I2:K3)."
        )

    title_rows = []
    for row in range(2, (sheet.max_row or 0) + 1):
        title = _normalized_cell(sheet.cell(row=row, column=9).value)
        model_key = _ALL_DATA_MODEL_TITLES.get(title)
        if model_key:
            title_rows.append((row, model_key))

    if not title_rows or title_rows[0] != (2, "zero_order"):
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}» не содержит таблицу Zero Order в ячейке I2."
        )

    coordinates = []
    for index, (title_row, model_key) in enumerate(title_rows):
        end_row = (
            title_rows[index + 1][0]
            if index + 1 < len(title_rows)
            else (sheet.max_row or 0) + 1
        )
        first_data_row = title_row + 2
        order = 0
        for row in range(first_data_row, end_row):
            x_value = sheet.cell(row=row, column=9).value
            y_value = sheet.cell(row=row, column=10).value
            sd_value = sheet.cell(row=row, column=11).value if model_key == "zero_order" else None
            if all(value is None or (isinstance(value, str) and not value.strip())
                   for value in (x_value, y_value, sd_value)):
                continue

            order += 1
            x = _optional_coordinate(x_value, "X")
            y = _optional_coordinate(y_value, "Y")
            sd = _optional_coordinate(sd_value, "SD") if model_key == "zero_order" else None
            status = "included"
            if x is None or y is None or (model_key == "zero_order" and sd_value not in (None, "") and sd is None):
                status = "invalid"
            elif model_key == "zero_order" and sd is not None and sd < 0:
                sd = None
                status = "invalid"
            coordinates.append(ImportedCoordinate(
                model_key=model_key,
                order=order - 1,
                x=x,
                y=y,
                sd=sd,
                status=status,
            ))
        if order == 0 and model_key != "zero_order":
            coordinates.append(ImportedCoordinate(
                model_key=model_key,
                order=0,
                x=None,
                y=None,
                status="empty_table",
            ))

    if not any(point.model_key == "zero_order" for point in coordinates):
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}» не содержит строк данных в таблице Zero Order."
        )
    return tuple(coordinates)


def _number(value, label, *, blank_message=None):
    if value is None or (isinstance(value, str) and not value.strip()):
        raise RawWorkbookValidationError(blank_message or f"Не заполнено поле «{label}».")
    try:
        result = float(str(value).strip().replace(" ", "").replace(",", "."))
    except (TypeError, ValueError, OverflowError) as exc:
        raise RawWorkbookValidationError(
            f"В поле «{label}» найдено нечисловое значение."
        ) from exc
    if not math.isfinite(result):
        raise RawWorkbookValidationError(
            f"В поле «{label}» требуется конечное число."
        )
    return result


def _optional_number(value, label):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    return _number(value, label)


def _count_source_replicates(sheet):
    """Mirror FindTryCount: scan row 5 from column O until zero/blank."""

    count = 0
    last_column = min(sheet.max_column or 0, MAX_COLUMNS)
    for column in range(15, last_column + 1):
        value = sheet.cell(row=5, column=column).value
        if value is None or (isinstance(value, str) and not value.strip()):
            break
        try:
            number = _number(value, f"O5 и далее, столбец {column}")
        except RawWorkbookValidationError:
            break
        if number == 0:
            break
        count += 1
    return count


def _parse_sheet(sheet, *, capture_source_results=False):
    if (sheet.max_row or 0) > MAX_ROWS or (sheet.max_column or 0) > MAX_COLUMNS:
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}» превышает предел {MAX_ROWS} строк "
            f"или {MAX_COLUMNS} столбцов."
        )

    name_value = sheet.cell(row=11, column=15).value
    name = str(name_value).strip() if name_value is not None else ""
    if not name:
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}»: имя опыта из ячейки O11 не заполнено."
        )

    parameters = CalculationParameters(
        drug_loading=_number(sheet.cell(row=6, column=15).value, "O6 (DL)"),
        entrapment_efficiency=_number(sheet.cell(row=7, column=15).value, "O7 (EE)"),
        reactor_volume_ml=_number(sheet.cell(row=8, column=15).value, "O8 (объём реактора)"),
        probe_volume_ml=_number(sheet.cell(row=9, column=15).value, "O9 (объём пробы)"),
        calibration_coefficient=_number(sheet.cell(row=10, column=15).value, "O10 (коэффициент калибровки)"),
    )
    expected_replicates = _count_source_replicates(sheet)
    if expected_replicates == 0:
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}»: в строке 5, начиная с O5, "
            "не найдено ни одного номера/значения повтора."
        )

    replicates = []
    start_row = 4
    while start_row <= (sheet.max_row or 0):
        sample_mass_value = sheet.cell(row=start_row, column=1).value
        if sample_mass_value is None or (
            isinstance(sample_mass_value, str) and not sample_mass_value.strip()
        ):
            break
        number = len(replicates) + 1
        sample_mass = _number(
            sample_mass_value,
            f"повтор {number}, столбец A, строка {start_row} (масса образца)",
        )
        points = []
        row = start_row
        while row <= (sheet.max_row or 0):
            time_value = sheet.cell(row=row, column=3).value
            if row > start_row:
                if time_value is None or (
                    isinstance(time_value, str) and not time_value.strip()
                ):
                    break
                time = _number(time_value, f"повтор {number}, время, ячейка C{row}")
                # The source C# reader uses time=0 as its end-of-block sentinel.
                if time == 0:
                    break
            else:
                time = _number(
                    time_value,
                    f"повтор {number}, время, ячейка C{row}",
                )

            signal = _number(
                sheet.cell(row=row, column=4).value,
                f"повтор {number}, сигнал A, ячейка D{row}",
            )
            source_results = (None, None, None)
            if capture_source_results:
                source_results = tuple(
                    _optional_number(
                        sheet.cell(row=row, column=column).value,
                        f"AllData, повтор {number}, {label}, строка {row}, столбец {column}",
                    )
                    for column, label in ((5, "Cabs"), (6, "Mreal"), (7, "Release"))
                )
            points.append(RawPoint(time, signal, *source_results))
            if len(points) > MAX_ROWS:
                raise RawWorkbookValidationError(
                    f"Повтор {number} содержит больше {MAX_ROWS} точек."
                )
            row += 1

        if not points:
            raise RawWorkbookValidationError(
                f"Повтор {number} на листе «{sheet.title}» не содержит точек."
            )
        replicates.append(ReplicateInput(number, sample_mass, tuple(points)))

        # This is the next-block calculation in NewSampleFromWorksheet.
        start_row += len(points) + 5

    if not replicates:
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}»: исходный блок измерений не найден начиная с A4."
        )
    if len(replicates) != expected_replicates:
        raise RawWorkbookValidationError(
            f"Лист «{sheet.title}»: в O5 и далее указано повторов: "
            f"{expected_replicates}; блоков данных начиная с A4 найдено: "
            f"{len(replicates)}."
        )

    return RawWorkbookSheet(name, parameters, tuple(replicates))


def read_csharp_raw_workbook(uploaded_file, *, capture_source_results=False):
    """Parse every worksheet according to WorkWithExcelFile.cs coordinates.

    The C# reader's layout is coordinate-based (O5:O11 metadata and raw
    blocks in A/C/D starting at row 4), not the precomputed Time/Release/SD
    table accepted by the legacy upload route.
    """

    if getattr(uploaded_file, "size", None) == 0:
        raise RawWorkbookValidationError("Файл Excel пуст.")
    if getattr(uploaded_file, "size", 0) > MAX_WORKBOOK_BYTES:
        raise RawWorkbookValidationError("Размер файла превышает допустимый предел 25 МБ.")

    workbook = None
    try:
        uploaded_file.seek(0)
        workbook = load_workbook(uploaded_file, read_only=True, data_only=True)
        if not workbook.worksheets:
            raise RawWorkbookValidationError("В книге нет листов экспериментов.")
        parsed = []
        for sheet in workbook.worksheets:
            try:
                parsed.append(_parse_sheet(
                    sheet,
                    capture_source_results=capture_source_results,
                ))
            except RawWorkbookValidationError as exc:
                if f"Лист «{sheet.title}»" in str(exc):
                    raise
                raise RawWorkbookValidationError(
                    f"Лист «{sheet.title}»: {exc}"
                ) from exc
        return tuple(parsed)
    except RawWorkbookValidationError:
        raise
    except (BadZipFile, InvalidFileException, OSError, ValueError) as exc:
        raise RawWorkbookValidationError(
            "Файл не является корректной книгой .xlsx для исходного импорта C#."
        ) from exc
    finally:
        if workbook is not None:
            workbook.close()


def read_csharp_all_data_workbook(uploaded_file):
    """Read raw replicate blocks and model tables in a C# AllData workbook.

    The source writes the same O5:O11/A-C-D raw layout as its input importer,
    then adds Zero Order (I2:K...) and transformed-coordinate tables in I:J.
    Imported model coordinates are kept separate from Django's recalculated
    measurements so a local recalculation never overwrites the source table.
    """

    raw_sheets = read_csharp_raw_workbook(
        uploaded_file,
        capture_source_results=True,
    )
    workbook = None
    try:
        uploaded_file.seek(0)
        workbook = load_workbook(uploaded_file, read_only=True, data_only=True)
        if len(workbook.worksheets) != len(raw_sheets):
            raise RawWorkbookValidationError(
                "Не удалось сопоставить листы книги AllData с исходными повторами."
            )
        return tuple(
            RawWorkbookSheet(
                name=raw_sheet.name,
                parameters=raw_sheet.parameters,
                replicates=raw_sheet.replicates,
                model_points=_parse_all_data_coordinates(sheet),
            )
            for raw_sheet, sheet in zip(raw_sheets, workbook.worksheets)
        )
    except RawWorkbookValidationError:
        raise
    except (BadZipFile, InvalidFileException, OSError, ValueError) as exc:
        raise RawWorkbookValidationError(
            "Файл не является корректной книгой AllData.xlsx."
        ) from exc
    finally:
        if workbook is not None:
            workbook.close()
