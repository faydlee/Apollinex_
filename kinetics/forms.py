"""Forms for entering original replicate-level release measurements."""

import csv
import math

from django import forms
from django.core.exceptions import ValidationError

from .calculations import (
    CalculationInputError,
    CalculationParameters,
    RawPoint,
    ReplicateInput,
    calculate_dataset,
)


MAX_RAW_ROWS = 100_000


def _parse_number(value, label, line_number):
    normalized = str(value).strip().replace(" ", "").replace(",", ".")
    try:
        result = float(normalized)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValidationError(
            f"Строка {line_number}: в столбце «{label}» требуется число."
        ) from exc
    if not math.isfinite(result):
        raise ValidationError(
            f"Строка {line_number}: в столбце «{label}» требуется конечное число."
        )
    return result


def parse_raw_rows(raw_text):
    """Parse pasted Excel rows: replicate, mass g, time min, signal A."""

    source_lines = str(raw_text or "").splitlines()
    populated_lines = [
        (line_number, line)
        for line_number, line in enumerate(source_lines, start=1)
        if line.strip()
    ]
    if not populated_lines:
        raise ValidationError("Добавьте строки исходных измерений.")
    if len(populated_lines) > MAX_RAW_ROWS:
        raise ValidationError(
            f"Можно ввести не более {MAX_RAW_ROWS} строк измерений."
        )

    first_line = populated_lines[0][1]
    delimiter = "\t" if "\t" in first_line else ";" if ";" in first_line else ","
    if delimiter == ",":
        reader = csv.reader(source_lines, delimiter=",")
        parsed_rows = []
        for row in reader:
            if not any(str(cell).strip() for cell in row):
                continue
            parsed_rows.append((reader.line_num, row))
    else:
        parsed_rows = [
            (line_number, line.split(delimiter))
            for line_number, line in enumerate(source_lines, start=1)
            if line.strip()
        ]

    if parsed_rows and len(parsed_rows[0][1]) >= 1:
        first_cell = "".join(
            character for character in str(parsed_rows[0][1][0]).lower()
            if character.isalnum()
        )
        if first_cell in {"replicate", "repeat", "повтор", "номерповтора"}:
            parsed_rows = parsed_rows[1:]

    grouped = {}
    for line_number, row in parsed_rows:
        if len(row) != 4:
            raise ValidationError(
                f"Строка {line_number}: нужны 4 столбца — повтор, "
                "масса образца (г), время (мин), сигнал A."
            )
        try:
            replicate_number = int(str(row[0]).strip())
        except (TypeError, ValueError, OverflowError) as exc:
            raise ValidationError(
                f"Строка {line_number}: номер повтора должен быть целым числом."
            ) from exc
        if replicate_number < 1:
            raise ValidationError(
                f"Строка {line_number}: номер повтора должен начинаться с 1."
            )

        sample_mass = _parse_number(row[1], "масса образца (г)", line_number)
        time = _parse_number(row[2], "время (мин)", line_number)
        signal = _parse_number(row[3], "сигнал A", line_number)
        if replicate_number not in grouped:
            grouped[replicate_number] = {"mass": sample_mass, "points": []}
        elif grouped[replicate_number]["mass"] != sample_mass:
            raise ValidationError(
                f"Строка {line_number}: для повтора {replicate_number} "
                "масса образца должна быть одинаковой во всех строках."
            )
        grouped[replicate_number]["points"].append(RawPoint(time, signal))

    if not grouped:
        raise ValidationError("Добавьте хотя бы одну строку измерений.")
    replicate_numbers = sorted(grouped)
    if replicate_numbers != list(range(1, len(replicate_numbers) + 1)):
        raise ValidationError("Номера повторов должны идти без пропусков от 1.")

    return tuple(
        ReplicateInput(
            number=number,
            sample_mass_g=grouped[number]["mass"],
            points=tuple(grouped[number]["points"]),
        )
        for number in replicate_numbers
    )


class RawExperimentForm(forms.Form):
    name = forms.CharField(
        label="Название эксперимента",
        max_length=200,
        help_text="Обязательное название, которое будет видно в истории и отчётах.",
        widget=forms.TextInput(attrs={"autocomplete": "off"}),
    )
    description = forms.CharField(
        label="Описание",
        required=False,
        help_text="Необязательные заметки о составе, приборе или условиях измерения.",
        widget=forms.Textarea(attrs={"rows": 3}),
    )
    drug_loading = forms.FloatField(
        label="Drug loading (DL), отношение",
        help_text="Безразмерное отношение из исходной формулы; вводится без умножения на 100.",
        min_value=0,
        widget=forms.NumberInput(attrs={"step": "any"}),
    )
    entrapment_efficiency = forms.FloatField(
        label="EE, прямой множитель формулы C#",
        help_text="Исходное значение из O7 без преобразования: формула использует введённое число напрямую и не делит его на 100. Значения 0,8 и 80 отличаются в расчёте в 100 раз; ожидаемая шкала исходником не подтверждена.",
        min_value=0,
        widget=forms.NumberInput(attrs={"step": "any"}),
    )
    reactor_volume_ml = forms.FloatField(
        label="Объём реактора, мл",
        help_text="Должен быть больше нуля.",
        min_value=0.0000001,
        widget=forms.NumberInput(attrs={"step": "any", "min": "0.0000001"}),
    )
    probe_volume_ml = forms.FloatField(
        label="Объём отбираемой пробы, мл",
        help_text="Объём, учитываемый для ранее отобранных проб; укажите 0, если пробы не отбирались.",
        min_value=0,
        widget=forms.NumberInput(attrs={"step": "any"}),
    )
    calibration_coefficient = forms.FloatField(
        label="Калибровочный коэффициент k (A × k → Cabs, мг/мл)",
        help_text="Неотрицательный коэффициент калибровки прибора для сигнала A.",
        min_value=0,
        widget=forms.NumberInput(attrs={"step": "any"}),
    )
    raw_rows = forms.CharField(
        label="Повтор, масса (г), время (мин), сигнал A",
        help_text="Вставьте столбцы через Tab или точку с запятой: номер повтора, масса образца в г, время в мин, сигнал A в единицах прибора. Порядок строк сохраняется; у повторов должно быть одинаковое число точек. Десятичная запятая поддерживается.",
        widget=forms.Textarea(attrs={"rows": 14, "spellcheck": "false"}),
    )

    def clean_raw_rows(self):
        return parse_raw_rows(self.cleaned_data["raw_rows"])

    def clean_name(self):
        name = self.cleaned_data["name"].strip()
        if not name:
            raise ValidationError("Укажите название эксперимента.")
        return name

    def clean(self):
        cleaned_data = super().clean()
        if any(
            field not in cleaned_data
            for field in (
                "drug_loading",
                "entrapment_efficiency",
                "reactor_volume_ml",
                "probe_volume_ml",
                "calibration_coefficient",
                "raw_rows",
            )
        ):
            return cleaned_data

        parameters = CalculationParameters(
            drug_loading=cleaned_data["drug_loading"],
            entrapment_efficiency=cleaned_data["entrapment_efficiency"],
            reactor_volume_ml=cleaned_data["reactor_volume_ml"],
            probe_volume_ml=cleaned_data["probe_volume_ml"],
            calibration_coefficient=cleaned_data["calibration_coefficient"],
        )
        try:
            calculate_dataset(cleaned_data["raw_rows"], parameters)
        except CalculationInputError as exc:
            self.add_error("raw_rows", str(exc))
        return cleaned_data
