import base64
import math
import struct
from datetime import datetime, timezone as datetime_timezone
from io import BytesIO, StringIO
from unittest.mock import patch

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.core.management import call_command
from django.core.management.base import CommandError
from django.db import DatabaseError, connection
from django.db.migrations.executor import MigrationExecutor
from django.test import Client, TestCase, TransactionTestCase, override_settings
from django.urls import resolve
from openpyxl import Workbook, load_workbook as openpyxl_load_workbook

from .calculations import (
    CalculationInputError,
    CalculationParameters,
    RawPoint,
    ReplicateInput,
    calculate_dataset,
    calculate_replicate,
    fit_linear_trendline,
    summarize_source_model_transforms,
    transform_source_model,
)
from .forms import parse_raw_rows
from .examples import SYNTHETIC_EXAMPLES
from .excel_import import (
    read_csharp_all_data_workbook,
    read_csharp_raw_workbook,
)
from .models import (
    Experiment,
    ImportedReplicateCalculation,
    ImportedModelPoint,
    Measurement,
    RawMeasurement,
    RawReplicate,
    ReplicateCalculation,
)
from .views import (
    MAX_PLOT_POINTS,
    _build_comparison_plot,
    _build_plot,
    _build_model_plot,
    _calculate_ttp_qtp,
    _fit_kinetic_models,
    _first_order_model_fit,
    _plot_points,
)


def workbook_upload(rows, headers=("Time", "Release"), filename="experiment.xlsx"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(list(headers))
    for row in rows:
        sheet.append(list(row))

    content = BytesIO()
    workbook.save(content)
    workbook.close()
    return SimpleUploadedFile(
        filename,
        content.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


def csharp_raw_workbook_upload(sheet_count=2, filename="source-workbook.xlsx"):
    """Build a test workbook from WorkWithExcelFile.cs's documented cells."""
    workbook = Workbook()
    first_sheet = workbook.active
    sheets = [first_sheet]
    for index in range(1, sheet_count):
        sheets.append(workbook.create_sheet(f"Source {index + 1}"))

    for index, sheet in enumerate(sheets, start=1):
        sheet.title = f"Source {index}"
        sheet.cell(6, 15, 0.25)
        sheet.cell(7, 15, 0.8)
        sheet.cell(8, 15, 100)
        sheet.cell(9, 15, 2)
        sheet.cell(10, 15, 0.5)
        sheet.cell(11, 15, f"C# sheet {index}")

        replicate_data = [
            (0.25, [(0, 1), (10, 2)]),
            (0.25, [(0, 1.2), (9, 1.8)]),
        ]
        if index == 2:
            replicate_data = [(0.3, [(0, 0.5), (5, 1)])]
        for replicate_index, _replicate in enumerate(replicate_data):
            sheet.cell(5, 15 + replicate_index, 0.25)

        start_row = 4
        for sample_mass, points in replicate_data:
            sheet.cell(start_row, 1, sample_mass)
            for point_index, (time, signal) in enumerate(points):
                sheet.cell(start_row + point_index, 3, time)
                sheet.cell(start_row + point_index, 4, signal)
            start_row += len(points) + 5

    content = BytesIO()
    workbook.save(content)
    workbook.close()
    return SimpleUploadedFile(
        filename,
        content.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


def csharp_all_data_workbook_upload(filename="AllData.xlsx"):
    """Build a compact fixture from the source writer's confirmed cell layout."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "AllData sample"
    sheet.cell(5, 15, 0.25)
    sheet.cell(5, 16, 0.30)
    for row, value in (
        (6, 0.25),
        (7, 0.8),
        (8, 100),
        (9, 2),
        (10, 0.5),
        (11, "AllData test experiment"),
    ):
        sheet.cell(row, 15, value)

    sheet.cell(4, 1, 0.25)
    sheet.cell(4, 3, 0)
    sheet.cell(4, 4, 1)
    sheet.cell(4, 5, 0.5)
    sheet.cell(4, 6, 50)
    sheet.cell(4, 7, 125)
    sheet.cell(5, 3, 1)
    sheet.cell(5, 4, 2)
    sheet.cell(5, 5, 1)
    sheet.cell(5, 6, 101)
    sheet.cell(5, 7, 252.5)
    sheet.cell(11, 1, 0.30)
    sheet.cell(11, 3, 0)
    sheet.cell(11, 4, 1.2)
    sheet.cell(11, 5, 0.6)
    sheet.cell(11, 6, 60)
    sheet.cell(11, 7, 125)
    sheet.cell(12, 3, 2)
    sheet.cell(12, 4, 1.8)
    sheet.cell(12, 5, 0.9)
    sheet.cell(12, 6, 93)
    sheet.cell(12, 7, 193.75)

    sheet.merge_cells("I2:K2")
    sheet.cell(2, 9, "Zero Order")
    sheet.cell(3, 9, "Time, min")
    sheet.cell(3, 10, "Drug release, %")
    sheet.cell(3, 11, "Standart deviation")
    for row, time, release, sd in ((4, 0, 10, 1.5), (5, 1, 25, None)):
        sheet.cell(row, 9, time)
        sheet.cell(row, 10, release)
        sheet.cell(row, 11, sd)

    tables = (
        ("First Order", ((0, math.log(90)), (1, math.log(75)))),
        ("Higuchi", ((0, 10), (1, 25))),
        ("Hixson Crowel", ((0, 0), (1, 0.1))),
        ("Korsmeyer Peppas", ((0, 0), (0, math.log(25)))),
    )
    title_row = 9
    for title, points in tables:
        sheet.merge_cells(start_row=title_row, start_column=9, end_row=title_row, end_column=10)
        sheet.cell(title_row, 9, title)
        sheet.cell(title_row + 1, 9, "X")
        sheet.cell(title_row + 1, 10, "Y")
        for offset, (x_value, y_value) in enumerate(points, start=2):
            sheet.cell(title_row + offset, 9, x_value)
            sheet.cell(title_row + offset, 10, y_value)
        title_row += len(points) + 5

    content = BytesIO()
    workbook.save(content)
    workbook.close()
    return SimpleUploadedFile(
        filename,
        content.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


class RawCalculationTests(TestCase):
    def test_csharp_formula_and_population_sd_are_preserved(self):
        parameters = CalculationParameters(
            drug_loading=0.25,
            entrapment_efficiency=0.8,
            reactor_volume_ml=100,
            probe_volume_ml=2,
            calibration_coefficient=0.5,
        )
        replicate_1 = ReplicateInput(
            number=1,
            sample_mass_g=0.25,
            points=(RawPoint(0, 1), RawPoint(10, 2)),
        )
        replicate_2 = ReplicateInput(
            number=2,
            sample_mass_g=0.25,
            points=(RawPoint(0, 1.2), RawPoint(9, 1.8)),
        )

        calculated_1 = calculate_replicate(replicate_1, parameters)
        self.assertEqual(calculated_1[0].c_abs, 0.5)
        self.assertEqual(calculated_1[0].m_real, 50)
        self.assertEqual(calculated_1[0].release, 125)
        self.assertEqual(calculated_1[1].m_real, 101)
        release_1_second = 252.50001525878906
        release_2_first = 150.00001525878906
        self.assertEqual(calculated_1[1].release, release_1_second)

        result = calculate_dataset((replicate_2, replicate_1), parameters)
        aggregate = result.aggregate_points
        self.assertEqual(
            [(point.time, point.release) for point in aggregate],
            [(0, 137.5), (10, 240.25)],
        )
        expected_sd_1 = math.sqrt(
            ((125 - 137.5) ** 2 + (release_2_first - 137.5) ** 2) / 2
        )
        expected_sd_2 = math.sqrt(
            ((release_1_second - 240.25) ** 2 + (228 - 240.25) ** 2) / 2
        )
        self.assertAlmostEqual(aggregate[0].sd, expected_sd_1, delta=1e-10)
        self.assertAlmostEqual(aggregate[1].sd, expected_sd_2, delta=1e-10)

    def test_raw_signal_calculation_rounds_like_csharp_float_fields(self):
        to_float32 = lambda value: struct.unpack("!f", struct.pack("!f", value))[0]
        signal = 0.123456789
        calibration = 0.987654321
        expected_c_abs = to_float32(to_float32(signal) * to_float32(calibration))
        result = calculate_replicate(
            ReplicateInput(1, 0.25, (RawPoint(0, signal),)),
            CalculationParameters(0.25, 0.8, 100, 0, calibration),
        )

        self.assertEqual(result[0].c_abs, expected_c_abs)

    def test_empty_replicate_and_different_point_counts_are_rejected(self):
        parameters = CalculationParameters(0.2, 0.9, 100, 1, 0.5)
        with self.assertRaisesRegex(CalculationInputError, "нет исходных измерений"):
            calculate_replicate(ReplicateInput(1, 0.1, ()), parameters)

        with self.assertRaisesRegex(CalculationInputError, "одинаковое количество точек"):
            calculate_dataset((
                ReplicateInput(1, 0.1, (RawPoint(0, 1),)),
                ReplicateInput(2, 0.1, (RawPoint(0, 1), RawPoint(1, 2))),
            ), parameters)

    def test_zero_denominator_special_case_matches_source(self):
        result = calculate_replicate(
            ReplicateInput(1, 0.1, (RawPoint(0, 0),)),
            CalculationParameters(0, 0.8, 100, 1, 0.5),
        )
        self.assertEqual(result[0].release, 0)
        with self.assertRaisesRegex(CalculationInputError, "масса препарата равна нулю"):
            calculate_replicate(
                ReplicateInput(1, 0.1, (RawPoint(0, 1),)),
                CalculationParameters(0, 0.8, 100, 1, 0.5),
            )


class RawExperimentWorkflowTests(TestCase):
    def raw_form_data(self, **overrides):
        data = {
            "name": "Исходный опыт",
            "description": "Сигналы двух повторов",
            "drug_loading": "0.25",
            "entrapment_efficiency": "0.8",
            "reactor_volume_ml": "100",
            "probe_volume_ml": "2",
            "calibration_coefficient": "0.5",
            "raw_rows": (
                "1\t0,25\t0\t1\n"
                "1\t0,25\t10\t2\n"
                "2\t0,25\t0\t1.2\n"
                "2\t0,25\t9\t1.8"
            ),
        }
        data.update(overrides)
        return data

    def test_raw_inputs_and_per_replicate_results_are_saved_separately(self):
        response = self.client.post(
            "/experiments/new/raw/",
            self.raw_form_data(),
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Исходные параметры и повторные измерения")
        self.assertContains(response, "Исходные измерения сохранены отдельно")
        experiment = Experiment.objects.get()
        self.assertEqual(experiment.input_mode, Experiment.InputMode.RAW)
        self.assertEqual(experiment.calculation_version, "source-float32-1")
        self.assertEqual(experiment.raw_replicates.count(), 2)
        self.assertEqual(RawMeasurement.objects.count(), 4)
        self.assertEqual(ReplicateCalculation.objects.count(), 4)
        self.assertEqual(
            list(RawMeasurement.objects.order_by("replicate__replicate_number", "order")
                 .values_list("time", "signal_a")),
            [(0, 1), (10, 2), (0, 1.2), (9, 1.8)],
        )
        output = list(experiment.measurements.order_by("order").values_list(
            "time", "release", "sd"
        ))
        self.assertEqual([(time, release) for time, release, _sd in output], [
            (0, 137.5),
            (10, 240.25),
        ])
        expected_sd_1 = math.sqrt(
            ((125 - 137.5) ** 2 + (150.00001525878906 - 137.5) ** 2) / 2
        )
        expected_sd_2 = math.sqrt(
            ((252.50001525878906 - 240.25) ** 2 + (228 - 240.25) ** 2) / 2
        )
        self.assertAlmostEqual(output[0][2], expected_sd_1, delta=1e-10)
        self.assertAlmostEqual(output[1][2], expected_sd_2, delta=1e-10)

    def test_authenticated_manual_work_is_owned_by_the_creator(self):
        user = get_user_model().objects.create_user(
            username="manual-work-owner",
            password="test-password",
        )
        self.client.force_login(user)

        response = self.client.post(
            "/experiments/new/raw/",
            self.raw_form_data(),
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        experiment = Experiment.objects.get()
        self.assertEqual(experiment.owner, user)
        self.assertContains(response, "Удалить")

    def test_recalculate_preserves_raw_rows_and_replaces_derived_values(self):
        self.client.post("/experiments/new/raw/", self.raw_form_data())
        experiment = Experiment.objects.get()
        raw_before = list(RawMeasurement.objects.order_by("pk").values_list(
            "pk", "time", "signal_a", "order"
        ))

        response = self.client.post(
            f"/experiments/{experiment.pk}/recalculate/",
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Расчёт повторён")
        self.assertEqual(
            list(RawMeasurement.objects.order_by("pk").values_list(
                "pk", "time", "signal_a", "order"
            )),
            raw_before,
        )
        self.assertEqual(RawMeasurement.objects.count(), 4)
        self.assertEqual(ReplicateCalculation.objects.count(), 4)
        self.assertEqual(experiment.measurements.count(), 2)

    def test_editing_raw_rows_and_parameters_recalculates_from_sources_once(self):
        self.client.post("/experiments/new/raw/", self.raw_form_data())
        experiment = Experiment.objects.get()
        replacement_rows = (
            "1\t0.25\t10\t2\n"
            "1\t0.25\t0\t1\n"
            "2\t0.25\t9\t1.8\n"
            "2\t0.25\t0\t1.2"
        )
        edit_page = self.client.get(f"/experiments/{experiment.pk}/edit/")
        self.assertContains(edit_page, "заменить текущие результаты Django")

        response = self.client.post(
            f"/experiments/{experiment.pk}/edit/",
            self.raw_form_data(
                name="Отредактированный опыт",
                entrapment_efficiency="1.6",
                raw_rows=replacement_rows,
            ),
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Изменения сохранены")
        experiment.refresh_from_db()
        self.assertEqual(experiment.name, "Отредактированный опыт")
        self.assertEqual(experiment.entrapment_efficiency, 1.6)
        self.assertFalse(experiment.source_results_stale)
        self.assertEqual(
            list(RawMeasurement.objects.order_by(
                "replicate__replicate_number", "order"
            ).values_list("time", "signal_a")),
            [(10, 2), (0, 1), (9, 1.8), (0, 1.2)],
        )

        expected = calculate_dataset(
            parse_raw_rows(replacement_rows),
            CalculationParameters(
                drug_loading=0.25,
                entrapment_efficiency=1.6,
                reactor_volume_ml=100,
                probe_volume_ml=2,
                calibration_coefficient=0.5,
            ),
        )
        actual = list(experiment.measurements.order_by("order").values_list(
            "time", "release", "sd"
        ))
        self.assertEqual(actual, [
            (point.time, point.release, point.sd)
            for point in expected.aggregate_points
        ])
        self.assertEqual(ReplicateCalculation.objects.count(), 4)

    def test_editing_imported_all_data_preserves_source_rows_and_marks_snapshot_stale(self):
        self.client.post("/experiments/new/raw/", self.raw_form_data())
        experiment = Experiment.objects.get()
        raw_before = list(RawMeasurement.objects.order_by("pk").values_list(
            "pk", "time", "signal_a", "order"
        ))
        source_row = RawMeasurement.objects.order_by("pk").first()
        ImportedReplicateCalculation.objects.create(
            measurement=source_row,
            c_abs=12.5,
            m_real=25,
            release=50,
        )
        ImportedModelPoint.objects.create(
            experiment=experiment,
            model_key="zero_order",
            order=1,
            x=10,
            y=50,
        )

        edit_page = self.client.get(f"/experiments/{experiment.pk}/edit/")
        self.assertEqual(edit_page.status_code, 200)
        self.assertTrue(edit_page.context["form"].fields["raw_rows"].disabled)
        response = self.client.post(
            f"/experiments/{experiment.pk}/edit/",
            self.raw_form_data(
                entrapment_efficiency="1.6",
                raw_rows="1\t0.25\t0\t999",
            ),
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "сохранены как исходный снимок")
        experiment.refresh_from_db()
        self.assertTrue(experiment.source_results_stale)
        self.assertEqual(
            list(RawMeasurement.objects.order_by("pk").values_list(
                "pk", "time", "signal_a", "order"
            )),
            raw_before,
        )
        imported = ImportedReplicateCalculation.objects.get()
        self.assertEqual(
            (imported.c_abs, imported.m_real, imported.release),
            (12.5, 25, 50),
        )
        self.assertEqual(ImportedModelPoint.objects.get().y, 50)
        self.assertContains(response, "не соответствуют текущему пересчёту Django")

    def test_excel_export_contains_parameters_means_repeats_and_source_results(self):
        self.client.post("/experiments/new/raw/", self.raw_form_data())
        experiment = Experiment.objects.get()
        first_raw = RawMeasurement.objects.order_by("pk").first()
        ImportedReplicateCalculation.objects.create(
            measurement=first_raw,
            c_abs=12.5,
            m_real=25,
            release=50,
        )
        ImportedModelPoint.objects.create(
            experiment=experiment,
            model_key="zero_order",
            order=1,
            x=10,
            y=50,
            sd=2.5,
        )

        response = self.client.get(
            f"/experiments/{experiment.pk}/export.xlsx"
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response["Content-Type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        workbook = openpyxl_load_workbook(BytesIO(response.content), data_only=True)
        self.assertEqual(
            workbook.sheetnames,
            [
                "Experiment", "Input parameters", "Mean Release SD",
                "Replicate readings", "Django replicate calculations",
                "C# imported results", "CSharp coordinates", "TTP QTP",
                "Django model fits", "Charts",
            ],
        )
        summary = dict(workbook["Experiment"].iter_rows(
            min_row=2, values_only=True
        ))
        self.assertEqual(summary["Calculation method/version"], "source-float32-1")
        self.assertEqual(summary["Synthetic example"], "no")
        input_parameters = {
            row[0]: row[1]
            for row in workbook["Input parameters"].iter_rows(
                min_row=2, values_only=True
            )
        }
        self.assertEqual(input_parameters["Entrapment efficiency (EE)"], 0.8)
        self.assertEqual(input_parameters["Number of replicates"], 2)
        self.assertEqual(summary["Source format"], "Ручной ввод исходных измерений")
        means = workbook["Mean Release SD"]
        self.assertEqual(means.cell(2, 3).value, 137.5)
        self.assertIsNotNone(means.cell(2, 4).value)
        raw = workbook["Replicate readings"]
        self.assertEqual(raw.cell(2, 5).value, 1)
        self.assertEqual(workbook["Django replicate calculations"].cell(2, 3).value, 0.5)
        self.assertEqual(workbook["C# imported results"].cell(2, 5).value, 12.5)
        self.assertEqual(workbook["CSharp coordinates"].cell(2, 4).value, 50)
        analysis_summary = {
            row[0]: row[1]
            for row in workbook["TTP QTP"].iter_rows(
                min_row=2, min_col=1, max_col=2, values_only=True
            )
        }
        self.assertEqual(analysis_summary["TTP"], 10)
        self.assertAlmostEqual(analysis_summary["QTP"], 240.25)
        self.assertIn("Strict crossing", analysis_summary["Status"])
        self.assertEqual(workbook["Django model fits"].cell(2, 1).value, "Zero Order")
        self.assertGreaterEqual(len(workbook["Charts"]._images), 1)
        workbook.close()

    def test_synthetic_examples_are_prefilled_labeled_and_match_reference_results(self):
        home = self.client.get("/")
        self.assertContains(home, "синтетический учебный пример A")
        self.assertContains(home, "synthetic-b")

        saved_experiments = []
        for key, example in SYNTHETIC_EXAMPLES.items():
            form_page = self.client.get(
                "/experiments/new/raw/",
                {"example": key},
            )
            self.assertEqual(form_page.status_code, 200)
            self.assertEqual(
                form_page.context["form"]["raw_rows"].value(),
                example["initial"]["raw_rows"],
            )
            self.assertContains(form_page, "Не реальные научные данные")

            response = self.client.post(
                f"/experiments/new/raw/?example={key}",
                example["initial"],
                follow=True,
            )
            self.assertEqual(response.status_code, 200)
            self.assertContains(response, "Синтетический учебный пример сохранён")
            experiment = Experiment.objects.get(source_file=example["source_file"])
            self.assertEqual(
                experiment.source_format,
                Experiment.SourceFormat.SYNTHETIC_RAW,
            )
            saved_experiments.append(experiment)
            actual = list(experiment.measurements.order_by("order").values_list(
                "time", "release", "sd"
            ))
            self.assertEqual(len(actual), len(example["expected"]))
            for actual_row, expected in zip(actual, example["expected"]):
                self.assertEqual(actual_row[0], expected["time"])
                self.assertAlmostEqual(actual_row[1], expected["release"], delta=1e-5)
                self.assertAlmostEqual(actual_row[2], expected["sd"], delta=1e-5)
            self.assertEqual(response.context["ttp"], example["ttp"])
            self.assertAlmostEqual(response.context["qtp"], example["qtp"], delta=1e-5)
            self.assertContains(response, "Скачать Excel")
            self.assertContains(response, "Изменить исходные данные")

        comparison = self.client.get(
            "/comparison/",
            {"experiment_ids": [item.pk for item in saved_experiments]},
        )
        self.assertEqual(comparison.status_code, 200)
        self.assertContains(
            comparison,
            "Сравнение выбранных экспериментов по фактическому времени",
        )
        self.assertEqual(
            {item.name for item in comparison.context["selected_experiments"]},
            {example["initial"]["name"] for example in SYNTHETIC_EXAMPLES.values()},
        )
        self.assertEqual(
            {dataset["method"] for dataset in comparison.context["datasets"]},
            {f"Django {experiment.calculation_version}" for experiment in saved_experiments},
        )
        self.assertEqual(comparison.context["compatibility_warnings"], [])
        for experiment in saved_experiments:
            self.assertContains(comparison, f"/experiments/{experiment.pk}/")
            detail = self.client.get(f"/experiments/{experiment.pk}/")
            self.assertContains(detail, "не являются реальными научными результатами")

    def test_method_documentation_explains_confirmed_formulas_and_limits(self):
        response = self.client.get("/method/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Mreal_i_mg")
        self.assertContains(response, "популяционное стандартное отклонение")
        self.assertContains(response, "строгое")
        self.assertContains(response, "не подтверждена")
        self.assertContains(response, "AllData.xlsx")

    def test_mismatched_repeat_lengths_are_rejected_without_persisting(self):
        response = self.client.post(
            "/experiments/new/raw/",
            self.raw_form_data(raw_rows="1\t0.25\t0\t1\n2\t0.25\t0\t1\n2\t0.25\t1\t2"),
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "одинаковое количество точек")
        self.assertEqual(Experiment.objects.count(), 0)
        self.assertEqual(RawMeasurement.objects.count(), 0)

    def test_semicolon_rows_support_decimal_comma_and_keep_input_order(self):
        rows = parse_raw_rows(
            "Повтор;Масса;Время;A\n"
            "1;0,25;10,5;1,2\n"
            "1;0,25;0;0,8"
        )

        self.assertEqual(
            [(point.time, point.signal_a) for point in rows[0].points],
            [(10.5, 1.2), (0, 0.8)],
        )

    def test_report_includes_raw_inputs_parameters_and_calculation_version(self):
        self.client.post("/experiments/new/raw/", self.raw_form_data())
        experiment = Experiment.objects.get()

        response = self.client.get(
            "/reports/",
            {"experiment_id": experiment.pk},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Исходные параметры и значения расчёта")
        self.assertContains(response, "source-float32-1")
        self.assertContains(response, "Сигнал A")
        self.assertEqual(len(response.context["report_raw_measurements"]), 4)

    def test_manual_form_explains_ee_direct_multiplier_scale(self):
        response = self.client.get("/experiments/new/raw/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "0,8 и 80 отличаются в расчёте в 100 раз")
        self.assertContains(response, "значение из O7 без преобразования")


class CSharpRawWorkbookImportTests(TestCase):
    def test_reader_uses_source_confirmed_cells_and_reads_all_sheets(self):
        sheets = read_csharp_raw_workbook(csharp_raw_workbook_upload())

        self.assertEqual([sheet.name for sheet in sheets], ["C# sheet 1", "C# sheet 2"])
        self.assertEqual([len(sheet.replicates) for sheet in sheets], [2, 1])
        self.assertEqual(
            [(point.time, point.signal_a) for point in sheets[0].replicates[0].points],
            [(0, 1), (10, 2)],
        )
        self.assertEqual(sheets[0].parameters.entrapment_efficiency, 0.8)

    def test_multi_sheet_import_persists_raw_and_calculated_rows(self):
        response = self.client.post(
            "/upload/csharp-raw/",
            {
                "file": csharp_raw_workbook_upload(),
                "upload_token": "00000000-0000-0000-0000-000000000001",
            },
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Импортировано листов экспериментов: 2")
        self.assertEqual(Experiment.objects.count(), 2)
        self.assertEqual(RawReplicate.objects.count(), 3)
        self.assertEqual(RawMeasurement.objects.count(), 6)
        self.assertEqual(ReplicateCalculation.objects.count(), 6)
        self.assertEqual(Measurement.objects.count(), 4)
        self.assertEqual(
            set(Experiment.objects.values_list("input_mode", flat=True)),
            {Experiment.InputMode.RAW},
        )
        self.assertEqual(
            set(Experiment.objects.values_list("source_format", flat=True)),
            {Experiment.SourceFormat.CSHARP_RAW},
        )

    def test_replayed_source_workbook_upload_does_not_duplicate_experiments(self):
        token = "00000000-0000-0000-0000-000000000002"
        for _ in range(2):
            response = self.client.post(
                "/upload/csharp-raw/",
                {"file": csharp_raw_workbook_upload(), "upload_token": token},
                follow=False,
            )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(Experiment.objects.count(), 2)
        self.assertEqual(RawMeasurement.objects.count(), 6)

    def test_precomputed_table_is_not_mistaken_for_csharp_raw_book(self):
        precomputed = workbook_upload([(0, 10), (1, 20)])
        response = self.client.post(
            "/upload/csharp-raw/",
            {
                "file": precomputed,
                "upload_token": "00000000-0000-0000-0000-000000000003",
            },
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(Experiment.objects.count(), 0)
        self.assertContains(response, "O11")


class AllDataWorkbookImportTests(TestCase):
    def test_reader_preserves_all_data_zero_order_sd_and_model_coordinates(self):
        sheets = read_csharp_all_data_workbook(csharp_all_data_workbook_upload())

        self.assertEqual(len(sheets), 1)
        sheet = sheets[0]
        self.assertEqual(sheet.name, "AllData test experiment")
        self.assertEqual(len(sheet.replicates), 2)
        self.assertEqual(sheet.parameters.entrapment_efficiency, 0.8)
        self.assertEqual(
            (
                sheet.replicates[0].points[0].source_c_abs,
                sheet.replicates[0].points[0].source_m_real,
                sheet.replicates[0].points[0].source_release,
            ),
            (0.5, 50, 125),
        )
        zero_order = [point for point in sheet.model_points if point.model_key == "zero_order"]
        self.assertEqual([(point.x, point.y) for point in zero_order], [(0, 10), (1, 25)])
        self.assertEqual(zero_order[0].sd, 1.5)
        self.assertIsNone(zero_order[1].sd)
        self.assertEqual(len(sheet.model_points), 10)

    def test_unified_upload_saves_raw_inputs_and_exact_all_data_coordinates(self):
        token = "00000000-0000-0000-0000-000000000101"
        user = get_user_model().objects.create_user(
            username="alldata-owner",
            password="test-password",
        )
        self.client.force_login(user)
        response = self.client.post(
            "/upload/",
            {"file": csharp_all_data_workbook_upload(), "upload_token": token},
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Импортировано листов AllData: 1")
        experiment = Experiment.objects.get()
        self.assertEqual(experiment.owner, user)
        self.assertEqual(experiment.input_mode, Experiment.InputMode.RAW)
        self.assertEqual(
            experiment.source_format,
            Experiment.SourceFormat.CSHARP_ALLDATA,
        )
        self.assertEqual(experiment.drug_loading, 0.25)
        self.assertEqual(RawReplicate.objects.filter(experiment=experiment).count(), 2)
        self.assertEqual(RawMeasurement.objects.filter(replicate__experiment=experiment).count(), 4)
        self.assertEqual(
            ImportedReplicateCalculation.objects.filter(
                measurement__replicate__experiment=experiment
            ).count(),
            4,
        )
        self.assertEqual(Measurement.objects.filter(experiment=experiment).count(), 2)
        self.assertEqual(ImportedModelPoint.objects.filter(experiment=experiment).count(), 10)
        self.assertIsNone(
            ImportedModelPoint.objects.get(
                experiment=experiment,
                model_key="zero_order",
                order=1,
            ).sd
        )

        model_response = self.client.get("/models/", {
            "experiment_id": experiment.pk,
            "source_model_key": "first_order",
        })
        self.assertEqual(model_response.status_code, 200)
        self.assertEqual(model_response.context["source_coordinate_source"], "imported")
        self.assertEqual(model_response.context["source_coordinate_count"], 2)
        self.assertContains(model_response, "координаты модели")
        self.assertContains(
            model_response,
            f"/experiments/{experiment.pk}/imported-transforms/first_order.csv",
        )

        report_response = self.client.get(
            "/reports/",
            {"experiment_id": experiment.pk},
        )
        self.assertEqual(report_response.status_code, 200)
        self.assertEqual(
            report_response.context["report_raw_measurements"][0]
            .source_calculation.c_abs,
            0.5,
        )
        self.assertContains(report_response, "C# AllData Release")
        self.assertContains(report_response, "координаты модели")

        csv_response = self.client.get(
            f"/experiments/{experiment.pk}/imported-transforms/first_order.csv"
        )
        self.assertEqual(csv_response.status_code, 200)
        self.assertIn("transformed_x", csv_response.content.decode("utf-8-sig"))

    def test_recalculation_does_not_change_raw_or_imported_values(self):
        self.client.post(
            "/upload/",
            {
                "file": csharp_all_data_workbook_upload(),
                "upload_token": "00000000-0000-0000-0000-000000000102",
            },
        )
        experiment = Experiment.objects.get()
        raw_before = list(
            RawMeasurement.objects.filter(replicate__experiment=experiment)
            .order_by("replicate_id", "order")
            .values_list("replicate__replicate_number", "time", "signal_a")
        )
        imported_before = list(
            ImportedModelPoint.objects.filter(experiment=experiment)
            .order_by("model_key", "order")
            .values_list("model_key", "order", "x", "y", "sd", "status")
        )
        imported_calculations_before = list(
            ImportedReplicateCalculation.objects.filter(
                measurement__replicate__experiment=experiment
            ).order_by("measurement__replicate__replicate_number", "measurement__order")
            .values_list(
                "measurement__replicate__replicate_number",
                "measurement__order",
                "c_abs",
                "m_real",
                "release",
            )
        )
        aggregate_before = list(
            Measurement.objects.filter(experiment=experiment)
            .order_by("order")
            .values_list("time", "release", "sd")
        )

        for _ in range(2):
            response = self.client.post(f"/experiments/{experiment.pk}/recalculate/")
            self.assertEqual(response.status_code, 302)

        raw_after = list(
            RawMeasurement.objects.filter(replicate__experiment=experiment)
            .order_by("replicate_id", "order")
            .values_list("replicate__replicate_number", "time", "signal_a")
        )
        imported_after = list(
            ImportedModelPoint.objects.filter(experiment=experiment)
            .order_by("model_key", "order")
            .values_list("model_key", "order", "x", "y", "sd", "status")
        )
        imported_calculations_after = list(
            ImportedReplicateCalculation.objects.filter(
                measurement__replicate__experiment=experiment
            ).order_by("measurement__replicate__replicate_number", "measurement__order")
            .values_list(
                "measurement__replicate__replicate_number",
                "measurement__order",
                "c_abs",
                "m_real",
                "release",
            )
        )
        aggregate_after = list(
            Measurement.objects.filter(experiment=experiment)
            .order_by("order")
            .values_list("time", "release", "sd")
        )
        self.assertEqual(raw_after, raw_before)
        self.assertEqual(imported_after, imported_before)
        self.assertEqual(imported_calculations_after, imported_calculations_before)
        self.assertEqual(aggregate_after, aggregate_before)

    def test_unified_upload_still_accepts_original_raw_csharp_workbook(self):
        response = self.client.post(
            "/upload/",
            {
                "file": csharp_raw_workbook_upload(sheet_count=1),
                "upload_token": "00000000-0000-0000-0000-000000000103",
            },
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Импортировано листов исходных измерений: 1")
        self.assertEqual(Experiment.objects.count(), 1)
        self.assertEqual(ImportedModelPoint.objects.count(), 0)


class SourceModelTransformationTests(TestCase):
    def test_linear_trendline_fits_free_intercept_and_counts_invalid_pairs(self):
        fit = fit_linear_trendline([
            (1, 5),
            (2, 7),
            (3, 9),
            (None, 11),
            (4, math.inf),
        ])

        self.assertTrue(fit["success"])
        self.assertAlmostEqual(fit["slope"], 2.0)
        self.assertAlmostEqual(fit["intercept"], 3.0)
        self.assertAlmostEqual(fit["r_squared"], 1.0)
        self.assertEqual(fit["point_count"], 3)
        self.assertEqual(fit["excluded_count"], 2)

    def test_linear_trendline_reports_constant_x_and_constant_y(self):
        no_x_variation = fit_linear_trendline([(2, 1), (2, 4)])
        self.assertFalse(no_x_variation["success"])
        self.assertIn("все значения X одинаковы", no_x_variation["error"])

        constant_y = fit_linear_trendline([(1, 7), (2, 7), (3, 7)])
        self.assertTrue(constant_y["success"])
        self.assertIsNone(constant_y["r_squared"])

    def test_csharp_coordinate_rules_and_boundary_handling(self):
        points = [(0, 0), (1, 20), (4, 100), (5, 120)]

        zero = transform_source_model(points, "zero_order")
        self.assertEqual(
            [(row["x"], row["y"]) for row in zero["records"]],
            points,
        )
        first = transform_source_model(points, "first_order")
        self.assertAlmostEqual(first["records"][0]["y"], math.log(100))
        self.assertEqual(first["included_count"], 2)
        self.assertEqual(first["invalid_domain_count"], 2)

        hixson = transform_source_model(points, "hixson_crowell")
        self.assertEqual(hixson["source_omission_count"], 1)
        self.assertEqual(hixson["invalid_domain_count"], 1)
        self.assertEqual(hixson["records"][0]["status"], "skipped_t0_source_rule")
        self.assertAlmostEqual(
            hixson["records"][1]["y"],
            math.pow(100, 1 / 3) - math.pow(80, 1 / 3),
        )

        peppas = transform_source_model(points, "korsmeyer_peppas")
        self.assertEqual(peppas["source_omission_count"], 1)
        self.assertEqual(peppas["included_count"], 3)
        self.assertEqual(peppas["records"][0]["status"], "skipped_t0_source_rule")
        self.assertAlmostEqual(peppas["records"][1]["x"], 0)
        self.assertAlmostEqual(peppas["records"][1]["y"], math.log(20))

        summaries = summarize_source_model_transforms(points)
        self.assertEqual(len(summaries), 5)
        self.assertEqual(summaries[-1]["source_omission_count"], 1)

    def test_csv_export_preserves_input_order_and_marks_skipped_origin(self):
        experiment = Experiment.objects.create(name="Transform", source_file="raw.xlsx")
        Measurement.objects.create(experiment=experiment, time=0, release=0, order=0)
        Measurement.objects.create(experiment=experiment, time=1, release=20, order=1)

        response = self.client.get(
            f"/experiments/{experiment.pk}/transforms/korsmeyer_peppas.csv"
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("transformed_x", response.content.decode("utf-8-sig"))
        self.assertIn("skipped_t0_by_source_rule", response.content.decode("utf-8-sig"))
        self.assertIn("included", response.content.decode("utf-8-sig"))

    def test_model_page_labels_csharp_transforms_separately_from_django_metrics(self):
        experiment = Experiment.objects.create(name="Model transforms", source_file="raw.xlsx")
        for order, (time, release) in enumerate(((0, 0), (1, 20), (2, 40))):
            Measurement.objects.create(
                experiment=experiment,
                time=time,
                release=release,
                order=order,
            )

        response = self.client.get("/models/", {"experiment_id": experiment.pk})

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Таблица показывает координаты моделей, найденные в C# исходнике")
        self.assertContains(response, "Эти подгонки Django используют собственные эмпирические уравнения")
        self.assertContains(response, f"/experiments/{experiment.pk}/transforms/zero_order.csv")
        self.assertContains(response, "Преобразованные координаты моделей")
        self.assertContains(response, "Исходное Time")
        self.assertEqual(response.context["source_coordinate_source"], "calculated")
        self.assertEqual(response.context["source_coordinate_count"], 3)
        linear_fit = response.context["source_coordinate_linear_fit"]
        self.assertTrue(linear_fit["success"])
        self.assertAlmostEqual(linear_fit["slope"], 20.0)
        self.assertAlmostEqual(linear_fit["intercept"], 0.0)
        self.assertAlmostEqual(linear_fit["r_squared"], 1.0)
        self.assertContains(response, "Линейная подгонка преобразованных координат")
        self.assertIsNotNone(response.context["source_coordinate_linear_plot"])

    def test_model_page_fits_imported_finite_pairs_and_reports_exclusions(self):
        experiment = Experiment.objects.create(
            name="Imported model points",
            source_file="AllData.xlsx",
        )
        ImportedModelPoint.objects.create(
            experiment=experiment,
            model_key="zero_order",
            order=0,
            x=0,
            y=5,
            status="invalid",
        )
        ImportedModelPoint.objects.create(
            experiment=experiment,
            model_key="zero_order",
            order=1,
            x=1,
            y=7,
            status="included",
        )
        ImportedModelPoint.objects.create(
            experiment=experiment,
            model_key="zero_order",
            order=2,
            x=None,
            y=None,
            status="invalid",
        )

        response = self.client.get("/models/", {"experiment_id": experiment.pk})

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["source_coordinate_source"], "imported")
        fit = response.context["source_coordinate_linear_fit"]
        self.assertTrue(fit["success"])
        self.assertAlmostEqual(fit["slope"], 2.0)
        self.assertAlmostEqual(fit["intercept"], 5.0)
        self.assertEqual(fit["point_count"], 2)
        self.assertEqual(fit["excluded_count"], 1)


class RawInputsMigrationTests(TransactionTestCase):
    """Ensure the new input tables preserve existing precomputed experiments."""

    migrate_from = [("kinetics", "0004_measurement_sd")]
    migrate_to = [("kinetics", "0005_experiment_calculation_version_and_more")]

    def setUp(self):
        super().setUp()
        executor = MigrationExecutor(connection)
        executor.migrate(self.migrate_from)
        old_apps = executor.loader.project_state(self.migrate_from).apps
        old_experiment = old_apps.get_model("kinetics", "Experiment")
        old_measurement = old_apps.get_model("kinetics", "Measurement")
        experiment = old_experiment.objects.create(
            name="Legacy experiment",
            description="",
            source_file="legacy.xlsx",
        )
        self.experiment_pk = experiment.pk
        old_measurement.objects.bulk_create([
            old_measurement(
                experiment_id=experiment.pk,
                time=index,
                release=index * 2.5,
                order=index,
                sd=None,
            )
            for index in range(40)
        ])

        MigrationExecutor(connection).migrate(self.migrate_to)

    def tearDown(self):
        MigrationExecutor(connection).migrate(
            MigrationExecutor(connection).loader.graph.leaf_nodes()
        )
        super().tearDown()

    def test_precomputed_rows_survive_and_new_fields_get_legacy_defaults(self):
        migrated_apps = MigrationExecutor(connection).loader.project_state(
            self.migrate_to
        ).apps
        experiment_model = migrated_apps.get_model("kinetics", "Experiment")
        measurement_model = migrated_apps.get_model("kinetics", "Measurement")
        experiment = experiment_model.objects.get(pk=self.experiment_pk)

        self.assertEqual(experiment.input_mode, "precalculated")
        self.assertEqual(experiment.calculation_version, "")
        self.assertIsNone(experiment.drug_loading)
        self.assertEqual(measurement_model.objects.filter(experiment_id=self.experiment_pk).count(), 40)


class ExperimentWorkflowTests(TestCase):
    def setUp(self):
        response = self.client.get("/")
        self.upload_token = response.context["upload_token"]

    def post_workbook(self, upload=None, token=None, follow=True):
        return self.client.post(
            "/upload/",
            {
                "file": upload or workbook_upload([(0, 0), (2, 12.4), (4, 28.1)]),
                "upload_token": token or self.upload_token,
            },
            follow=follow,
        )

    def test_home_opens_with_russian_empty_state(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Рабочее пространство")
        self.assertContains(response, "Загрузите файл .xlsx")

    def test_empty_data_page_and_sidebar_navigation(self):
        data_response = self.client.get("/data/")
        self.assertEqual(data_response.status_code, 200)
        self.assertContains(data_response, "Данные пока не загружены")

        home_html = self.client.get("/").content.decode()
        self.assertRegex(
            home_html,
            r'<a class="nav-item" href="/#overview" aria-current="page">',
        )
        self.assertEqual(home_html.count('aria-current="page">'), 1)

        data_html = data_response.content.decode()
        self.assertRegex(
            data_html,
            r'<a class="nav-item" href="/data/" aria-current="page">',
        )
        self.assertEqual(data_html.count('aria-current="page">'), 1)

    def test_home_summary_and_plot_use_saved_rows_without_rendering_the_table(self):
        self.post_workbook()

        response = self.client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "experiment.xlsx")
        self.assertContains(response, "Точек данных")
        self.assertContains(response, "data:image/png;base64,")
        self.assertContains(response, "28,1")
        self.assertNotContains(response, "28,1%")
        self.assertNotContains(response, "<td>12,4</td>")

    def test_valid_workbook_is_saved_and_shown_from_database(self):
        response = self.post_workbook()

        self.assertEqual(response.status_code, 200)
        self.assertEqual(len(response.redirect_chain), 1)
        experiment = Experiment.objects.get()
        self.assertEqual(experiment.name, "experiment")
        self.assertEqual(experiment.source_file, "experiment.xlsx")
        self.assertEqual(
            experiment.source_format,
            Experiment.SourceFormat.TIME_RELEASE_SD,
        )
        self.assertIsNotNone(experiment.created_at)
        self.assertIsNotNone(experiment.updated_at)
        self.assertContains(response, "data:image/png;base64,")
        self.assertContains(response, "Эксперимент и измерения сохранены.")
        self.assertEqual(
            list(experiment.measurements.order_by("order").values_list("time", "release", "order")),
            [(0.0, 0.0, 0), (2.0, 12.4, 1), (4.0, 28.1, 2)],
        )

        data_response = self.client.get("/data/")
        self.assertEqual(data_response.status_code, 200)
        self.assertContains(data_response, "experiment.xlsx")
        self.assertContains(data_response, "<th scope=\"col\">Высвобождение</th>", html=True)
        self.assertNotContains(data_response, "Высвобождение (%)")
        self.assertContains(data_response, "<td>12,4</td>", html=True)

    def test_upload_is_post_redirect_get_and_replay_does_not_duplicate(self):
        user = get_user_model().objects.create_user(
            username="xlsx-owner",
            password="test-password",
        )
        self.client.force_login(user)
        upload = workbook_upload([(0, 1), (1, 2)])
        first_response = self.post_workbook(upload=upload, follow=False)
        second_response = self.post_workbook(
            upload=workbook_upload([(0, 1), (1, 2)]),
            follow=False,
        )

        self.assertEqual(first_response.status_code, 302)
        self.assertEqual(second_response.status_code, 302)
        self.assertEqual(Experiment.objects.count(), 1)
        self.assertEqual(Experiment.objects.get().owner, user)
        self.assertEqual(Measurement.objects.count(), 2)

    def test_large_workbook_saves_measurements_in_order_across_batches(self):
        upload = workbook_upload(
            [(index, index / 10) for index in range(1_001)],
        )

        response = self.post_workbook(upload=upload, follow=False)

        self.assertEqual(response.status_code, 302)
        experiment = Experiment.objects.get()
        measurements = experiment.measurements.order_by("order")
        self.assertEqual(measurements.count(), 1_001)
        self.assertEqual(
            list(measurements.values_list("time", "release", "order")[:1]),
            [(0.0, 0.0, 0)],
        )
        self.assertEqual(
            list(measurements.values_list("time", "release", "order")[1_000:]),
            [(1_000.0, 100.0, 1_000)],
        )

    def test_invalid_extension_and_corrupt_workbook_are_rejected(self):
        wrong_extension = workbook_upload([(0, 1)], filename="data.csv")
        response = self.post_workbook(upload=wrong_extension)
        self.assertContains(response, "Поддерживается только формат .xlsx.")
        self.assertEqual(Experiment.objects.count(), 0)

        empty_file = SimpleUploadedFile("empty.xlsx", b"")
        response = self.post_workbook(upload=empty_file)
        self.assertContains(response, "Файл пуст.")
        self.assertEqual(Experiment.objects.count(), 0)

        corrupt_file = SimpleUploadedFile(
            "broken.xlsx",
            b"this is not an Excel workbook",
            content_type="application/octet-stream",
        )
        response = self.post_workbook(upload=corrupt_file)
        self.assertContains(response, "не является корректной книгой .xlsx")
        self.assertEqual(Experiment.objects.count(), 0)

        valid_xlsx = workbook_upload([(0, 1)])
        truncated_file = SimpleUploadedFile(
            "truncated.xlsx",
            valid_xlsx.read()[:-20],
            content_type="application/octet-stream",
        )
        response = self.post_workbook(upload=truncated_file)
        self.assertContains(response, "не является корректной книгой .xlsx")
        self.assertEqual(Experiment.objects.count(), 0)

    def test_excel_without_a_table_sheet_is_rejected_cleanly(self):
        class WorkbookWithoutWorksheet:
            active = None

            def close(self):
                return None

        upload = SimpleUploadedFile("chart-only.xlsx", b"not-read-by-mock")
        with patch(
            "kinetics.views.load_workbook",
            return_value=WorkbookWithoutWorksheet(),
        ):
            response = self.post_workbook(upload=upload)

        self.assertContains(response, "В книге нет листа с табличными измерениями.")
        self.assertEqual(Experiment.objects.count(), 0)

    def test_required_headers_and_finite_numeric_values_are_validated(self):
        bad_headers = workbook_upload([(0, 1)], headers=("Elapsed", "Amount"))
        response = self.post_workbook(upload=bad_headers)
        self.assertContains(response, "нужны заголовки Time (Время) и Release")
        self.assertEqual(Experiment.objects.count(), 0)

        duplicate_headers = workbook_upload(
            [(0, 1, 2)],
            headers=("Time", "Время", "Release"),
        )
        response = self.post_workbook(upload=duplicate_headers)
        self.assertContains(response, "Лист «Sheet»")
        self.assertContains(response, "повторный заголовок времени")
        self.assertContains(response, "столбце B")
        self.assertEqual(Experiment.objects.count(), 0)

        non_finite = workbook_upload([("NaN", 1)])
        response = self.post_workbook(upload=non_finite)
        self.assertContains(response, "должно быть конечным числом")
        self.assertEqual(Experiment.objects.count(), 0)

        for value in ("Infinity", "-Infinity"):
            with self.subTest(value=value):
                response = self.post_workbook(upload=workbook_upload([(value, 1)]))
                self.assertContains(response, "должно быть конечным числом")
                self.assertEqual(Experiment.objects.count(), 0)

        non_numeric = workbook_upload([(0, "не число")])
        response = self.post_workbook(upload=non_numeric)
        self.assertContains(response, "Лист «Sheet»")
        self.assertContains(response, "строке 2")
        self.assertContains(response, "столбец B")
        self.assertContains(response, "должно быть числом")
        self.assertEqual(Experiment.objects.count(), 0)

    def test_empty_data_and_partial_rows_are_rejected(self):
        blank_sheet = workbook_upload([], headers=[])
        response = self.post_workbook(upload=blank_sheet)
        self.assertContains(response, "Лист Excel пуст.")

        empty = workbook_upload([])
        response = self.post_workbook(upload=empty)
        self.assertContains(response, "нет строк с измерениями")

        partial = workbook_upload([(None, 3)])
        response = self.post_workbook(upload=partial)
        self.assertContains(response, "Пустое")
        self.assertEqual(Experiment.objects.count(), 0)

    def test_formula_and_excel_error_cells_are_rejected_without_saving(self):
        formula_workbook = Workbook()
        formula_sheet = formula_workbook.active
        formula_sheet.append(["Time", "Release"])
        formula_sheet.append(["=1+1", 10])
        formula_bytes = BytesIO()
        formula_workbook.save(formula_bytes)
        formula_workbook.close()
        formula_upload = SimpleUploadedFile(
            "formula.xlsx",
            formula_bytes.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response = self.post_workbook(upload=formula_upload)
        self.assertContains(response, "Пустое или некорректное значение")
        self.assertEqual(Experiment.objects.count(), 0)

        error_workbook = Workbook()
        error_sheet = error_workbook.active
        error_sheet.append(["Time", "Release"])
        error_sheet.append([0, "#DIV/0!"])
        error_bytes = BytesIO()
        error_workbook.save(error_bytes)
        error_workbook.close()
        error_upload = SimpleUploadedFile(
            "excel-error.xlsx",
            error_bytes.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response = self.post_workbook(upload=error_upload)
        self.assertContains(response, "должно быть числом")
        self.assertEqual(Experiment.objects.count(), 0)

    def test_header_aliases_and_unrestricted_measurements_remain_supported(self):
        upload = workbook_upload(
            [(-2, -5), (4, 125), (2, 75), (4, 140)],
            headers=("  ВРЕМЯ (часы)  ", "Высвобождение"),
        )

        response = self.post_workbook(upload=upload)

        self.assertEqual(response.status_code, 200)
        experiment = Experiment.objects.get()
        self.assertEqual(
            list(experiment.measurements.order_by("order").values_list("time", "release")),
            [(-2.0, -5.0), (4.0, 125.0), (2.0, 75.0), (4.0, 140.0)],
        )

    def test_csharp_export_column_names_are_accepted_in_a_tabular_upload(self):
        response = self.post_workbook(upload=workbook_upload(
            [(0, 12.5, 0.5), (5, 25, 1.25)],
            headers=("Time, min", "Drug release, %", "Standart deviation"),
        ))

        self.assertEqual(response.status_code, 200)
        experiment = Experiment.objects.get()
        self.assertEqual(
            list(experiment.measurements.order_by("order").values_list(
                "time", "release", "sd",
            )),
            [(0.0, 12.5, 0.5), (5.0, 25.0, 1.25)],
        )

    def test_stardartd_is_imported_and_kept_with_its_measurement(self):
        upload = workbook_upload(
            [(0, 0.4, 0), (1, 1.25, 10), (2, None, 20)],
            headers=("Time", "StardartD", "Release"),
        )

        response = self.post_workbook(upload=upload)

        self.assertEqual(response.status_code, 200)
        experiment = Experiment.objects.get()
        self.assertEqual(
            list(experiment.measurements.order_by("order").values_list("time", "release", "sd")),
            [(0.0, 0.0, 0.4), (1.0, 10.0, 1.25), (2.0, 20.0, None)],
        )
        plot_points, _ = _plot_points(experiment)
        self.assertEqual(
            plot_points,
            [(0.0, 0.0, 0.4), (1.0, 10.0, 1.25), (2.0, 20.0, None)],
        )
        detail = self.client.get(f"/experiments/{experiment.pk}/")
        self.assertContains(detail, "<th scope=\"col\">SD</th>", html=True)
        self.assertContains(detail, "<td>1,25</td>", html=True)
        self.assertContains(detail, "TTP и QTP")
        self.assertIsNone(detail.context["ttp"])
        self.assertIsNone(detail.context["qtp"])
        self.assertEqual(detail.context["model_fits"][0]["model_key"], "zero_order")

    def test_old_workbooks_without_stardartd_remain_supported(self):
        response = self.post_workbook(upload=workbook_upload([(0, 1), (1, 2)]))

        self.assertEqual(response.status_code, 200)
        experiment = Experiment.objects.get()
        self.assertEqual(
            list(experiment.measurements.order_by("order").values_list("sd", flat=True)),
            [None, None],
        )

    def test_invalid_stardartd_values_are_rejected_without_partial_import(self):
        for invalid_value, message in (
            (-0.1, "не может быть отрицательным"),
            ("не число", "должно быть числом"),
            ("Infinity", "должно быть конечным числом"),
        ):
            with self.subTest(value=invalid_value):
                response = self.post_workbook(upload=workbook_upload(
                    [(0, 1, 0.2), (1, invalid_value, 2)],
                    headers=("Time", "StardartD", "Release"),
                ))
                self.assertContains(response, message)
                self.assertEqual(Experiment.objects.count(), 0)
                self.assertEqual(Measurement.objects.count(), 0)

    def test_rejected_upload_preserves_existing_history_and_comparison_data(self):
        first = Experiment.objects.create(name="Сохранённый опыт", source_file="saved.xlsx")
        Measurement.objects.create(experiment=first, time=0, release=1, order=0)
        second = Experiment.objects.create(name="Второй опыт", source_file="second.xlsx")
        Measurement.objects.create(experiment=second, time=2, release=3, order=0)

        invalid = SimpleUploadedFile("bad.xlsx", b"not an xlsx")
        response = self.post_workbook(upload=invalid)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(Experiment.objects.count(), 2)
        self.assertEqual(Measurement.objects.count(), 2)
        history = self.client.get("/experiments/")
        comparison = self.client.get(
            "/comparison/",
            {"experiment_ids": [first.pk, second.pk]},
        )
        self.assertEqual(history.status_code, 200)
        self.assertContains(history, "Сохранённый опыт")
        self.assertContains(history, "Второй опыт")
        self.assertEqual(comparison.status_code, 200)
        self.assertContains(comparison, "Совместный график")

    def test_single_measurement_is_stored_but_model_page_explains_it_cannot_be_fitted(self):
        response = self.post_workbook(upload=workbook_upload([(7, 130)]))

        self.assertEqual(response.status_code, 200)
        experiment = Experiment.objects.get()
        model_response = self.client.get("/models/", {"experiment_id": experiment.pk})
        self.assertEqual(model_response.status_code, 200)
        self.assertContains(model_response, "как минимум три измерения")
        self.assertIsNone(model_response.context["model_fit"])

    def test_extremely_large_finite_measurements_do_not_raise_server_errors(self):
        with self.assertLogs("kinetics.views", level="ERROR"):
            response = self.post_workbook(
                upload=workbook_upload([(0, 1e308), (1, 1.7e308)]),
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(Experiment.objects.count(), 1)
        self.assertEqual(Measurement.objects.count(), 2)
        self.assertTrue(response.context["plot_error"])
        self.assertContains(response, "Не удалось построить график по сохранённым измерениям")

    def test_excel_close_failure_does_not_turn_a_valid_upload_into_a_server_error(self):
        def loader_with_close_failure(*args, **kwargs):
            workbook = openpyxl_load_workbook(*args, **kwargs)
            close = workbook.close

            def close_then_fail():
                close()
                raise OSError("simulated close failure")

            workbook.close = close_then_fail
            return workbook

        with self.assertLogs("kinetics.views", level="WARNING"):
            with patch(
                "kinetics.views.load_workbook",
                side_effect=loader_with_close_failure,
            ):
                response = self.post_workbook()

        self.assertEqual(response.status_code, 200)
        self.assertEqual(Experiment.objects.count(), 1)
        self.assertEqual(Measurement.objects.count(), 3)

    def test_upload_size_and_row_limits_are_enforced_before_saving(self):
        with patch("kinetics.views.MAX_EXCEL_UPLOAD_SIZE", 1):
            response = self.post_workbook()
        self.assertContains(response, "Размер файла превышает допустимый предел")
        self.assertEqual(Experiment.objects.count(), 0)

        with patch("kinetics.views.MAX_EXCEL_ROWS", 2):
            response = self.post_workbook(
                upload=workbook_upload([(0, 1), (1, 2), (2, 3)]),
            )
        self.assertContains(response, "не более 2 строк данных")
        self.assertEqual(Experiment.objects.count(), 0)

    def test_database_error_rolls_back_experiment_and_measurements(self):
        bulk_create = Measurement.objects.bulk_create
        call_count = 0

        def fail_after_first_batch(objects, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                raise DatabaseError("simulated write failure")
            return bulk_create(objects, **kwargs)

        large_upload = workbook_upload(
            [(index, index / 10) for index in range(1_001)],
        )
        with self.assertLogs("kinetics.views", level="ERROR"):
            with patch.object(
                Measurement.objects,
                "bulk_create",
                side_effect=fail_after_first_batch,
            ):
                response = self.post_workbook(upload=large_upload)

        self.assertContains(response, "Данные не были записаны частично.")
        self.assertEqual(call_count, 2)
        self.assertEqual(Experiment.objects.count(), 0)
        self.assertEqual(Measurement.objects.count(), 0)

    def test_history_and_saved_detail_use_the_selected_experiment(self):
        first = Experiment.objects.create(name="Первый", source_file="one.xlsx")
        Measurement.objects.create(experiment=first, time=1, release=10, order=0)
        second = Experiment.objects.create(name="Второй", source_file="two.xlsx")
        Measurement.objects.create(experiment=second, time=2, release=20, order=0)

        history = self.client.get("/experiments/")
        self.assertEqual(history.status_code, 200)
        self.assertContains(history, "Первый")
        self.assertContains(history, "Второй")
        self.assertContains(history, "Не определён (старая запись)")
        history_html = history.content.decode()
        self.assertEqual(history_html.count('aria-current="page">'), 1)

        with patch("kinetics.views._get_or_build_plot", return_value="cGxvdA==") as plot_builder:
            detail = self.client.get(f"/experiments/{first.pk}/")
        self.assertEqual(detail.status_code, 200)
        self.assertContains(detail, "one.xlsx")
        self.assertContains(detail, "data:image/png;base64,")
        self.assertContains(detail, "<td>10,0</td>", html=True)
        self.assertEqual(plot_builder.call_args.args[0], [(1.0, 10.0, None)])

    def test_detail_table_paginates_measurements_in_source_order(self):
        experiment = Experiment.objects.create(name="Большой опыт", source_file="many.xlsx")
        Measurement.objects.bulk_create([
            Measurement(
                experiment=experiment,
                time=float(index),
                release=250.5 if index == 100 else float(index),
                order=index,
            )
            for index in range(101)
        ])

        response = self.client.get(f"/experiments/{experiment.pk}/?page=2")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Страница 2 из 2")
        self.assertContains(
            response,
            "<tr><td>101</td><td>100,0</td><td>250,5</td><td>—</td></tr>",
            html=True,
        )
        self.assertEqual(
            [row.time for row in response.context["measurements_page"]],
            [100.0],
        )

    def test_excel_export_keeps_missing_sd_blank_and_source_order(self):
        experiment = Experiment.objects.create(
            name="Экспорт без SD",
            source_file="release.xlsx",
        )
        Measurement.objects.create(
            experiment=experiment, time=4, release=40, sd=None, order=0
        )
        Measurement.objects.create(
            experiment=experiment, time=1, release=10, sd=2.5, order=1
        )

        response = self.client.get(
            f"/experiments/{experiment.pk}/export.xlsx"
        )

        self.assertEqual(response.status_code, 200)
        workbook = openpyxl_load_workbook(BytesIO(response.content), data_only=True)
        results = workbook["Mean Release SD"]
        self.assertEqual(
            [(results.cell(row, 2).value, results.cell(row, 3).value)
             for row in (2, 3)],
            [(4, 40), (1, 10)],
        )
        self.assertIsNone(results.cell(2, 4).value)
        self.assertEqual(results.cell(3, 4).value, 2.5)
        analysis = dict(workbook["TTP QTP"].iter_rows(
            min_row=2, min_col=1, max_col=2, values_only=True
        ))
        self.assertEqual(analysis["TTP"], 4)
        self.assertEqual(analysis["Threshold P = Q_last − SD_last / 2"], 8.75)
        self.assertGreaterEqual(len(workbook["Charts"]._images), 1)
        workbook.close()

    def test_excel_export_distinguishes_zero_from_missing_parameters(self):
        experiment = Experiment.objects.create(
            name="Неполные старые параметры",
            source_file="legacy.xlsx",
            drug_loading=0,
            entrapment_efficiency=None,
            reactor_volume_ml=100,
            probe_volume_ml=0,
            calibration_coefficient=0,
        )
        Measurement.objects.create(
            experiment=experiment, time=0, release=0, sd=None, order=0
        )

        response = self.client.get(f"/experiments/{experiment.pk}/export.xlsx")

        self.assertEqual(response.status_code, 200)
        workbook = openpyxl_load_workbook(BytesIO(response.content), data_only=True)
        parameters = {
            row[0]: row
            for row in workbook["Input parameters"].iter_rows(
                min_row=2, values_only=True
            )
        }
        self.assertEqual(parameters["Drug loading (DL)"][1], 0)
        self.assertEqual(parameters["Probe volume"][1], 0)
        self.assertIsNone(parameters["Entrapment efficiency (EE)"][1])
        self.assertIn("Not saved", parameters["Entrapment efficiency (EE)"][3])
        self.assertIsNone(workbook["Mean Release SD"].cell(2, 4).value)
        self.assertEqual(workbook["Replicate readings"].max_row, 1)
        workbook.close()

    def test_excel_report_does_not_invent_ttp_when_last_source_sd_is_missing(self):
        experiment = Experiment.objects.create(
            name="Без SD последней точки",
            source_file="missing-last-sd.xlsx",
        )
        Measurement.objects.bulk_create([
            Measurement(
                experiment=experiment, time=0, release=10, sd=1, order=0
            ),
            Measurement(
                experiment=experiment, time=5, release=50, sd=None, order=1
            ),
        ])

        response = self.client.get(
            f"/experiments/{experiment.pk}/export.xlsx"
        )

        workbook = openpyxl_load_workbook(BytesIO(response.content), data_only=True)
        analysis = dict(workbook["TTP QTP"].iter_rows(
            min_row=2, min_col=1, max_col=2, values_only=True
        ))
        self.assertIsNone(analysis["TTP"])
        self.assertIsNone(analysis["QTP"])
        self.assertIsNone(analysis["Threshold P = Q_last − SD_last / 2"])
        self.assertIn("SD", analysis["Status"])
        workbook.close()

    def test_comparison_export_contains_selected_experiments_in_actual_time_order(self):
        first = Experiment.objects.create(name="Опыт A", source_file="a.xlsx")
        second = Experiment.objects.create(name="Опыт B", source_file="b.xlsx")
        Measurement.objects.create(
            experiment=first, time=2, release=20, sd=1, order=0
        )
        Measurement.objects.create(
            experiment=first, time=1, release=10, sd=None, order=1
        )
        Measurement.objects.create(
            experiment=second, time=1, release=15, sd=2, order=0
        )

        comparison_page = self.client.get(
            "/comparison/",
            {"experiment_ids": [first.pk, second.pk]},
        )
        self.assertContains(comparison_page, "Скачать сравнение Excel")
        response = self.client.get(
            "/comparison/export.xlsx",
            {"experiment_ids": [first.pk, second.pk]},
        )

        self.assertEqual(response.status_code, 200)
        workbook = openpyxl_load_workbook(BytesIO(response.content), data_only=True)
        self.assertEqual(
            workbook.sheetnames,
            ["Experiments", "Measurements by time", "Comparison notes"],
        )
        rows = list(workbook["Measurements by time"].iter_rows(
            min_row=2, values_only=True
        ))
        self.assertEqual(
            [(row[1], row[4], row[6], row[7]) for row in rows],
            [("Опыт A", 1, 10, None), ("Опыт B", 1, 15, 2), ("Опыт A", 2, 20, 1)],
        )
        summary_headers = {
            value: index
            for index, value in enumerate(
                next(workbook["Experiments"].iter_rows(values_only=True)),
                start=1,
            )
        }
        self.assertEqual(
            workbook["Experiments"].cell(2, summary_headers["Source format"]).value,
            "Не определён (старая запись)",
        )
        self.assertIn("единицы Time/Release не сохраняются", " ".join(
            str(row[1]) for row in workbook["Comparison notes"].iter_rows(
                min_row=2, values_only=True
            )
        ))
        workbook.close()

    def test_comparison_can_show_and_export_django_model_fits(self):
        first = Experiment.objects.create(name="Модель A", source_file="a.xlsx")
        second = Experiment.objects.create(name="Модель B", source_file="b.xlsx")
        values = [(0, 0), (1, 20), (2, 35), (3, 48)]
        for experiment, offset in ((first, 0), (second, 4)):
            Measurement.objects.bulk_create([
                Measurement(
                    experiment=experiment,
                    time=time,
                    release=release + offset,
                    order=index,
                )
                for index, (time, release) in enumerate(values)
            ])
        query = {"experiment_ids": [first.pk, second.pk], "compare_models": "1"}

        response = self.client.get("/comparison/", query)

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Подгонки моделей между экспериментами")
        self.assertContains(response, "Эти значения не являются результатами C#")
        self.assertEqual(len(response.context["model_comparison_rows"]), 10)
        self.assertTrue(any(
            row["experiment"].pk == first.pk
            and row["model"]["model_key"] == "zero_order"
            and row["model"]["success"]
            for row in response.context["model_comparison_rows"]
        ))

        export = self.client.get("/comparison/export.xlsx", query)
        self.assertEqual(export.status_code, 200)
        workbook = openpyxl_load_workbook(BytesIO(export.content), data_only=True)
        self.assertIn("Django model fits", workbook.sheetnames)
        fits = workbook["Django model fits"]
        self.assertEqual(fits.max_row, 11)
        self.assertEqual(fits.cell(2, 6).value, "Zero Order")
        self.assertIsNotNone(fits.cell(2, 10).value)
        workbook.close()

    def test_plot_uses_actual_time_order_while_the_table_keeps_source_order(self):
        experiment = Experiment.objects.create(
            name="Несортированное время",
            source_file="unsorted.xlsx",
        )
        Measurement.objects.bulk_create([
            Measurement(experiment=experiment, time=time, release=release, order=order)
            for order, (time, release) in enumerate([(4, 40), (0, 0), (2, 20)])
        ])

        with patch(
            "kinetics.views._get_or_build_plot",
            return_value="c29ydGVk",
        ) as plot_builder:
            response = self.client.get(f"/experiments/{experiment.pk}/")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            plot_builder.call_args.args[0],
            [(0.0, 0.0, None), (2.0, 20.0, None), (4.0, 40.0, None)],
        )
        self.assertEqual(
            [(row.time, row.release) for row in response.context["measurements_page"]],
            [(4.0, 40.0), (0.0, 0.0), (2.0, 20.0)],
        )

    def test_plot_sampling_is_bounded_and_keeps_the_last_measurement(self):
        experiment = Experiment.objects.create(name="График", source_file="plot.xlsx")
        Measurement.objects.bulk_create([
            Measurement(
                experiment=experiment,
                time=float(index),
                release=float(index),
                order=index,
            )
            for index in range(MAX_PLOT_POINTS * 2 + 1)
        ])

        points, fingerprint = _plot_points(experiment)

        self.assertLessEqual(len(points), MAX_PLOT_POINTS)
        self.assertEqual(points[0], (0.0, 0.0, None))
        self.assertEqual(
            points[-1],
            (float(MAX_PLOT_POINTS * 2), float(MAX_PLOT_POINTS * 2), None),
        )
        self.assertEqual(len(fingerprint), 64)

    def test_unknown_experiment_returns_404(self):
        response = self.client.get("/experiments/999999/")
        self.assertEqual(response.status_code, 404)

    def test_delete_requires_post_and_cascades_measurements(self):
        user = get_user_model().objects.create_user(
            username="delete-owner",
            password="test-password",
        )
        experiment = Experiment.objects.create(
            name="На удаление", source_file="x.xlsx", owner=user
        )
        Measurement.objects.create(experiment=experiment, time=1, release=2, order=0)
        self.client.force_login(user)

        self.assertEqual(
            self.client.get(f"/experiments/{experiment.pk}/delete/").status_code,
            405,
        )
        history = self.client.get("/experiments/")
        self.assertContains(history, "Удалить")
        self.assertContains(
            history,
            f'<form class="history-delete-form" action="/experiments/{experiment.pk}/delete/" method="post"',
            html=False,
        )
        self.assertContains(history, f'action="/experiments/{experiment.pk}/delete/"')
        csrf_client = Client(enforce_csrf_checks=True)
        csrf_client.force_login(user)
        csrf_client.get(f"/experiments/{experiment.pk}/")
        csrf_response = csrf_client.post(
            f"/experiments/{experiment.pk}/delete/",
        )
        self.assertEqual(csrf_response.status_code, 403)
        self.assertTrue(Experiment.objects.filter(pk=experiment.pk).exists())

        response = self.client.post(f"/experiments/{experiment.pk}/delete/", follow=True)

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Эксперимент «На удаление» удалён.")
        self.assertEqual(Experiment.objects.count(), 0)
        self.assertEqual(Measurement.objects.count(), 0)

    def test_delete_database_error_keeps_experiment_and_shows_message(self):
        user = get_user_model().objects.create_user(
            username="delete-error-owner",
            password="test-password",
        )
        experiment = Experiment.objects.create(
            name="Сохранить при ошибке", source_file="safe.xlsx", owner=user
        )
        Measurement.objects.create(experiment=experiment, time=1, release=2, order=0)
        self.client.force_login(user)

        with self.assertLogs("kinetics.views", level="ERROR"):
            with patch.object(
                Experiment,
                "delete",
                side_effect=DatabaseError("simulated delete failure"),
            ):
                response = self.client.post(
                    f"/experiments/{experiment.pk}/delete/",
                    follow=True,
                )

        self.assertContains(response, "Не удалось удалить эксперимент. Изменения отменены.")
        self.assertTrue(Experiment.objects.filter(pk=experiment.pk).exists())
        self.assertEqual(Measurement.objects.filter(experiment=experiment).count(), 1)

    def test_upload_and_delete_forms_require_csrf_token(self):
        client = Client(enforce_csrf_checks=True)
        home = client.get("/")
        self.assertEqual(home.status_code, 200)
        response = client.post(
            "/upload/",
            {
                "file": workbook_upload([(0, 1)]),
                "upload_token": home.context["upload_token"],
            },
        )
        self.assertEqual(response.status_code, 403)


class AuthenticationWorkflowTests(TestCase):
    def test_login_and_registration_are_exempt_from_production_login_middleware(self):
        for path in ("/accounts/login/", "/accounts/register/"):
            with self.subTest(path=path):
                view = resolve(path).func
                self.assertFalse(getattr(view, "login_required", True))

    def test_registration_hashes_password_logs_in_and_links_owner_to_upload(self):
        password = "Quartz!Forest-831-River"
        response = self.client.get("/accounts/register/")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Войти")
        self.assertContains(response, "Регистрация")

        response = self.client.post("/accounts/register/", {
            "username": "new-lab-user",
            "password1": password,
            "password2": password,
        })

        self.assertRedirects(response, "/")
        user = get_user_model().objects.get(username="new-lab-user")
        self.assertTrue(user.check_password(password))
        self.assertNotEqual(user.password, password)
        home = self.client.get("/")
        self.assertContains(home, "new-lab-user")
        self.assertContains(home, "Выйти")
        self.assertNotContains(home, "Регистрация")

        self.client.post("/upload/", {
            "file": workbook_upload([(0, 1), (1, 2)]),
            "upload_token": home.context["upload_token"],
        })
        experiment = Experiment.objects.get()
        self.assertEqual(experiment.owner_id, user.pk)

    def test_registration_validation_and_csrf(self):
        response = self.client.post("/accounts/register/", {
            "username": "invalid-registration",
            "password1": "Quartz!Forest-831-River",
            "password2": "different-password",
        })
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.context["form"].errors)
        self.assertFalse(
            get_user_model().objects.filter(username="invalid-registration").exists()
        )

        csrf_client = Client(enforce_csrf_checks=True)
        self.assertEqual(csrf_client.get("/accounts/register/").status_code, 200)
        rejected = csrf_client.post("/accounts/register/", {
            "username": "csrf-user",
            "password1": "Quartz!Forest-831-River",
            "password2": "Quartz!Forest-831-River",
        })
        self.assertEqual(rejected.status_code, 403)
        self.assertFalse(get_user_model().objects.filter(username="csrf-user").exists())

        self.assertEqual(csrf_client.get("/accounts/login/").status_code, 200)
        self.assertEqual(csrf_client.post("/accounts/login/", {
            "username": "csrf-user",
            "password": "Quartz!Forest-831-River",
        }).status_code, 403)

        user = get_user_model().objects.create_user(
            username="csrf-logout-user",
            password="test-password",
        )
        csrf_client.force_login(user)
        self.assertEqual(csrf_client.post("/accounts/logout/").status_code, 403)

    def test_login_logout_and_invalid_login(self):
        user = get_user_model().objects.create_user(
            username="auth-user",
            password="Quartz!Forest-831-River",
        )
        self.assertContains(self.client.get("/accounts/login/"), "Войти")

        invalid = self.client.post("/accounts/login/", {
            "username": user.username,
            "password": "incorrect-password",
        })
        self.assertEqual(invalid.status_code, 200)
        self.assertTrue(invalid.context["form"].errors)

        response = self.client.post("/accounts/login/", {
            "username": user.username,
            "password": "Quartz!Forest-831-River",
        }, follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Вы вошли как auth-user.")
        self.assertContains(response, "auth-user")
        self.assertContains(response, "Выйти")

        response = self.client.post("/accounts/logout/", follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Вы вышли из учётной записи.")
        self.assertContains(response, "Регистрация")
        self.assertNotContains(response, "auth-user")
        self.assertEqual(self.client.get("/accounts/logout/").status_code, 405)


class AccountSettingsTests(TestCase):
    def setUp(self):
        self.current_password = "Current!Password-5829"
        self.user = get_user_model().objects.create_user(
            username="settings-owner",
            password=self.current_password,
            first_name="Анна",
            last_name="Иванова",
        )
        self.client.force_login(self.user)

    def profile_data(self, **overrides):
        data = {
            "section": "profile",
            "first_name": "Мария",
            "last_name": "Соколова",
        }
        data.update(overrides)
        return data

    def password_data(self, **overrides):
        data = {
            "section": "password",
            "old_password": self.current_password,
            "new_password1": "Different!Strong-Password-6317",
            "new_password2": "Different!Strong-Password-6317",
        }
        data.update(overrides)
        return data

    def test_anonymous_user_is_redirected_to_login(self):
        self.client.logout()

        response = self.client.get("/settings/")

        self.assertRedirects(response, "/accounts/login/?next=/settings/")

    def test_signed_in_user_can_open_settings_and_sees_account_and_app_details(self):
        response = self.client.get("/settings/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Настройки")
        self.assertContains(response, "settings-owner")
        self.assertContains(response, self.user.date_joined.strftime("%d.%m.%Y"))
        self.assertContains(response, "Анна Иванова")
        self.assertContains(response, "Учётная запись")
        self.assertContains(response, "Безопасность")
        self.assertContains(response, "Моделирование кинетики высвобождения действующих веществ")
        self.assertContains(response, "v1.0")
        self.assertContains(response, 'href="/settings/" aria-current="page"')

    def test_display_name_is_saved_for_the_signed_in_user(self):
        response = self.client.post("/settings/", self.profile_data())

        self.assertRedirects(response, "/settings/")
        self.user.refresh_from_db()
        self.assertEqual((self.user.first_name, self.user.last_name), ("Мария", "Соколова"))
        self.assertContains(self.client.get("/settings/"), "Мария Соколова")

    def test_submitted_user_id_cannot_change_another_account(self):
        other_user = get_user_model().objects.create_user(
            username="other-settings-user",
            password="Other!Password-9354",
            first_name="Другая",
            last_name="Пользователь",
        )

        response = self.client.post(
            "/settings/",
            self.profile_data(user_id=other_user.pk),
        )

        self.assertRedirects(response, "/settings/")
        self.user.refresh_from_db()
        other_user.refresh_from_db()
        self.assertEqual((self.user.first_name, self.user.last_name), ("Мария", "Соколова"))
        self.assertEqual(
            (other_user.first_name, other_user.last_name),
            ("Другая", "Пользователь"),
        )

    def test_valid_password_change_keeps_the_user_authenticated(self):
        response = self.client.post("/settings/", self.password_data())

        self.assertRedirects(response, "/settings/")
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password("Different!Strong-Password-6317"))
        self.assertEqual(self.client.get("/settings/").status_code, 200)

    def test_wrong_current_password_is_rejected(self):
        response = self.client.post(
            "/settings/",
            self.password_data(old_password="Wrong!Password-7492"),
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("old_password", response.context["password_form"].errors)
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password(self.current_password))

    def test_password_confirmation_must_match(self):
        response = self.client.post(
            "/settings/",
            self.password_data(new_password2="Another!Password-1111"),
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("new_password2", response.context["password_form"].errors)
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password(self.current_password))

    def test_django_password_validators_reject_a_weak_new_password(self):
        response = self.client.post(
            "/settings/",
            self.password_data(
                new_password1="12345678",
                new_password2="12345678",
            ),
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("new_password2", response.context["password_form"].errors)
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password(self.current_password))

    def test_account_forms_require_csrf(self):
        csrf_client = Client(enforce_csrf_checks=True)
        csrf_client.force_login(self.user)
        self.assertEqual(csrf_client.get("/settings/").status_code, 200)

        profile_response = csrf_client.post("/settings/", self.profile_data())
        password_response = csrf_client.post("/settings/", self.password_data())

        self.assertEqual(profile_response.status_code, 403)
        self.assertEqual(password_response.status_code, 403)
        self.user.refresh_from_db()
        self.assertEqual((self.user.first_name, self.user.last_name), ("Анна", "Иванова"))
        self.assertTrue(self.user.check_password(self.current_password))

    def test_home_history_and_documentation_still_work(self):
        for path in ("/", "/experiments/", "/method/"):
            with self.subTest(path=path):
                self.assertEqual(self.client.get(path).status_code, 200)


class ExperimentVisibilityTests(TestCase):
    def setUp(self):
        user_model = get_user_model()
        self.alice = user_model.objects.create_user(
            username="alice",
            password="test-password",
        )
        self.bob = user_model.objects.create_user(
            username="bob",
            password="test-password",
        )
        self.alice_experiment = Experiment.objects.create(
            name="Alice private data",
            source_file="alice.xlsx",
            owner=self.alice,
            source_format=Experiment.SourceFormat.TIME_RELEASE_SD,
        )
        self.bob_experiment = Experiment.objects.create(
            name="Bob private data",
            source_file="bob.xlsx",
            owner=self.bob,
            source_format=Experiment.SourceFormat.TIME_RELEASE_SD,
        )
        for experiment in (self.alice_experiment, self.bob_experiment):
            Measurement.objects.create(
                experiment=experiment,
                time=0,
                release=10,
                order=0,
            )
            Measurement.objects.create(
                experiment=experiment,
                time=1,
                release=20,
                order=1,
            )
        self.shared_example = Experiment.objects.create(
            name="Общий пример",
            source_file="example.xlsx",
            source_format=Experiment.SourceFormat.SYNTHETIC_RAW,
        )

    @override_settings(
        DEBUG=False,
        APOLLINEX_ALLOW_GUEST_UNOWNED_EXPERIMENTS=False,
    )
    def test_users_only_see_owned_and_shared_experiments_across_routes(self):
        legacy_experiment = Experiment.objects.create(
            name="Unassigned legacy work",
            source_file="legacy.xlsx",
        )
        guest_history = self.client.get("/experiments/")
        self.assertContains(guest_history, "Общий пример")
        self.assertNotContains(guest_history, "Unassigned legacy work")
        self.assertEqual(
            self.client.get(f"/experiments/{legacy_experiment.pk}/").status_code,
            404,
        )

        self.client.force_login(self.alice)

        history = self.client.get("/experiments/")
        self.assertContains(history, "Alice private data")
        self.assertContains(history, "Общий пример")
        self.assertNotContains(history, "Bob private data")

        self.assertEqual(
            self.client.get(f"/experiments/{self.bob_experiment.pk}/").status_code,
            404,
        )
        self.assertEqual(
            self.client.get(
                f"/experiments/{self.bob_experiment.pk}/export.xlsx"
            ).status_code,
            404,
        )
        self.assertEqual(
            self.client.get(
                f"/experiments/{self.bob_experiment.pk}/transforms/zero_order.csv"
            ).status_code,
            404,
        )
        self.assertEqual(
            self.client.get(
                f"/experiments/{self.bob_experiment.pk}/imported-transforms/zero_order.csv"
            ).status_code,
            404,
        )

        data_page = self.client.get(
            "/data/",
            {"experiment_id": self.bob_experiment.pk},
        )
        self.assertNotContains(data_page, "Bob private data")
        models_page = self.client.get(
            "/models/",
            {"experiment_id": self.bob_experiment.pk},
        )
        self.assertNotContains(models_page, "Bob private data")
        report_page = self.client.get(
            "/reports/",
            {"experiment_id": self.bob_experiment.pk},
        )
        self.assertNotContains(report_page, "Bob private data")

        comparison = self.client.get("/comparison/", {
            "experiment_ids": [self.alice_experiment.pk, self.bob_experiment.pk],
        })
        self.assertNotContains(comparison, "Bob private data")
        export = self.client.get("/comparison/export.xlsx", {
            "experiment_ids": [self.alice_experiment.pk, self.bob_experiment.pk],
        })
        self.assertEqual(export.status_code, 302)

        for path, data in (
            (f"/experiments/{self.bob_experiment.pk}/edit/", {}),
            (f"/experiments/{self.bob_experiment.pk}/recalculate/", {}),
            (f"/experiments/{self.bob_experiment.pk}/delete/", {}),
        ):
            with self.subTest(path=path):
                self.assertEqual(self.client.post(path, data).status_code, 404)
        self.assertTrue(Experiment.objects.filter(pk=self.bob_experiment.pk).exists())
        self.assertEqual(Measurement.objects.filter(
            experiment=self.bob_experiment,
        ).count(), 2)

    def test_local_guest_workflow_remains_available_without_exposing_owned_work(self):
        unowned = Experiment.objects.create(
            name="Local unowned legacy data",
            source_file="legacy.xlsx",
        )

        history = self.client.get("/experiments/")

        self.assertEqual(history.status_code, 200)
        self.assertContains(history, "Local unowned legacy data")
        self.assertContains(history, "Общий пример")
        self.assertNotContains(history, "Alice private data")
        self.assertNotContains(history, "Bob private data")
        self.assertTrue(Experiment.objects.filter(pk=unowned.pk).exists())


class ExperimentDeletionOwnershipTests(TestCase):
    def setUp(self):
        user_model = get_user_model()
        self.owner = user_model.objects.create_user(
            username="work-owner",
            password="test-password",
        )
        self.other_user = user_model.objects.create_user(
            username="other-owner",
            password="test-password",
        )
        self.client.force_login(self.owner)

    def create_experiment(self, owner, *, source_format=Experiment.SourceFormat.MANUAL_RAW):
        return Experiment.objects.create(
            name="Опыт пользователя",
            source_file="input.xlsx",
            owner=owner,
            source_format=source_format,
        )

    def test_owner_can_delete_only_their_work_and_its_cascaded_data(self):
        own = self.create_experiment(self.owner)
        other = self.create_experiment(self.other_user)
        for experiment in (own, other):
            Measurement.objects.create(
                experiment=experiment,
                time=1,
                release=25,
                sd=1,
                order=0,
            )
            replicate = RawReplicate.objects.create(
                experiment=experiment,
                replicate_number=1,
                sample_mass_g=0.25,
            )
            raw_measurement = RawMeasurement.objects.create(
                replicate=replicate,
                time=1,
                signal_a=2,
                order=0,
            )
            ReplicateCalculation.objects.create(
                measurement=raw_measurement,
                c_abs=1,
                m_real=2,
                release=8,
                algorithm_version="source-float32-1",
            )
            ImportedReplicateCalculation.objects.create(
                measurement=raw_measurement,
                c_abs=1,
                m_real=2,
                release=8,
            )
            ImportedModelPoint.objects.create(
                experiment=experiment,
                model_key="zero_order",
                order=0,
                x=1,
                y=25,
            )

        response = self.client.post(
            f"/experiments/{own.pk}/delete/",
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Эксперимент «Опыт пользователя» удалён.")
        self.assertFalse(Experiment.objects.filter(pk=own.pk).exists())
        self.assertTrue(Experiment.objects.filter(pk=other.pk).exists())
        self.assertEqual(Measurement.objects.count(), 1)
        self.assertEqual(RawReplicate.objects.count(), 1)
        self.assertEqual(RawMeasurement.objects.count(), 1)
        self.assertEqual(ReplicateCalculation.objects.count(), 1)
        self.assertEqual(ImportedReplicateCalculation.objects.count(), 1)
        self.assertEqual(ImportedModelPoint.objects.count(), 1)

    def test_foreign_and_shared_synthetic_experiments_cannot_be_deleted_by_id(self):
        foreign = self.create_experiment(self.other_user)
        shared_example = self.create_experiment(
            None,
            source_format=Experiment.SourceFormat.SYNTHETIC_RAW,
        )
        history = self.client.get("/experiments/")
        self.assertNotContains(
            history,
            f'action="/experiments/{foreign.pk}/delete/"',
        )
        self.assertNotContains(
            history,
            f'action="/experiments/{shared_example.pk}/delete/"',
        )

        for experiment in (foreign, shared_example):
            with self.subTest(experiment=experiment.pk):
                response = self.client.post(
                    f"/experiments/{experiment.pk}/delete/"
                )
                self.assertEqual(response.status_code, 404)
                self.assertTrue(Experiment.objects.filter(pk=experiment.pk).exists())

    def test_user_owned_copy_of_synthetic_example_can_be_deleted(self):
        copy = self.create_experiment(
            self.owner,
            source_format=Experiment.SourceFormat.SYNTHETIC_RAW,
        )

        response = self.client.post(
            f"/experiments/{copy.pk}/delete/",
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Эксперимент «Опыт пользователя» удалён.")
        self.assertFalse(Experiment.objects.filter(pk=copy.pk).exists())


class AssignExperimentOwnerCommandTests(TestCase):
    command_name = "assign_experiment_owner"

    def setUp(self):
        user_model = get_user_model()
        self.user = user_model.objects.create_user(
            username="legacy-owner",
            password="test-password",
        )
        self.client.force_login(self.user)

    def test_admin_cannot_bypass_the_guarded_assignment_command(self):
        from .admin import ExperimentAdmin

        self.assertIn("owner", ExperimentAdmin.readonly_fields)

    def test_explicit_assignment_changes_only_owner_and_shows_delete_button(self):
        experiment = Experiment.objects.create(
            name="Старая работа",
            description="Описание для проверки",
            source_file="legacy.xlsx",
            source_format=Experiment.SourceFormat.LEGACY_UNKNOWN,
        )
        measurement = Measurement.objects.create(
            experiment=experiment,
            time=2,
            release=30,
            sd=1.5,
            order=0,
        )
        replicate = RawReplicate.objects.create(
            experiment=experiment,
            replicate_number=1,
            sample_mass_g=0.25,
        )
        raw_measurement = RawMeasurement.objects.create(
            replicate=replicate,
            time=2,
            signal_a=0.04,
            order=0,
        )
        calculation = ReplicateCalculation.objects.create(
            measurement=raw_measurement,
            c_abs=0.5,
            m_real=0.125,
            release=30,
            algorithm_version="legacy-test",
        )
        imported_calculation = ImportedReplicateCalculation.objects.create(
            measurement=raw_measurement,
            c_abs=0.5,
            m_real=0.125,
            release=30,
            source_format="csharp-alldata",
        )
        imported_point = ImportedModelPoint.objects.create(
            experiment=experiment,
            model_key="zero_order",
            order=0,
            x=2,
            y=30,
            sd=1.5,
        )
        original_updated_at = experiment.updated_at

        output = StringIO()
        call_command(
            self.command_name,
            str(experiment.pk),
            str(self.user.pk),
            stdout=output,
        )

        experiment.refresh_from_db()
        self.assertEqual(experiment.owner_id, self.user.pk)
        self.assertEqual(experiment.updated_at, original_updated_at)
        self.assertEqual(
            (experiment.name, experiment.description, experiment.source_file,
             experiment.source_format),
            ("Старая работа", "Описание для проверки", "legacy.xlsx",
             Experiment.SourceFormat.LEGACY_UNKNOWN),
        )
        measurement.refresh_from_db()
        self.assertEqual(
            (measurement.time, measurement.release, measurement.sd, measurement.order),
            (2, 30, 1.5, 0),
        )
        raw_measurement.refresh_from_db()
        calculation.refresh_from_db()
        imported_calculation.refresh_from_db()
        imported_point.refresh_from_db()
        self.assertEqual((raw_measurement.time, raw_measurement.signal_a), (2, 0.04))
        self.assertEqual(
            (calculation.c_abs, calculation.m_real, calculation.release,
             calculation.algorithm_version),
            (0.5, 0.125, 30, "legacy-test"),
        )
        self.assertEqual(
            (imported_calculation.c_abs, imported_calculation.m_real,
             imported_calculation.release, imported_calculation.source_format),
            (0.5, 0.125, 30, "csharp-alldata"),
        )
        self.assertEqual(
            (imported_point.model_key, imported_point.x, imported_point.y,
             imported_point.sd),
            ("zero_order", 2, 30, 1.5),
        )
        self.assertIn(f"Эксперимент {experiment.pk} назначен пользователю", output.getvalue())

        history = self.client.get("/experiments/")
        self.assertContains(
            history,
            f'action="/experiments/{experiment.pk}/delete/"',
        )
        self.assertContains(history, "Удалить")

        unselected = Experiment.objects.create(
            name="Другая старая работа",
            source_file="another-legacy.xlsx",
        )
        history = self.client.get("/experiments/")
        self.assertNotContains(
            history,
            f'action="/experiments/{unselected.pk}/delete/"',
        )
        unselected.refresh_from_db()
        self.assertIsNone(unselected.owner_id)

    def test_shared_demo_examples_are_never_assigned(self):
        example_definition = next(iter(SYNTHETIC_EXAMPLES.values()))
        examples = [
            Experiment.objects.create(
                name="Синтетический пример",
                source_file="synthetic.xlsx",
                source_format=Experiment.SourceFormat.SYNTHETIC_RAW,
            ),
            Experiment.objects.create(
                name="Старый общий пример",
                source_file=example_definition["source_file"],
                source_format=Experiment.SourceFormat.LEGACY_UNKNOWN,
            ),
            Experiment.objects.create(
                name=example_definition["initial"]["name"],
                source_file="renamed-example.xlsx",
                source_format=Experiment.SourceFormat.LEGACY_UNKNOWN,
            ),
            Experiment.objects.create(
                name="Ещё одна старая копия примера",
                source_file="renamed-example-2.xlsx",
                description=example_definition["initial"]["description"],
                source_format=Experiment.SourceFormat.LEGACY_UNKNOWN,
            ),
        ]

        for experiment in examples:
            with self.subTest(experiment=experiment.pk):
                with self.assertRaisesRegex(CommandError, "демонстрационным примером"):
                    call_command(
                        self.command_name,
                        str(experiment.pk),
                        str(self.user.pk),
                    )
                experiment.refresh_from_db()
                self.assertIsNone(experiment.owner_id)

        history = self.client.get("/experiments/")
        for experiment in examples:
            self.assertNotContains(
                history,
                f'action="/experiments/{experiment.pk}/delete/"',
            )

    def test_owned_missing_experiment_and_missing_user_are_refused(self):
        owned = Experiment.objects.create(
            name="Уже назначена",
            source_file="owned.xlsx",
            owner=self.user,
        )
        with self.assertRaisesRegex(CommandError, "уже есть владелец"):
            call_command(self.command_name, str(owned.pk), str(self.user.pk))

        with self.assertRaisesRegex(CommandError, "не найден"):
            call_command(self.command_name, "999999", str(self.user.pk))

        unassigned = Experiment.objects.create(
            name="Останется без владельца",
            source_file="unassigned.xlsx",
        )
        with self.assertRaisesRegex(CommandError, "Пользователь с ID 999999 не найден"):
            call_command(self.command_name, str(unassigned.pk), "999999")
        unassigned.refresh_from_db()
        self.assertIsNone(unassigned.owner_id)


class AnalysisSectionsTests(TestCase):
    def create_experiment(self, name, rows, source_file=None):
        experiment = Experiment.objects.create(
            name=name,
            source_file=source_file or f"{name}.xlsx",
        )
        measurements = []
        for order, row in enumerate(rows):
            time, release, *optional_sd = row
            measurements.append(Measurement(
                experiment=experiment,
                time=time,
                release=release,
                sd=optional_sd[0] if optional_sd else None,
                order=order,
            ))
        Measurement.objects.bulk_create(measurements)
        return experiment

    def test_ttp_qtp_matches_source_algorithm_and_handles_edge_cases(self):
        result = _calculate_ttp_qtp([
            (5, 21, None),
            (1, 10, None),
            (3, 20, 2),
        ])
        self.assertEqual(result["ttp"], 5)
        self.assertEqual(result["qtp"], 21)
        self.assertEqual(result["ttp_qtp_threshold"], 19)

        # Equality to P does not cross it: the comparison is strictly >.
        boundary = _calculate_ttp_qtp([(1, 9, None), (2, 10, 2)])
        self.assertEqual(boundary["ttp"], 2)
        self.assertEqual(boundary["qtp"], 10)

        single = _calculate_ttp_qtp([(0, 10, 2)])
        self.assertEqual(single["ttp"], 0)
        self.assertEqual(single["qtp"], 10)

        repeated_time = _calculate_ttp_qtp([
            (2, 20, None),
            (2, 21, None),
            (2, 15, 2),
        ])
        self.assertEqual(repeated_time["ttp"], 2)
        self.assertEqual(repeated_time["qtp"], 20)

        zero_sd = _calculate_ttp_qtp([(0, 4, None), (1, 5, 0)])
        self.assertIsNone(zero_sd["ttp"])
        self.assertIn("не пересекают порог", zero_sd["ttp_qtp_message"])

        for rows, message_part in (
            ([], "нет измерений"),
            ([(0, 1, None), (1, 2, None)], "отсутствует SD"),
            ([(0, 1, 1), (1, 2, None)], "отсутствует SD"),
            ([(0, float("nan"), None), (1, 2, 1)], "некорректные числовые"),
            ([(0, 1, None), (1, 2, -1)], "неотрицательным числом"),
            ([(0, 1, None), (1, 2, float("inf"))], "неотрицательным числом"),
        ):
            with self.subTest(rows=rows):
                unavailable = _calculate_ttp_qtp(rows)
                self.assertIsNone(unavailable["ttp"])
                self.assertIsNone(unavailable["qtp"])
                self.assertIn(message_part, unavailable["ttp_qtp_message"])

        # The source implementation processes rows in their supplied order;
        # it does not silently sort an irregular time series.
        negative_time = _calculate_ttp_qtp([(-2, 11, None), (0, 5, 2)])
        self.assertEqual(negative_time["ttp"], -2)

    def test_ttp_qtp_is_identical_on_detail_model_and_report_pages(self):
        experiment = self.create_experiment(
            "TTP/QTP consistency",
            [(1, 2, None), (2, 4, None), (3, 6, 1)],
        )

        responses = (
            self.client.get(f"/experiments/{experiment.pk}/"),
            self.client.get("/models/", {"experiment_id": experiment.pk}),
            self.client.get("/reports/", {"experiment_id": experiment.pk}),
        )
        self.assertTrue(all(response.status_code == 200 for response in responses))
        for response in responses:
            with self.subTest(path=response.request["PATH_INFO"]):
                self.assertEqual(response.context["ttp"], 3)
                self.assertEqual(response.context["qtp"], 6)
                self.assertEqual(response.context["ttp_qtp_threshold"], 5.5)
                self.assertContains(response, "Q<sub>last</sub>")

    def test_analysis_routes_render_and_navigation_links_are_active(self):
        routes = (
            ("kinetic_models", "/models/", "Кинетические модели"),
            ("comparison", "/comparison/", "Сравнение экспериментов"),
            ("reports", "/reports/", "Отчёты"),
        )
        for route_name, path, page_title in routes:
            with self.subTest(path=path):
                response = self.client.get(path)
                self.assertEqual(response.status_code, 200)
                self.assertContains(response, page_title)
                self.assertEqual(response.request["PATH_INFO"], path)
                self.assertEqual(
                    response.content.decode().count('aria-current="page">'),
                    1,
                )
                self.assertEqual(response.resolver_match.url_name, route_name)
                if route_name == "kinetic_models":
                    self.assertNotContains(response, "Подогнать модель")
                if route_name == "reports":
                    self.assertNotContains(response, "Сформировать отчёт")

    def test_first_order_fit_returns_parameters_and_rejects_insufficient_data(self):
        expected_rate = 0.35
        expected_plateau = 80.0
        points = [
            (time, expected_plateau * (1 - math.exp(-expected_rate * time)))
            for time in (0, 1, 2, 4, 8)
        ]

        fit, error = _first_order_model_fit(points)

        self.assertEqual(error, "")
        self.assertAlmostEqual(fit["rate"], expected_rate, places=5)
        self.assertAlmostEqual(fit["plateau"], expected_plateau, places=4)
        self.assertLess(fit["rmse"], 1e-5)
        self.assertIsNotNone(fit["r_squared"])
        self.assertIsNone(_first_order_model_fit([(0, 0), (1, 2)])[0])

    def test_first_order_fit_handles_constant_zero_repeated_and_extreme_values(self):
        constant_fit, constant_error = _first_order_model_fit(
            [(0, 5), (1, 5), (2, 5)],
        )
        self.assertEqual(constant_error, "")
        self.assertIsNone(constant_fit["r_squared"])
        self.assertTrue(all(
            math.isfinite(value)
            for value in (constant_fit["plateau"], constant_fit["rate"], constant_fit["rmse"])
        ))

        for points in (
            [(0, 0), (0, 1), (0, 2)],
            [(0, 0), (1, 0), (2, 0)],
            [(0, 1), (1, 2), (float("inf"), 3)],
            [(0, 1), (5e-324, 2), (1e-323, 3)],
        ):
            with self.subTest(points=points):
                fit, error = _first_order_model_fit(points)
                self.assertIsNone(fit)
                self.assertTrue(error)

    def test_plot_builder_failure_is_reported_for_detail_and_report_pages(self):
        experiment = self.create_experiment(
            "Сбой графика",
            [(0, 0), (1, 2), (2, 3)],
        )

        with self.assertLogs("kinetics.views", level="ERROR"):
            with patch(
                "kinetics.views._build_plot",
                side_effect=RuntimeError("simulated plotting failure"),
            ):
                detail = self.client.get(f"/experiments/{experiment.pk}/")

        self.assertEqual(detail.status_code, 200)
        self.assertTrue(detail.context["plot_error"])
        self.assertContains(detail, "Не удалось построить график по сохранённым измерениям")

        with patch("kinetics.views._get_or_build_plot", return_value=None):
            report = self.client.get("/reports/", {"experiment_id": experiment.pk})

        self.assertEqual(report.status_code, 200)
        self.assertTrue(report.context["report_plot_error"])
        self.assertContains(report, "Не удалось построить график по сохранённым измерениям")

    def test_model_and_comparison_plot_builders_render_pngs(self):
        fit, error = _first_order_model_fit([(0, 0), (1, 4), (2, 6)])
        self.assertEqual(error, "")
        model_plot = base64.b64decode(
            _build_model_plot([(0, 0), (1, 4), (2, 6)], fit)
        )
        first = Experiment(name="Первый", source_file="first.xlsx")
        second = Experiment(name="Второй", source_file="second.xlsx")
        comparison_plot = base64.b64decode(_build_comparison_plot([
            {"experiment": first, "points": [(0, 0), (2, 4)]},
            {"experiment": second, "points": [(1, 3), (4, 8)]},
        ]))

        self.assertTrue(model_plot.startswith(b"\x89PNG\r\n\x1a\n"))
        self.assertTrue(comparison_plot.startswith(b"\x89PNG\r\n\x1a\n"))

    def test_scientific_plot_draws_supplied_standard_deviation_error_bars(self):
        with patch("kinetics.views.plt.Axes.errorbar") as errorbar:
            image = base64.b64decode(
                _build_plot([(0, 1, 0.25), (1, 2, 0.5), (2, 3, None)])
            )

        self.assertTrue(image.startswith(b"\x89PNG\r\n\x1a\n"))
        errorbar.assert_called_once()
        self.assertEqual(errorbar.call_args.kwargs["yerr"], [0.25, 0.5])

        with patch("kinetics.views.plt.Axes.errorbar") as missing_errorbar:
            _build_plot([(0, 1, None), (1, 2, None)])
        missing_errorbar.assert_not_called()

    def test_model_comparison_calculates_raw_scale_r_squared_and_selects_the_best(self):
        points = [(time, 4.0 * time) for time in (0, 1, 2, 3, 4)]

        models, best = _fit_kinetic_models(points)

        self.assertEqual(
            [model["name"] for model in models],
            ["Zero Order", "First Order", "Higuchi", "Hixson–Crowell", "Korsmeyer–Peppas"],
        )
        zero_order = models[0]
        self.assertTrue(zero_order["success"])
        self.assertAlmostEqual(zero_order["r_squared"], 1.0, places=12)
        self.assertEqual(
            best["r_squared"],
            max(model["r_squared"] for model in models if model["success"] and model["r_squared"] is not None),
        )
        self.assertTrue(best["best"])
        self.assertTrue(all(
            model["r_squared"] is None or math.isfinite(model["r_squared"])
            for model in models if model["success"]
        ))

    def test_hixson_crowell_and_korsmeyer_peppas_fit_their_declared_curves(self):
        hixson_points = [
            (time, 100 * (1 - max(0, 1 - 0.08 * time) ** 3))
            for time in (0, 1, 2, 3, 5, 7)
        ]
        power_points = [
            (time, 2.5 * time ** 0.62)
            for time in (0, 0.5, 1, 2, 4, 8)
        ]

        hixson, hixson_error = _fit_kinetic_models(hixson_points)
        power, power_best = _fit_kinetic_models(power_points)

        hixson_result = next(model for model in hixson if model["model_key"] == "hixson_crowell")
        power_result = next(model for model in power if model["model_key"] == "korsmeyer_peppas")
        self.assertTrue(hixson_result["success"], hixson_error)
        self.assertAlmostEqual(hixson_result["r_squared"], 1.0, places=8)
        self.assertTrue(power_result["success"], power_best)
        self.assertAlmostEqual(power_result["parameters"]["n"], 0.62, places=8)
        self.assertAlmostEqual(power_result["r_squared"], 1.0, places=12)

    def test_model_comparison_handles_insufficient_and_invalid_data(self):
        for points in (
            [(0, 0), (1, 2)],
            [(0, 0), (0, 1), (0, 2)],
            [(0, 1), (1, -2), (2, 3)],
        ):
            with self.subTest(points=points):
                models, best = _fit_kinetic_models(points)
                self.assertIsNone(best)
                self.assertTrue(all(not model["success"] for model in models))

    def test_model_page_fits_a_selected_experiment_and_shows_missing_data_message(self):
        experiment = self.create_experiment(
            "Подгонка",
            [(0, 0), (1, 12), (2, 19), (4, 29)],
        )
        with patch("kinetics.views._build_model_plot", return_value="bW9kZWw="):
            response = self.client.get(
                "/models/",
                {"experiment_id": experiment.pk},
            )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "R∞")
        self.assertContains(response, "Оценка плато")
        self.assertContains(response, "data:image/png;base64,bW9kZWw=")
        self.assertIsNotNone(response.context["model_fit"])
        self.assertEqual(len(response.context["model_fits"]), 5)
        self.assertContains(response, "Лучшая модель по R² среди проверенных")
        self.assertContains(response, "TTP и QTP")

        short = self.create_experiment("Мало точек", [(0, 0), (1, 2)])
        response = self.client.get("/models/", {"experiment_id": short.pk})
        self.assertContains(response, "как минимум три измерения")
        self.assertIsNone(response.context["model_plot"])

    def test_model_page_displays_small_rate_without_rounding_it_to_zero(self):
        scale = 10_000_000
        rate = 0.35 / scale
        experiment = self.create_experiment(
            "small-rate",
            [
                (scale * time, 80 * (1 - math.exp(-0.35 * time)))
                for time in (0, 1, 2, 4, 8)
            ],
        )

        with patch("kinetics.views._build_model_plot", return_value="bW9kZWw="):
            response = self.client.get("/models/", {"experiment_id": experiment.pk})

        self.assertEqual(response.status_code, 200)
        self.assertAlmostEqual(response.context["model_fit"]["rate"], rate, delta=rate * 1e-5)
        display_rate = response.context["model_fit_display"]["rate"]
        self.assertIn("e-08", display_rate)
        self.assertContains(response, display_rate)

    def test_model_plot_failure_is_reported_when_fit_succeeds(self):
        experiment = self.create_experiment(
            "model-plot-failure",
            [(0, 0), (1, 12), (2, 19), (4, 29)],
        )

        with self.assertLogs("kinetics.views", level="ERROR"):
            with patch(
                "kinetics.views._build_model_plot",
                side_effect=RuntimeError("simulated model plotting failure"),
            ):
                response = self.client.get("/models/", {"experiment_id": experiment.pk})

        self.assertEqual(response.status_code, 200)
        self.assertIsNotNone(response.context["model_fit"])
        self.assertTrue(response.context["model_error"])
        self.assertContains(response, response.context["model_error"])

    def test_comparison_uses_real_time_values_for_each_series_and_table(self):
        first = self.create_experiment(
            "Первый",
            [(0, 0), (2, 20), (4, 40)],
        )
        second = self.create_experiment(
            "Второй",
            [(0, 1), (1, 11), (6, 61)],
        )
        with patch("kinetics.views._build_comparison_plot", return_value="Y29tcGFyZQ==") as plot_builder:
            response = self.client.get(
                "/comparison/",
                {"experiment_ids": [first.pk, second.pk]},
            )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Совместный график")
        self.assertContains(response, "data:image/png;base64,Y29tcGFyZQ==")
        self.assertContains(response, f"/experiments/{first.pk}/")
        self.assertContains(response, "единицы Time/Release не сохраняются")
        self.assertEqual(
            response.context["datasets"][0]["time_unit"],
            "по файлу; не задано",
        )
        self.assertEqual(
            [dataset["points"] for dataset in plot_builder.call_args.args[0]],
            [[(0.0, 0.0, None), (2.0, 20.0, None), (4.0, 40.0, None)], [(0.0, 1.0, None), (1.0, 11.0, None), (6.0, 61.0, None)]],
        )
        rows = response.context["measurement_page"].object_list
        self.assertEqual(
            [(row.time, row.release) for row in rows],
            [(0.0, 0.0), (0.0, 1.0), (1.0, 11.0), (2.0, 20.0), (4.0, 40.0), (6.0, 61.0)],
        )

    def test_comparison_reports_selection_and_empty_data_requirements(self):
        first = self.create_experiment("Есть данные", [(0, 1), (1, 2)])
        empty = Experiment.objects.create(name="Пустой опыт", source_file="empty.xlsx")

        response = self.client.get("/comparison/", {"experiment_ids": [first.pk]})
        self.assertContains(response, "как минимум два эксперимента")

        response = self.client.get(
            "/comparison/",
            {"experiment_ids": [first.pk, empty.pk]},
        )
        self.assertContains(response, "минимум из двух выбранных экспериментов")
        self.assertContains(response, "Пустой опыт")
        self.assertIsNone(response.context["comparison_plot"])

    def test_report_shows_experiment_metadata_measurements_and_print_action(self):
        experiment = self.create_experiment(
            "Отчётный опыт",
            [(0, 0), (2, 24), (5, 57)],
            source_file="release-study.xlsx",
        )
        with patch("kinetics.views._get_or_build_plot", return_value="cmVwb3J0"):
            response = self.client.get(
                "/reports/",
                {"experiment_id": experiment.pk},
            )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Отчётный опыт")
        self.assertContains(response, "release-study.xlsx")
        self.assertContains(response, "Печать / сохранить PDF")
        self.assertContains(response, "data:image/png;base64,cmVwb3J0")
        self.assertContains(response, "<td>57,0</td>", html=True)
        self.assertEqual(len(response.context["report_measurements"]), 3)
        self.assertEqual(len(response.context["report_model_fits"]), 5)
        self.assertContains(response, "TTP и QTP")
        self.assertContains(response, "В исходном файле не было повторов")

    def test_experiment_pages_and_history_link_to_the_same_saved_experiment(self):
        experiment = self.create_experiment(
            "Связанный опыт",
            [(0, 5), (1, 20)],
            source_file="linked.xlsx",
        )

        detail = self.client.get(f"/experiments/{experiment.pk}/")
        data = self.client.get("/data/", {"experiment_id": experiment.pk})
        history = self.client.get("/experiments/")

        self.assertEqual(detail.status_code, 200)
        self.assertContains(detail, f"/data/?experiment_id={experiment.pk}")
        self.assertContains(detail, f"/models/?experiment_id={experiment.pk}")
        self.assertContains(detail, f"/reports/?experiment_id={experiment.pk}")
        self.assertEqual(data.context["experiment"].pk, experiment.pk)
        self.assertContains(history, f"/experiments/{experiment.pk}/")
        self.assertContains(history, f"/reports/?experiment_id={experiment.pk}")

    def test_report_with_no_measurements_has_a_clear_empty_state(self):
        experiment = Experiment.objects.create(name="Пустой отчёт", source_file="empty.xlsx")
        response = self.client.get("/reports/", {"experiment_id": experiment.pk})

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "У этого эксперимента нет измерений.")
        self.assertContains(response, "Для графика пока нет корректных измерений.")
        self.assertEqual(response.context["report_summary"]["measurement_count"], 0)


class RenameExperimentTimestampMigrationTests(TransactionTestCase):
    """Проверяет, что RenameField сохраняет значение старой колонки даты."""

    migrate_from = [("kinetics", "0002_measurement")]
    migrate_to = [("kinetics", "0003_rename_update_at_and_add_upload_token")]

    def setUp(self):
        super().setUp()
        executor = MigrationExecutor(connection)
        executor.migrate(self.migrate_from)
        old_apps = executor.loader.project_state(self.migrate_from).apps
        old_experiment = old_apps.get_model("kinetics", "Experiment")
        old_measurement = old_apps.get_model("kinetics", "Measurement")
        created = old_experiment.objects.create(
            name="Сохранённая дата",
            description="",
            source_file="date.xlsx",
        )
        self.experiment_pk = created.pk
        old_measurement.objects.create(
            experiment_id=created.pk,
            time=1.25,
            release=2.5,
            order=0,
        )
        self.expected_updated_at = datetime(
            2024, 3, 4, 5, 6, 7, tzinfo=datetime_timezone.utc,
        )
        old_experiment.objects.filter(pk=created.pk).update(
            update_at=self.expected_updated_at,
        )

        executor = MigrationExecutor(connection)
        executor.migrate(self.migrate_to)

    def tearDown(self):
        executor = MigrationExecutor(connection)
        executor.migrate(executor.loader.graph.leaf_nodes())
        super().tearDown()

    def test_timestamp_value_survives_the_column_rename(self):
        migrated_apps = MigrationExecutor(connection).loader.project_state(
            self.migrate_to,
        ).apps
        migrated_experiment = migrated_apps.get_model("kinetics", "Experiment")
        migrated_measurement = migrated_apps.get_model("kinetics", "Measurement")
        experiment = migrated_experiment.objects.get(pk=self.experiment_pk)
        self.assertEqual(experiment.updated_at, self.expected_updated_at)
        self.assertIsNone(experiment.upload_token)
        self.assertEqual(
            list(migrated_measurement.objects.filter(
                experiment_id=self.experiment_pk,
            ).values_list("time", "release", "order")),
            [(1.25, 2.5, 0)],
        )
        with connection.cursor() as cursor:
            columns = {
                row[1] for row in cursor.execute(
                    "PRAGMA table_info(kinetics_measurement)",
                ).fetchall()
            }
            self.assertNotIn("sd", columns)
            self.assertEqual(cursor.execute("PRAGMA foreign_key_check").fetchall(), [])

    def test_rollback_restores_old_column_name_and_keeps_timestamp(self):
        current_apps = MigrationExecutor(connection).loader.project_state(
            self.migrate_to,
        ).apps
        current_experiment = current_apps.get_model("kinetics", "Experiment")
        current_measurement = current_apps.get_model("kinetics", "Measurement")
        experiment = current_experiment.objects.create(
            name="Дата после отката",
            source_file="rollback.xlsx",
        )
        expected = datetime(
            2023, 8, 9, 10, 11, 12, tzinfo=datetime_timezone.utc,
        )
        current_experiment.objects.filter(pk=experiment.pk).update(updated_at=expected)
        current_measurement.objects.create(
            experiment_id=experiment.pk,
            time=3.5,
            release=4.5,
            order=0,
        )

        executor = MigrationExecutor(connection)
        executor.migrate(self.migrate_from)
        old_apps = executor.loader.project_state(self.migrate_from).apps
        old_experiment = old_apps.get_model("kinetics", "Experiment")
        old_measurement = old_apps.get_model("kinetics", "Measurement")

        restored = old_experiment.objects.get(pk=experiment.pk)
        self.assertEqual(restored.update_at, expected)
        self.assertEqual(
            list(old_measurement.objects.filter(experiment_id=experiment.pk).values_list(
                "time", "release", "order",
            )),
            [(3.5, 4.5, 0)],
        )
        with connection.cursor() as cursor:
            columns = {
                row[1] for row in cursor.execute(
                    "PRAGMA table_info(kinetics_experiment)",
                ).fetchall()
            }
            self.assertIn("update_at", columns)
            self.assertNotIn("updated_at", columns)
            self.assertEqual(cursor.execute("PRAGMA foreign_key_check").fetchall(), [])


class MeasurementSdMigrationTests(TransactionTestCase):
    """Проверяет добавление nullable SD без потери исторических измерений."""

    migrate_from = [("kinetics", "0003_rename_update_at_and_add_upload_token")]
    migrate_to = [("kinetics", "0004_measurement_sd")]

    def setUp(self):
        super().setUp()
        executor = MigrationExecutor(connection)
        executor.migrate(self.migrate_from)
        old_apps = executor.loader.project_state(self.migrate_from).apps
        old_experiment = old_apps.get_model("kinetics", "Experiment")
        old_measurement = old_apps.get_model("kinetics", "Measurement")
        experiment = old_experiment.objects.create(
            name="Измерение до SD",
            description="",
            source_file="before-sd.xlsx",
        )
        self.experiment_pk = experiment.pk
        old_measurement.objects.create(
            experiment_id=experiment.pk,
            time=1.5,
            release=12.25,
            order=0,
        )

        executor = MigrationExecutor(connection)
        executor.migrate(self.migrate_to)
        self.new_measurement = executor.loader.project_state(
            self.migrate_to,
        ).apps.get_model("kinetics", "Measurement")

    def tearDown(self):
        executor = MigrationExecutor(connection)
        executor.migrate(executor.loader.graph.leaf_nodes())
        super().tearDown()

    def test_sd_migration_preserves_rows_and_keeps_sd_nullable(self):
        measurement = self.new_measurement.objects.get(
            experiment_id=self.experiment_pk,
        )
        self.assertEqual(
            (measurement.time, measurement.release, measurement.order),
            (1.5, 12.25, 0),
        )
        self.assertIsNone(measurement.sd)

        with_sd = self.new_measurement.objects.create(
            experiment_id=self.experiment_pk,
            time=2.5,
            release=24.5,
            order=1,
            sd=1.25,
        )
        self.assertEqual(with_sd.sd, 1.25)
