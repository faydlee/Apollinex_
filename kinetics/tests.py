from io import BytesIO

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from openpyxl import Workbook


class ExperimentDataPageTests(TestCase):
    def upload_workbook(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['Time', 'Release'])
        sheet.append([0, 0])
        sheet.append([2, 12.4])
        sheet.append([4, 28.1])
        content = BytesIO()
        workbook.save(content)
        return SimpleUploadedFile(
            'experiment.xlsx',
            content.getvalue(),
            content_type=(
                'application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.sheet'
            ),
        )

    def test_uploaded_data_is_available_on_data_page(self):
        response = self.client.post(
            '/upload/',
            {'file': self.upload_workbook()},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'data:image/png;base64,')
        self.assertRegex(
            response.content.decode(),
            r'<a class="nav-item" href="/#overview" aria-current="page">',
        )

        data_response = self.client.get('/data/')

        self.assertEqual(data_response.status_code, 200)
        self.assertContains(data_response, 'experiment.xlsx')
        self.assertContains(data_response, 'Количество строк')
        self.assertContains(data_response, '<td>12.4</td>', html=True)
        self.assertContains(data_response, '<td>28.1</td>', html=True)

    def test_home_page_keeps_summary_and_chart_without_full_table(self):
        self.client.post('/upload/', {'file': self.upload_workbook()})
        self.client.get('/data/')

        response = self.client.get('/')

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'experiment.xlsx')
        self.assertContains(response, 'Точек данных')
        self.assertNotContains(response, '<td>12.4</td>')
        self.assertContains(response, 'data:image/png;base64,')

    def test_data_page_shows_empty_state_before_upload(self):
        response = self.client.get('/data/')

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Данные пока не загружены')

    def test_sidebar_highlights_only_the_current_implemented_page(self):
        home_response = self.client.get('/')
        home_html = home_response.content.decode()

        self.assertEqual(home_response.status_code, 200)
        self.assertRegex(
            home_html,
            r'<a class="nav-item" href="/#overview" aria-current="page">',
        )
        self.assertEqual(home_html.count('aria-current="page">'), 1)

        data_response = self.client.get('/data/')
        data_html = data_response.content.decode()

        self.assertEqual(data_response.status_code, 200)
        self.assertRegex(
            data_html,
            r'<a class="nav-item" href="/data/" aria-current="page">',
        )
        self.assertEqual(data_html.count('aria-current="page">'), 1)
        self.assertEqual(data_html.count('href="/data/"'), 1)
        self.assertContains(data_response, 'Графики')
        self.assertContains(data_response, 'aria-disabled="true"')
