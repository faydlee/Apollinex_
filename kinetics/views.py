from django.shortcuts import render
from django.http import HttpResponse
from django.core.cache import cache
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile
import traceback
import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt
import io
import base64
import hashlib
import json
import math


SESSION_DATA_KEY = 'kinetics_data'
SESSION_FILENAME_KEY = 'kinetics_filename'


def _build_plot(data):
    times = [row['time'] for row in data]
    releases = [row['release'] for row in data]
    fig, ax = plt.subplots()
    image_buffer = io.BytesIO()
    try:
        ax.plot(times, releases, marker='o')
        ax.set_xlabel('Время')
        ax.set_ylabel('Высвобождение (%)')
        ax.set_title('Кинетика высвобождения')
        ax.grid(True)
        fig.savefig(image_buffer, format='png', bbox_inches='tight')
        image_buffer.seek(0)
        return base64.b64encode(image_buffer.getvalue()).decode('utf-8')
    finally:
        plt.close(fig)
        image_buffer.close()


def _plot_cache_key(data):
    def normalize(value):
        if isinstance(value, float) and not math.isfinite(value):
            if math.isnan(value):
                special_value = 'nan'
            else:
                special_value = 'inf' if value > 0 else '-inf'
            return {'__float__': special_value}
        return value

    payload = json.dumps(
        [[normalize(row['time']), normalize(row['release'])] for row in data],
        separators=(',', ':'),
        ensure_ascii=False,
        allow_nan=False,
    ).encode('utf-8')
    digest = hashlib.sha256(payload).hexdigest()
    return f'kinetics:plot:{digest}'


def _get_or_build_plot(data):
    key = _plot_cache_key(data)
    plot = cache.get(key)
    if plot is None:
        plot = _build_plot(data)
        cache.set(key, plot, timeout=None)
    return plot


def _get_saved_experiment(request):
    data = request.session.get(SESSION_DATA_KEY, [])
    if not isinstance(data, list):
        data = []
    return {
        'filename': request.session.get(SESSION_FILENAME_KEY, ''),
        'data': data,
        'row_count': len(data),
    }


def home(request):
    context = _get_saved_experiment(request)
    data = context['data']
    if data:
        times = [row['time'] for row in data]
        releases = [row['release'] for row in data]
        context.update({
            'time_min': min(times),
            'time_max': max(times),
            'max_release': max(releases),
            'plot': _get_or_build_plot(data),
        })
    return render(request, 'kinetics/home.html', context)

def data_page(request):
    return render(
        request,
        'kinetics/data.html',
        _get_saved_experiment(request),
    )

    
def upload(request):
    if request.method != 'POST':
        return HttpResponse('Используйте форму для загрузки файла.')
    
    excel_file = request.FILES.get('file')
    
    if not excel_file:
        return render(request, 'kinetics/home.html',{
            'error': 'Выберите Excel-файл для загрузки.'
        })

    if not excel_file.name.lower().endswith('.xlsx'):
        return render(request, 'kinetics/home.html', {
            'error': 'Поддерживается только формат .xlsx.'
        })

    workbook = None
    try:
        workbook = load_workbook(
            excel_file,
            read_only=True,
            data_only=True
            )
    
        sheet = workbook.active
        rows = list(sheet.iter_rows(values_only=True))
        
        if len(rows) < 2:
            return render(request, 'kinetics/home.html', {
                'error': 'Файл должен содержать заголовки и хотя бы одну строку данных.'
            })
            
        if (
            len(rows[0]) < 2
            or rows[0][0] is None
            or rows[0][1] is None
            or not str(rows[0][0]).strip()
            or not str(rows[0][1]).strip()
        ):
            return render(request, 'kinetics/home.html', {
                'error': 'В файле должны быть минимум два столбца.'
            })
            
        data = []
            
        for row_number, row in enumerate(rows[1:], start=2):
            if len(row) < 2 or row[0] is None or row[1] is None:
                return render(request, 'kinetics/home.html', {
                    'error': f'Пустое значение в строке {row_number}.'
                })
                
            if not isinstance(row[0], (int, float)):
                return render(request, 'kinetics/home.html', {
                    'error': f'Время в строке {row_number}: должно быть числом.'
                })
                
            if not isinstance(row[1], (int, float)):
                return render(request, 'kinetics/home.html', {
                    'error': f'Высвобождение в строке {row_number}: должно быть числом.'
            })
                
            data.append({
                'time': row[0],
                'release': row[1]
            })

        times = [row['time'] for row in data]
        releases = [row['release'] for row in data]
        time_min = min(times)
        time_max = max(times)
        max_release = max(releases)
        plot = _get_or_build_plot(data)

        request.session[SESSION_FILENAME_KEY] = excel_file.name
        request.session[SESSION_DATA_KEY] = data
            
        return render(request, 'kinetics/home.html', {
            'filename': excel_file.name,
            'row_count': len(data),
            'data': data,
            'time_min': time_min,
            'time_max': time_max,
            'max_release': max_release,
            'success': 'Файл успешно загружен и проверен.',
            'plot': plot
        })
        
    except (BadZipFile, InvalidFileException):
        return render(request, 'kinetics/home.html', {
            'error': 'Файл поврежден или не является корректным файлом .xlsx.'
        })
    except Exception:
        traceback.print_exc()
        return render(request, 'kinetics/home.html', {
            'error': 'Не удалось прочитать файл. Подробности находятся в терминале Django.'
        })
    finally:
        if workbook is not None:
            workbook.close()
        
