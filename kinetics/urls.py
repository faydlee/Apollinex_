from django.urls import path
from . import auth_views
from . import views

# Маршруты обзора, таблицы, истории и сохранённых экспериментов.
urlpatterns = [
    path('accounts/login/', auth_views.LoginView.as_view(), name='login'),
    path('accounts/register/', auth_views.RegistrationView.as_view(), name='register'),
    path('accounts/logout/', auth_views.LogoutView.as_view(), name='logout'),
    path('settings/', auth_views.settings_page, name='settings'),
    # Обзор эксперимента.
    path('', views.home, name='home'),
    path('method/', views.method_documentation, name='method_documentation'),
    # Таблица данных выбранного или последнего сохранённого эксперимента.
    path('data/', views.data_page, name='data'),
    # Список сохранённых экспериментов.
    path('experiments/', views.experiment_history, name='experiment_history'),
    # Эмпирическая подгонка модели по данным выбранного эксперимента.
    path('models/', views.kinetic_models, name='kinetic_models'),
    # Совместный график и таблица выбранных экспериментов.
    path('comparison/', views.comparison, name='comparison'),
    path(
        'comparison/export.xlsx',
        views.comparison_export_xlsx,
        name='comparison_export_xlsx',
    ),
    # Печатная страница отчёта выбранного эксперимента.
    path('reports/', views.reports, name='reports'),
    # Просмотр сохранённого эксперимента, его точек и графика.
    path('experiments/<int:pk>/', views.experiment_detail, name='experiment_detail'),
    path(
        'experiments/<int:pk>/export.xlsx',
        views.export_experiment_xlsx,
        name='experiment_export_xlsx',
    ),
    path('experiments/new/raw/', views.raw_experiment_new, name='raw_experiment_new'),
    path(
        'experiments/<int:pk>/edit/',
        views.raw_experiment_edit,
        name='experiment_edit',
    ),
    path(
        'experiments/<int:pk>/recalculate/',
        views.recalculate_raw_experiment,
        name='experiment_recalculate',
    ),
    path(
        'experiments/<int:pk>/transforms/<slug:model_key>.csv',
        views.source_transform_csv,
        name='source_transform_csv',
    ),
    path(
        'experiments/<int:pk>/imported-transforms/<slug:model_key>.csv',
        views.imported_model_points_csv,
        name='imported_model_points_csv',
    ),
    # Удаление допускается только POST-запросом.
    path('experiments/<int:pk>/delete/', views.delete_experiment, name='experiment_delete'),
    # Приём загружаемой книги .xlsx.
    path('upload/', views.upload, name='upload'),
    path(
        'upload/csharp-raw/',
        views.upload_csharp_workbook,
        name='upload_csharp_workbook',
    ),
]
