from django.urls import path
from . import views

urlpatterns = [
    path('', views.index_view, name='index'),
    path('api/start/', views.start_scraping_view, name='start_scraping'),
    path('api/status/<str:task_id>/', views.check_status_view, name='check_status'),
    path('api/download/<str:task_id>/', views.download_report_view, name='download_report'),
]
