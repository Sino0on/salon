import uuid
import threading
import os
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
import logging
from .services import run as run_scraper

logger = logging.getLogger(__name__)

from django.core.cache import cache

def update_task(task_id, updates):
    task = cache.get(task_id, {})
    task.update(updates)
    cache.set(task_id, task, timeout=86400)
def index_view(request):
    return render(request, 'scraper/index.html')

def background_task(task_id, api_key, regions, keywords):
    logger.info(f"Начало фоновой задачи {task_id}. Регионы: {regions}")
    update_task(task_id, {'status': 'running'})
    output_filename = f"report_{task_id}.xlsx"
    output_path = os.path.join("scraper_reports", output_filename)
    os.makedirs("scraper_reports", exist_ok=True)
    
    def status_callback(msg):
        update_task(task_id, {'message': msg})

    try:
        run_scraper(api_key, regions, keywords, output_file=output_path, status_callback=status_callback)
        update_task(task_id, {'status': 'completed', 'file_path': output_path})
        logger.info(f"Задача {task_id} успешно завершена. Файл: {output_path}")
    except Exception as e:
        update_task(task_id, {'status': 'error', 'message': str(e)})
        logger.error(f"Ошибка в задаче {task_id}: {e}", exc_info=True)

@csrf_exempt
def start_scraping_view(request):
    if request.method == 'POST':
        api_key = request.POST.get('api_key')
        regions_raw = request.POST.get('regions', '')
        keywords_raw = request.POST.get('keywords', '')
        
        regions = [r.strip() for r in regions_raw.split(',') if r.strip()]
        keywords = [k.strip() for k in keywords_raw.split(',') if k.strip()]
        
        if not api_key or not regions or not keywords:
             return JsonResponse({'error': 'Все поля должны быть заполнены'}, status=400)

        task_id = str(uuid.uuid4())
        logger.info(f"Создана новая задача {task_id} от пользователя")
        cache.set(task_id, {'status': 'pending', 'message': 'Инициализация парсера...'}, timeout=86400)
        
        thread = threading.Thread(target=background_task, args=(task_id, api_key, regions, keywords))
        thread.start()
        
        return JsonResponse({'task_id': task_id})
    return JsonResponse({'error': 'Invalid method'}, status=400)

def check_status_view(request, task_id):
    task = cache.get(task_id)
    if task:
        return JsonResponse(task)
    return JsonResponse({'error': 'Task not found'}, status=404)

def download_report_view(request, task_id):
    task = cache.get(task_id)
    if task and task['status'] == 'completed':
        file_path = task['file_path']
        if os.path.exists(file_path):
            return FileResponse(open(file_path, 'rb'), as_attachment=True, filename="beauty_salons_report.xlsx")
    return JsonResponse({'error': 'File not found'}, status=404)
