import uuid
import threading
import os
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from .services import run as run_scraper

# Временное хранилище статусов (в памяти)
tasks = {}

def index_view(request):
    return render(request, 'scraper/index.html')

def background_task(task_id, api_key, regions, keywords):
    tasks[task_id]['status'] = 'running'
    output_filename = f"report_{task_id}.xlsx"
    output_path = os.path.join("scraper_reports", output_filename)
    os.makedirs("scraper_reports", exist_ok=True)
    
    def status_callback(msg):
        tasks[task_id]['message'] = msg

    try:
        run_scraper(api_key, regions, keywords, output_file=output_path, status_callback=status_callback)
        tasks[task_id]['status'] = 'completed'
        tasks[task_id]['file_path'] = output_path
    except Exception as e:
        tasks[task_id]['status'] = 'error'
        tasks[task_id]['message'] = str(e)

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
        tasks[task_id] = {'status': 'pending', 'message': 'Инициализация парсера...'}
        
        thread = threading.Thread(target=background_task, args=(task_id, api_key, regions, keywords))
        thread.start()
        
        return JsonResponse({'task_id': task_id})
    return JsonResponse({'error': 'Invalid method'}, status=400)

def check_status_view(request, task_id):
    task = tasks.get(task_id)
    if task:
        return JsonResponse(task)
    return JsonResponse({'error': 'Task not found'}, status=404)

def download_report_view(request, task_id):
    task = tasks.get(task_id)
    if task and task['status'] == 'completed':
        file_path = task['file_path']
        if os.path.exists(file_path):
            return FileResponse(open(file_path, 'rb'), as_attachment=True, filename="beauty_salons_report.xlsx")
    return JsonResponse({'error': 'File not found'}, status=404)
