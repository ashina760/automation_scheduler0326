import threading
import subprocess
import queue
import re
import datetime
import time

task_queue = queue.PriorityQueue()
tasks = []
task_lock = threading.Lock()

def execute_task(script_name, task):
    """ 执行子进程任务 """
    try:
        import_path = task.get('import_path')  # 获取文件路径
        command = ['python', '-u', script_name]
        if import_path:
            command.append(import_path)

        with task_lock:  # 确保任务串行执行
            process = subprocess.Popen(command,  
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, bufsize=1)

            destination_folder = None
            while True:
                retcode = process.poll()  # 检查进程是否结束
                line = process.stdout.readline().strip()  # 读取标准输出
                err_line = process.stderr.readline().strip()  # 读取标准错误

                if line:
                    print(f"子进程输出: {line}")
                    task['status'] = line  # 更新状态
                    match = re.search(r"DESTINATION_FOLDER:\s*(.+)", line)
                    if match:
                        destination_folder = match.group(1)
                        task['dwload_path'] = destination_folder  # 存入任务
                
                if err_line:
                    print(f"子进程错误: {err_line}")  # 处理错误信息
                
                if retcode is not None:  # 进程结束，跳出循环
                    break  

            print("子进程已结束")

    except Exception as e:
        task['status'] = f'执行失败: {str(e)}'
        print(f"任务执行失败: {str(e)}")

def task_worker():
    """ 任务队列处理线程 """
    while True:
        task_time, task = task_queue.get()
        script_name = task['script']

        # 如果任务已被删除，则跳过执行
        if task.get('cancelled', False):
            print(f"任务 {script_name} 已被删除，跳过执行")
            task_queue.task_done()
            continue

        task_time = task['time']  
        now = datetime.datetime.now().replace(microsecond=0)
        task_time = task_time.replace(tzinfo=None)
        remaining_time = (task_time - now).total_seconds()

        if remaining_time > 0:
            print(f"任务 {script_name} 还未到执行时间，等待 {remaining_time} 秒...")
            time.sleep(min(remaining_time, 60))  
            task_queue.put((task_time, task))  # 重新放入队列
        else:
            print(f"Executing task: {task}")
            execute_task(script_name, task)
            print(f"Finish task: {task}")

        task_queue.task_done()
        
# 启动后台线程，监听任务队列
threading.Thread(target=task_worker, daemon=True).start()
