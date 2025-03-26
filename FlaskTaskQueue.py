import json
import subprocess
import os
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from task_manager import task_queue, tasks
from datetime import datetime


app = Flask(__name__)

# 读取 config.json
CONFIG_PATH = r"C:\Users\beite.dai.bp\OneDrive - Coca-Cola Bottlers Japan\デスクトップ\py\automation_scheduler\config.json"
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)

app.secret_key = config["secret_key"]
scripts = config["scripts"]
opload_folder = config["opload_folder"]
download_folder = config["download_folder"]
destination_folder = config["destination_folder"]

# 确保目录存在
for folder in [opload_folder, destination_folder]:
    os.makedirs(folder, exist_ok=True)

# 执行任务
@app.route('/run_now', methods=['POST'])
def run_now():
    script_name = request.form.get('script_name')
    execute_time = request.form.get('execute_time')
    file = request.files.get('file')
    user = request.form.get('user')  # 获取用户 ID 
    import_path = None

    if file and file.filename:
        import_path = os.path.join(opload_folder, file.filename)
        file.save(import_path)

    execute_time_dt = datetime.fromisoformat(execute_time)
    if script_name and execute_time_dt:
        job_id = len(tasks) + 1
        task = {
            'script': script_name,
            'time': execute_time_dt,
            'status': 'Wait',
            'job_id': job_id,
            'import_path': import_path,
            "user": user 
        }
        task_queue.put((execute_time_dt, task))
        tasks.append(task)
        return jsonify({'status': 'success', 'message': f'任务 {script_name} 已成功添加！'})
    else:
        return jsonify({'status': 'error', 'message': '未填写脚本名称或执行时间！'})
    
# 删除任务
@app.route('/delete_task/<int:job_id>')
def delete_task(job_id):
    global tasks
    for task in tasks:
        if task['job_id'] == job_id:
            task['cancelled'] = True  # 标记任务已删除
            task['status'] = 'Termination'  # 修改状态
    flash('Task Termination', 'success')
    return redirect(url_for('index'))

# 打开子程序附件的保存路径
@app.route('/open_folder', methods=['POST'])
def open_folder():
    destination_folder = os.path.normpath(next(t['dwload_path'] for t in tasks if t['job_id'] == request.json['job_id']))
    subprocess.Popen(f'explorer "{destination_folder}"', shell=True)
    return jsonify({"status": "success"})

# 主路由
@app.route('/')
def index():
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    correct_id = config.get('correct_id')
    return render_template('index.html', scripts=scripts, tasks=tasks, correct_id=correct_id)

if __name__ == '__main__':
    app.run(debug=True, port=5000)

    # app.run(host="0.0.0.0",debug=True,port=5000)