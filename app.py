from flask import Flask, jsonify, request
from flask_cors import CORS
from datetime import datetime
from mock_data import mock_tasks

app = Flask(__name__)
CORS(app)

tasks = mock_tasks
current_year = datetime.now().year
current_meeting = 1

def update_case_numbers():
    global tasks
    for index, task in enumerate(tasks, start=1):
        task['caseNumber'] = f"{index:02d}/{str(current_year)[2:]} - {roman_numeral(current_meeting)}"

def roman_numeral(num):
    roman = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
    return roman[num - 1] if num <= 10 else str(num)

@app.route('/tasks', methods=['GET'])
def get_tasks():
    return jsonify(tasks)

@app.route('/tasks', methods=['POST'])
def add_task():
    new_task = request.json
    new_task['id'] = max([task['id'] for task in tasks]) + 1 if tasks else 1
    new_task['stage'] = "Idea Description"  # Set default stage
    tasks.append(new_task)
    update_case_numbers()
    return jsonify(new_task), 201

@app.route('/tasks/<int:task_id>', methods=['PUT'])
def update_task(task_id):
    task = next((task for task in tasks if task['id'] == task_id), None)
    if task is None:
        return jsonify({'error': 'Task not found'}), 404

    updated_data = request.json
    task.update(updated_data)
    return jsonify(task)

@app.route('/tasks/<int:task_id>', methods=['DELETE'])
def delete_task(task_id):
    global tasks
    tasks = [task for task in tasks if task['id'] != task_id]
    update_case_numbers()
    return '', 204

@app.route('/tasks/reorder', methods=['POST'])
def reorder_tasks():
    new_order = request.json
    global tasks
    
    # Create a dictionary mapping task IDs to tasks
    task_dict = {task['id']: task for task in tasks}
    
    # Create a new list of tasks based on the new order
    reordered_tasks = []
    for task_id in new_order:
        if task_id in task_dict:
            reordered_tasks.append(task_dict[task_id])
    
    # Add any tasks that weren't in the new_order to the end
    for task in tasks:
        if task['id'] not in new_order:
            reordered_tasks.append(task)
    
    tasks = reordered_tasks
    update_case_numbers()
    return jsonify(tasks)

@app.route('/tasks/generate_report', methods=['GET'])
def generate_report():
    report = "Sakspapirer\n\n"
    for task in tasks:
        report += f"Sak {task['caseNumber']}: {task['title']}\n"
        report += f"Idéeier: {task['owner']}\n"
        report += f"Beskrivelse: {task['description']}\n"
        report += f"Relevans for BI: {task['relevanceForBI']}\n"
        report += f"Behov for kurset/idéen: {task['needForCourse']}\n"
        report += f"Målgruppe: {task['targetGroup']}\n"
        report += f"Vekstpotensial: {task['growthPotential']}\n"
        report += f"Faglige ressurser: {task['facultyResources']}\n\n"
    return jsonify({"report": report})

if __name__ == '__main__':
    app.run(debug=True)
