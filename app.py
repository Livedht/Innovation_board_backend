from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
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
    
    task_dict = {task['id']: task for task in tasks}
    
    reordered_tasks = []
    for task_id in new_order:
        if task_id in task_dict:
            reordered_tasks.append(task_dict[task_id])
    
    for task in tasks:
        if task['id'] not in new_order:
            reordered_tasks.append(task)
    
    tasks = reordered_tasks
    update_case_numbers()
    return jsonify(tasks)

@app.route('/tasks/generate_report', methods=['GET'])
def generate_report():
    document = Document()
    
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    title = document.add_paragraph("Innovation Board Executive")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.bold = True
    title_run.font.size = Pt(14)

    meeting_date = datetime.now().strftime("%d.%m.%Y")
    details = [
        ("Place", "A4Y-117"),
        ("Date and time", meeting_date)
    ]
    for key, value in details:
        p = document.add_paragraph()
        p.add_run(f"{key}\n").bold = True
        p.add_run(value)

    members = [
        ("Lars Olsen", "Dean Executive"),
        ("Cecilie Asting", "Associate Dean Bachelor of Management"),
        ("Geir Høidal Bjønnes", "Associate Dean Master of Management"),
        ("Lise Hammergren", "Executive Vice President, BI Executive"),
        ("Gry Varre", "Manager Open Enrolment"),
        ("Wenche Martinussen", "Director Sales/Marketing"),
        ("Tonje Omland", "Manager Programme administration - Executive"),
        ("Jørgen Bjørnson Aanderaa", "Director Learning Center"),
        ("Line Lervik-Olsen", "Head of Department, Marketing"),
        ("John Christian Langli", "Head of Department, Accounting and Operations Management")
    ]
    
    document.add_paragraph("Members", style='Heading 1')
    for name, title in members:
        document.add_paragraph(f"{name} {title}")

    document.add_paragraph("Programme Administration / Secretariat", style='Heading 1')
    secretariat = [
        ("Nora Iversen Røed", "Adviser"),
        ("Haakon Tveter", "Senior Adviser")
    ]
    for name, title in secretariat:
        document.add_paragraph(f"{name} {title}")

    # Add a page break before the agenda
    document.add_page_break()

    document.add_paragraph("Agenda", style='Heading 1')
    for index, task in enumerate(tasks, start=1):
        document.add_paragraph(f"{task['caseNumber']} {task['title']} page {index}")

    for task in tasks:
        # Add a page break before each proposal
        document.add_page_break()

        document.add_paragraph("Proposal – New course idea", style='Heading 2')
        
        # Make headings bold and add content
        headings = [
            ("Item number", task['caseNumber']),
            ("Idea title", task['title']),
            ("Idea owner", task['owner']),
            ("Briefly describe the idea", task['description']),
            ("Why is this relevant for BI?", task.get('relevanceForBI', '')),
            ("Why does individuals and/or organizations need such a course/idea?", task.get('needForCourse', '')),
            ("What would be the relevant target group?", task.get('targetGroup', '')),
            ("What are your thoughts on the future growth potential of the market for this course/idea?", task.get('growthPotential', '')),
            ("Faculty resources – which academic departments should be involved?", task.get('facultyResources', ''))
        ]

        for heading, content in headings:
            document.add_paragraph(heading, style='Heading 3')
            document.add_paragraph(content)

    f = BytesIO()
    document.save(f)
    f.seek(0)

    return send_file(f, as_attachment=True, download_name='innovation_board_sakspapirer.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    app.run(debug=True)