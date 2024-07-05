from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import requests
import logging
import os

app = Flask(__name__)
CORS(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG)

tasks = []
meetings = []
current_year = datetime.now().year
current_meeting = 1

NETTSKJEMA_API_URL = "https://nettskjema.no/api/v2"
NETTSKJEMA_TOKEN = "vd3vtss32mf08me2vh584q6nb2i2b8t0icoao4dablv0b4r8crji5o5lrgl5ii4q5hhanmbpan4phdf03biglqglvtj9igsa609oc7mut4n7etk210e3hra1rb8segs6mq6q"
NETTSKJEMA_FORM_ID = "397131"


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

# Modify this function to accept a task as an argument
def add_task_internal(task):
    task['id'] = max([t['id'] for t in tasks]) + 1 if tasks else 1
    task['stage'] = "Idea Description"  # Set default stage
    tasks.append(task)
    update_case_numbers()
    return task

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

@app.route('/meetings', methods=['GET'])
def get_meetings():
    return jsonify(meetings)

@app.route('/meetings', methods=['POST'])
def add_meeting():
    new_meeting = request.json
    new_meeting['id'] = max([meeting['id'] for meeting in meetings]) + 1 if meetings else 1
    new_meeting['tasks'] = []
    meetings.append(new_meeting)
    return jsonify(new_meeting), 201

@app.route('/meetings/<int:meeting_id>', methods=['PUT'])
def update_meeting(meeting_id):
    meeting = next((meeting for meeting in meetings if meeting['id'] == meeting_id), None)
    if meeting is None:
        return jsonify({'error': 'Meeting not found'}), 404

    updated_data = request.json
    meeting.update(updated_data)
    return jsonify(meeting)

@app.route('/meetings/<int:meeting_id>', methods=['DELETE'])
def delete_meeting(meeting_id):
    global meetings
    meetings = [meeting for meeting in meetings if meeting['id'] != meeting_id]
    return '', 204

@app.route('/meetings/<int:meeting_id>/tasks', methods=['POST'])
def add_task_to_meeting(meeting_id):
    meeting = next((meeting for meeting in meetings if meeting['id'] == meeting_id), None)
    if meeting is None:
        return jsonify({'error': 'Meeting not found'}), 404

    task_id = request.json['task_id']
    task = next((task for task in tasks if task['id'] == task_id), None)
    if task is None:
        return jsonify({'error': 'Task not found'}), 404

    meeting['tasks'].append(task)
    return jsonify(meeting)

@app.route('/meetings/<int:meeting_id>/generate_report', methods=['GET'])
def generate_report(meeting_id):
    meeting = next((meeting for meeting in meetings if meeting['id'] == meeting_id), None)
    if meeting is None:
        return jsonify({'error': 'Meeting not found'}), 404

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

    meeting_date = meeting['date']
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

    document.add_page_break()

    document.add_paragraph("Agenda", style='Heading 1')
    for index, task in enumerate(meeting['tasks'], start=1):
        document.add_paragraph(f"{task['caseNumber']} {task['title']} page {index}")

    for task in meeting['tasks']:
        document.add_page_break()

        document.add_paragraph("Proposal – New course idea", style='Heading 2')
        
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

def get_nettskjema_data(form_id):
    headers = {
        "Authorization": f"Bearer {NETTSKJEMA_TOKEN}"
    }
    response = requests.get(f"{NETTSKJEMA_API_URL}/forms/{form_id}/submissions", headers=headers)
    response.raise_for_status()
    return response.json()

def transform_submission_to_task(submission):
    logging.debug(f"Submission structure: {submission}")
    
    answers = {}
    for answer in submission.get('answers', []):
        question_id = str(answer.get('questionId'))
        answer_value = answer.get('textAnswer', '')
        answers[question_id] = answer_value
    
    return {
        "submissionId": submission.get('submissionId'),  # Unique identifier
        "title": answers.get('7168624', f"Submission {submission.get('submissionId', 'Unknown')}"),
        "owner": answers.get('6795267', submission.get('respondentEmail', 'Unknown')),
        "description": answers.get('6795369', f"Imported from Nettskjema submission {submission.get('submissionId', 'Unknown')}"),
        "relevanceForBI": answers.get('6795370', ''),
        "needForCourse": answers.get('6795371', ''),
        "targetGroup": answers.get('6795373', ''),
        "growthPotential": answers.get('6826847', ''),
        "facultyResources": answers.get('6795374', ''),
        "stage": "Idea Description",
    }

@app.route('/meetings/<int:meeting_id>/tasks/<int:task_id>', methods=['DELETE'])
def remove_task_from_meeting(meeting_id, task_id):
    meeting = next((meeting for meeting in meetings if meeting['id'] == meeting_id), None)
    if meeting is None:
        return jsonify({'error': 'Meeting not found'}), 404

    meeting['tasks'] = [task for task in meeting['tasks'] if task['id'] != task_id]
    return jsonify(meeting)



@app.route('/import-nettskjema', methods=['POST'])
def import_nettskjema():
    try:
        logging.info("Starting Nettskjema import process")
        submissions = get_nettskjema_data(NETTSKJEMA_FORM_ID)
        imported_tasks = []

        existing_submission_ids = {task.get('submissionId') for task in tasks}

        for submission in submissions:
            if submission.get('submissionId') in existing_submission_ids:
                logging.info(f"Skipping already existing submission with ID: {submission.get('submissionId')}")
                continue
            
            try:
                task = transform_submission_to_task(submission)
                new_task = add_task_internal(task)  # Use the new internal function
                imported_tasks.append(new_task)
            except Exception as e:
                logging.error(f"Error processing submission: {str(e)}")
                continue

        logging.info(f"Successfully imported {len(imported_tasks)} tasks")
        return jsonify({
            "message": f"Successfully imported {len(imported_tasks)} tasks",
            "imported_tasks": imported_tasks
        }), 200
    except requests.RequestException as e:
        logging.error(f"Error fetching data from Nettskjema: {str(e)}")
        return jsonify({"error": f"Error fetching data from Nettskjema: {str(e)}"}), 500
    except Exception as e:
        logging.error(f"Error processing import: {str(e)}")
        return jsonify({"error": f"Error processing import: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
