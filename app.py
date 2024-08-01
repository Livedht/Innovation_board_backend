import os
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from datetime import datetime, timezone
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import requests
import logging
from supabase import create_client, Client
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session
from dotenv import load_dotenv
import json
import postgrest.exceptions

load_dotenv()

app = Flask(__name__)
CORS(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG)

# Initialize Supabase client
url: str = os.environ.get("SUPABASE_URL")
key: str = os.environ.get("SUPABASE_KEY")
supabase: Client = create_client(url, key)

# Nettskjema API configuration
NETTSKJEMA_API_URL = "https://api.nettskjema.no/v3"
NETTSKJEMA_AUTH_URL = "https://authorization.nettskjema.no/oauth2/token"
NETTSKJEMA_CLIENT_ID = os.environ.get("NETTSKJEMA_CLIENT_ID")
NETTSKJEMA_CLIENT_SECRET = os.environ.get("NETTSKJEMA_CLIENT_SECRET")
NETTSKJEMA_FORM_ID = os.environ.get("NETTSKJEMA_FORM_ID")

current_year = datetime.now().year
current_meeting = 1

def roman_numeral(num):
    roman = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
    return roman[num - 1] if num <= 10 else str(num)

def get_oauth_session():
    client = BackendApplicationClient(client_id=NETTSKJEMA_CLIENT_ID)
    oauth = OAuth2Session(client=client)
    token = oauth.fetch_token(
        token_url=NETTSKJEMA_AUTH_URL,
        client_id=NETTSKJEMA_CLIENT_ID,
        client_secret=NETTSKJEMA_CLIENT_SECRET
    )
    return oauth

def get_nettskjema_data(form_id):
    oauth = get_oauth_session()
    response = oauth.get(f"{NETTSKJEMA_API_URL}/form/{form_id}/answers")
    response.raise_for_status()
    
    # Log the raw response for debugging
    logging.debug(f"Raw response from Nettskjema: {response.text}")
    
    # Parse the response line by line
    data = []
    for line in response.iter_lines():
        if line:
            try:
                json_obj = json.loads(line.decode('utf-8'))
                data.append(json_obj)
            except json.JSONDecodeError as e:
                logging.error(f"Error decoding JSON line: {e}")
                logging.error(f"Problematic line: {line}")
    
    return data

def transform_submission_to_task(submission):
    logging.debug(f"Processing submission: {submission}")
    
    submissions_dict = {}
    for answer in submission:
        submissionId = answer['submissionId']
        if submissionId not in submissions_dict:
            submissions_dict[submissionId] = []
        submissions_dict[submissionId].append(answer)
    
    tasks = []
    for submissionId, answers in submissions_dict.items():
        task = {
            "submission_id": submissionId,
            "title": "",
            "owner": "",
            "description": "",
            "relevance_for_bi": "",
            "need_for_course": "",
            "target_group": "",
            "growth_potential": "",
            "faculty_resources": "",
            "stage": "Idea Description",
        }
        
        for answer in answers:
            elementId = str(answer['elementId'])
            if elementId == '6461993':  # Title
                task['title'] = answer.get('textAnswer', '')
            elif elementId == '6130890':  # Owner
                task['owner'] = answer.get('textAnswer', '')
            elif elementId == '6130957':  # Description
                task['description'] = answer.get('textAnswer', '')
            elif elementId == '6130958':  # Relevance for BI
                task['relevance_for_bi'] = answer.get('textAnswer', '')
            elif elementId == '6130959':  # Need for course
                task['need_for_course'] = answer.get('textAnswer', '')
            elif elementId == '6130961':  # Target group
                task['target_group'] = answer.get('textAnswer', '')
            elif elementId == '6158929':  # Growth potential
                task['growth_potential'] = answer.get('textAnswer', '')
            elif elementId == '6130962':  # Faculty resources
                task['faculty_resources'] = answer.get('textAnswer', '')
        
        tasks.append(task)
    
    return tasks

@app.route('/tasks', methods=['GET'])
def get_tasks():
    response = supabase.table('tasks').select('*').execute()
    tasks = response.data
    for index, task in enumerate(tasks, start=1):
        task['caseNumber'] = f"{index:02d}/{str(current_year)[2:]} - {roman_numeral(current_meeting)}"
    return jsonify(tasks)

@app.route('/tasks', methods=['POST'])
def add_task():
    new_task = request.json
    response = supabase.table('tasks').insert(new_task).execute()
    return jsonify(response.data[0]), 201

@app.route('/tasks/<int:task_id>', methods=['PUT'])
def update_task(task_id):
    updated_data = request.json
    response = supabase.table('tasks').update(updated_data).eq('id', task_id).execute()
    return jsonify(response.data[0])

@app.route('/tasks/<int:task_id>', methods=['DELETE'])
def delete_task(task_id):
    supabase.table('tasks').delete().eq('id', task_id).execute()
    return '', 204

@app.route('/tasks/reorder', methods=['POST'])
def reorder_tasks():
    new_order = request.json
    tasks = supabase.table('tasks').select('*').execute().data
    task_dict = {task['id']: task for task in tasks}
    
    reordered_tasks = []
    for task_id in new_order:
        if task_id in task_dict:
            reordered_tasks.append(task_dict[task_id])
    
    for task in tasks:
        if task['id'] not in new_order:
            reordered_tasks.append(task)
    
    for index, task in enumerate(reordered_tasks, start=1):
        task['caseNumber'] = f"{index:02d}/{str(current_year)[2:]} - {roman_numeral(current_meeting)}"
        supabase.table('tasks').update({'caseNumber': task['caseNumber']}).eq('id', task['id']).execute()
    
    return jsonify(reordered_tasks)

@app.route('/meetings', methods=['GET'])
def get_meetings():
    response = supabase.table('meetings').select('*').execute()
    return jsonify(response.data)

@app.route('/meetings', methods=['POST'])
def add_meeting():
    new_meeting = request.json

    # Ensure only the columns that exist in your database are included
    meeting_data = {
        'number': new_meeting.get('number'),
        'date': new_meeting.get('date')
    }

    # Validate required fields
    if not meeting_data['number'] or not meeting_data['date']:
        return jsonify({'error': 'Meeting number and date are required'}), 400

    # Convert date string to timestamp with timezone if it's not already
    if isinstance(meeting_data['date'], str):
        try:
            meeting_data['date'] = datetime.fromisoformat(meeting_data['date']).replace(tzinfo=timezone.utc)
        except ValueError:
            return jsonify({'error': 'Invalid date format. Please use ISO format (YYYY-MM-DDTHH:MM:SS.sssZ)'}), 400

    # Convert datetime object to ISO format string
    if isinstance(meeting_data['date'], datetime):
        meeting_data['date'] = meeting_data['date'].isoformat()

    try:
        response = supabase.table('meetings').insert(meeting_data).execute()
        if response.data:
            new_meeting_id = response.data[0]['id']
            return jsonify({'id': new_meeting_id, **meeting_data}), 201
        else:
            return jsonify({'error': 'Failed to create meeting'}), 500
    except postgrest.exceptions.APIError as e:
        return jsonify({'error': str(e)}), 400

    new_meeting = request.json
    
    # Ensure only the columns that exist in your database are included
    meeting_data = {
        'number': new_meeting.get('number'),
        'date': new_meeting.get('date')
    }
    
    # Validate required fields
    if not meeting_data['number'] or not meeting_data['date']:
        return jsonify({'error': 'Meeting number and date are required'}), 400
    
    # Convert date string to timestamp with timezone if it's not already
    if isinstance(meeting_data['date'], str):
        try:
            meeting_data['date'] = datetime.fromisoformat(meeting_data['date']).replace(tzinfo=timezone.utc)
        except ValueError:
            return jsonify({'error': 'Invalid date format. Please use ISO format (YYYY-MM-DDTHH:MM:SS.sssZ)'}), 400
    
    try:
        response = supabase.table('meetings').insert(meeting_data).execute()
        if response.data:
            new_meeting_id = response.data[0]['id']
            return jsonify({'id': new_meeting_id, **meeting_data}), 201
        else:
            return jsonify({'error': 'Failed to create meeting'}), 500
    except postgrest.exceptions.APIError as e:
        return jsonify({'error': str(e)}), 400

@app.route('/meetings/<int:meeting_id>', methods=['PUT'])
def update_meeting(meeting_id):
    updated_data = request.json
    response = supabase.table('meetings').update(updated_data).eq('id', meeting_id).execute()
    return jsonify(response.data[0])

@app.route('/meetings/<int:meeting_id>', methods=['DELETE'])
def delete_meeting(meeting_id):
    supabase.table('meetings').delete().eq('id', meeting_id).execute()
    return '', 204

@app.route('/meetings/<int:meeting_id>/tasks', methods=['POST'])
def add_task_to_meeting(meeting_id):
    task_id = request.json['task_id']
    response = supabase.table('meeting_tasks').insert({
        'meeting_id': meeting_id,
        'task_id': task_id
    }).execute()
    return jsonify(response.data[0])

@app.route('/meetings/<int:meeting_id>/tasks/<int:task_id>', methods=['DELETE'])
def remove_task_from_meeting(meeting_id, task_id):
    supabase.table('meeting_tasks').delete().eq('meeting_id', meeting_id).eq('task_id', task_id).execute()
    return '', 204

@app.route('/meetings/<int:meeting_id>/tasks/<int:task_id>', methods=['PUT'])
def update_meeting_task(meeting_id, task_id):
    data = request.json
    if 'minutes' in data:
        response = supabase.table('meeting_tasks').update({'minutes': data['minutes']}).eq('meeting_id', meeting_id).eq('task_id', task_id).execute()
    return jsonify(response.data[0]), 200

@app.route('/meetings/<int:meeting_id>/generate_report', methods=['GET'])
def generate_report(meeting_id):
    meeting_response = supabase.table('meetings').select('*').eq('id', meeting_id).execute()
    if not meeting_response.data:
        return jsonify({'error': 'Meeting not found'}), 404
    
    meeting = meeting_response.data[0]
    
    tasks_response = supabase.table('meeting_tasks').select('tasks(*)').eq('meeting_id', meeting_id).execute()
    tasks = [item['tasks'] for item in tasks_response.data]

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
    for index, task in enumerate(tasks, start=1):
        document.add_paragraph(f"{task['caseNumber']} {task['title']} page {index}")

    for task in tasks:
        document.add_page_break()

        document.add_paragraph("Proposal – New course idea", style='Heading 2')
        
        headings = [
            ("Item number", task['caseNumber']),
            ("Idea title", task['title']),
            ("Idea owner", task['owner']),
            ("Briefly describe the idea", task['description']),
            ("Why is this relevant for BI?", task.get('relevance_for_bi', '')),
            ("Why does individuals and/or organizations need such a course/idea?", task.get('need_for_course', '')),
            ("What would be the relevant target group?", task.get('target_group', '')),
            ("What are your thoughts on the future growth potential of the market for this course/idea?", task.get('growth_potential', '')),
            ("Faculty resources – which academic departments should be involved?", task.get('faculty_resources', ''))
        ]

        for heading, content in headings:
            document.add_paragraph(heading, style='Heading 3')
            document.add_paragraph(content)

    f = BytesIO()
    document.save(f)
    f.seek(0)

    return send_file(f, as_attachment=True, download_name='innovation_board_sakspapirer.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/meetings/<int:meeting_id>/generate_minutes', methods=['GET'])
def generate_minutes(meeting_id):
    meeting_response = supabase.table('meetings').select('*').eq('id', meeting_id).execute()
    if not meeting_response.data:
        return jsonify({'error': 'Meeting not found'}), 404
    
    meeting = meeting_response.data[0]
    
    tasks_response = supabase.table('meeting_tasks').select('tasks(*), minutes').eq('meeting_id', meeting_id).execute()
    tasks = tasks_response.data

    document = Document()
    
    title = document.add_paragraph(f"Møtereferat for møte {meeting['number']}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True
    title.runs[0].font.size = Pt(14)

    date_paragraph = document.add_paragraph(f"Dato: {meeting['date']}")
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for task in tasks:
        item_title = document.add_paragraph(f"{task['tasks']['caseNumber']} - {task['tasks']['title']}")
        item_title.runs[0].bold = True
        
        if task['minutes']:
            document.add_paragraph(task['minutes'])
        else:
            document.add_paragraph("Ingen referat tilgjengelig for dette punktet.")
        
        document.add_paragraph()

    f = BytesIO()
    document.save(f)
    f.seek(0)

    return send_file(f, as_attachment=True, download_name=f'meeting_{meeting_id}_minutes.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/import-nettskjema', methods=['POST'])
def import_nettskjema():
    try:
        logging.info("Starting Nettskjema import process")
        submissions = get_nettskjema_data(NETTSKJEMA_FORM_ID)
        logging.info(f"Received {len(submissions)} submissions from Nettskjema")
        
        if not submissions:
            logging.warning("No submissions received from Nettskjema")
            return jsonify({"message": "No submissions to import"}), 200

        imported_tasks = []
        existing_submission_ids = set(supabase.table('tasks').select('submission_id').execute().data)

        for submission in submissions:
            logging.debug(f"Processing submission: {submission}")
            if not isinstance(submission, list) or not submission:
                logging.warning(f"Invalid submission format: {submission}")
                continue

            submission_id = submission[0].get('submissionId')
            if submission_id in existing_submission_ids:
                logging.info(f"Skipping already existing submission with ID: {submission_id}")
                continue
            
            try:
                task = transform_submission_to_task(submission)
                logging.debug(f"Transformed task: {task}")
                response = supabase.table('tasks').insert(task).execute()
                if response.data:
                    imported_tasks.append(response.data[0])
                else:
                    logging.warning(f"No data returned when inserting task: {task}")
            except postgrest.exceptions.APIError as e:
                logging.error(f"Error inserting task into Supabase: {str(e)}")
            except Exception as e:
                logging.error(f"Error processing submission: {str(e)}", exc_info=True)

        logging.info(f"Successfully imported {len(imported_tasks)} tasks")
        return jsonify({
            "message": f"Successfully imported {len(imported_tasks)} tasks",
            "imported_tasks": imported_tasks
        }), 200
    except requests.RequestException as e:
        logging.error(f"Error fetching data from Nettskjema: {str(e)}", exc_info=True)
        return jsonify({"error": f"Error fetching data from Nettskjema: {str(e)}"}), 500
    except Exception as e:
        logging.error(f"Error processing import: {str(e)}", exc_info=True)
        return jsonify({"error": f"Error processing import: {str(e)}"}), 500

@app.route('/form/<int:form_id>/settings', methods=['GET'])
def get_form_settings(form_id):
    response = supabase.table('forms').select('*').eq('id', form_id).execute()
    if not response.data:
        return jsonify({'error': 'Form not found'}), 404
    return jsonify(response.data[0])

@app.route('/form/<int:form_id>/settings', methods=['PATCH'])
def patch_form_settings(form_id):
    updated_data = request.json
    force_cleanup = request.args.get('forceCleanup', 'false').lower() == 'true'
    
    response = supabase.table('forms').update(updated_data).eq('id', form_id).execute()
    if not response.data:
        return jsonify({'error': 'Form not found'}), 404
    
    if force_cleanup:
        # Implement cleanup logic here if needed
        pass
    
    return jsonify(response.data[0])

@app.route('/me', methods=['GET'])
def get_environment():
    # This endpoint might need to be adjusted based on how you're handling user authentication with Supabase
    return jsonify({"message": "User environment information"}), 200

@app.route('/form/<int:form_id>/submission-metadata', methods=['GET'])
def get_all_metadata_answers(form_id):
    response = supabase.table('submissions').select('*').eq('form_id', form_id).execute()
    return jsonify(response.data)

@app.route('/form/<int:form_id>/submission-metadata-postponed', methods=['GET'])
def get_all_postponed_metadata_answers(form_id):
    response = supabase.table('submissions').select('*').eq('form_id', form_id).eq('postponed', True).execute()
    return jsonify(response.data)

@app.route('/form/<int:form_id>/spss-syntax', methods=['GET'])
def get_spss_syntax_file(form_id):
    # Implement SPSS syntax file generation logic here
    return jsonify({"message": "SPSS syntax file generation not implemented"}), 501

@app.route('/form/<int:form_id>/excel-report', methods=['GET'])
def get_excel_report(form_id):
    # Implement Excel report generation logic here
    return jsonify({"message": "Excel report generation not implemented"}), 501

@app.route('/form/<int:form_id>/csv-report', methods=['GET'])
def get_csv_report(form_id):
    # Implement CSV report generation logic here
    return jsonify({"message": "CSV report generation not implemented"}), 501

@app.errorhandler(404)
def not_found(error):
    return jsonify({"error": "Not found"}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"error": "Internal server error"}), 500

if __name__ == '__main__':
    app.run(debug=True)
