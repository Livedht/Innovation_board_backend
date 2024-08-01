import requests

# Replace with your actual API token and form ID
api_token = 'tfagt4tvmauppgjn6ler1b2f3dhm4i5mfktr9tp103gcqdd7r0kgau607o03k1ot5idpf6jarfc79lthloqqqnh1mpi56dopmsjoi0qojmtf58m2m8v5ocur95nmqul1ilgo'
form_id = '397131'

# Nettskjema API endpoint for form submissions
api_url = f'https://nettskjema.no/api/v2/forms/{form_id}/submissions'

# Set the headers for the request
headers = {
    'Authorization': f'Bearer {api_token}',
    'Content-Type': 'application/json'
}

# Make a GET request to retrieve form submissions
response = requests.get(api_url, headers=headers)

if response.status_code == 200:
    data = response.json()
    # Process the data as needed
    for submission in data:
        print(submission)
else:
    print(f"Failed to retrieve data: {response.status_code}")
    print(response.text)
