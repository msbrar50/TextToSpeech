import os
import re
import requests
from docx import Document
from bs4 import BeautifulSoup

# Function to read text from a DOCX file and remove initial numbers
def read_docx(file_path):
    doc = Document(file_path)
    text = []
    for paragraph in doc.paragraphs:
        # Remove leading numbers and any following period, colon, or tab
        cleaned_text = re.sub(r'^\d+\.\s*', '', paragraph.text)
        if cleaned_text.strip():  # Only add non-empty lines
            text.append(cleaned_text)
    return "\n".join(text)

# Function to convert text to speech using TTSMP3 with a specified French voice
def text_to_speech(text, voice="Celine"):
    url = 'https://ttsmp3.com/make.php'
    data = {
        'msg': text,
        'lang': 'fr',       # Language code for French
        'source': 'ttsmp3',
        'name': voice,      # Speaker name, e.g., "Celine"
        'speed': '1.0',
    }
    
    # Send request to TTSMP3 and parse response
    response = requests.post(url, data=data)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Find the audio link from the response (requires inspection on the actual site)
    audio_element = soup.find('audio')
    if not audio_element:
        raise ValueError("Audio URL not found in response")
    
    audio_url = audio_element['src']
    return audio_url

# Function to download audio from a URL and save it locally
def download_audio(audio_url, filename):
    response = requests.get(audio_url)
    with open(filename, 'wb') as audio_file:
        audio_file.write(response.content)

# Main process
def main(docx_file_path, output_folder="downloaded_audio", speaker="Celine"):
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Read French text from DOCX, excluding initial numbers
    french_text = read_docx(docx_file_path)

    # Convert text to speech with specified speaker and get audio URL
    print(f"Converting text to speech with voice '{speaker}'...")
    audio_url = text_to_speech(french_text, voice=speaker)

    # Define output file path
    audio_filename = os.path.join(output_folder, f'{speaker}_output.mp3')

    # Download and save the audio file
    print("Downloading audio...")
    download_audio(audio_url, audio_filename)
    print(f'Audio downloaded successfully with voice "{speaker}": {audio_filename}')

# Specify your DOCX file path and desired speaker
docx_file = 'french_3000.docx'  # Replace with your DOCX file path
main(docx_file, speaker="Celine")  # Replace "Celine" with the desired speaker's name
