from flask import Flask, render_template, request, jsonify, session
import speech_recognition as sr
import smtplib
import imaplib
import email
import email.utils as email_utils
import pyttsx3
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
import re
import threading
import time
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
import mimetypes
import wave
import pyaudio
import email.header

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'your-secret-key')  # Add a secret key for session management

# Initialize text-to-speech engine
def create_engine():
    engine = pyttsx3.init()
    engine.setProperty('rate', 150)  # Slower speaking rate
    engine.setProperty('volume', 1.0)  # Maximum volume
    return engine

# Thread-local storage for TTS engine
thread_local = threading.local()

def get_engine():
    if not hasattr(thread_local, 'engine'):
        thread_local.engine = create_engine()
    return thread_local.engine

# Email configuration
EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

# Email mapping dictionary
EMAIL_MAPPINGS = {
    'alpesh': 'alpeshchandankhede@gmail.com',  # Using the email with '250'
    'anish': 'anishpahune79@gmail.com',
    'Utsav': 'utsavgavai22@gmail.com',
    'rishi': 'rishinaik616@gmail.com',
    'Prajval': 'prajvalkawale@gmail.com',  
    'test': '',
    'support': '',
    'info': '',
    'admin': '',
    'contact': '',
    'help': '',
    # Add more mappings as needed
}

# Password mapping dictionary
PASSWORD_MAPPINGS = {
    'abc': 'zladryfkszykmgyh',
    # Add more mappings as needed
}

def map_email(spoken_text):
    """Map spoken email keywords to actual email addresses"""
    spoken_text = spoken_text.lower().strip()
    return EMAIL_MAPPINGS.get(spoken_text, spoken_text)

def map_password(spoken_text):
    """Map spoken password keywords to actual passwords"""
    spoken_text = spoken_text.lower().strip()
    return PASSWORD_MAPPINGS.get(spoken_text, spoken_text)

def speech_to_text():
    try:
        # Initialize recognizer
        r = sr.Recognizer()
        
        # Get list of microphones before initializing
        mics = sr.Microphone.list_microphone_names()
        print("\n=== Speech Recognition Started ===")
        print("Available Microphones:", mics)
        
        # Try to find the default microphone
        try:
            default_mic_index = mics.index("Microphone (Realtek(R) Audio)")
        except ValueError:
            default_mic_index = 0  # Use first available microphone if default not found
        
        # Initialize microphone with explicit device index
        with sr.Microphone(device_index=default_mic_index) as source:
            print(f"Using microphone: {mics[default_mic_index]}")
            
            # More aggressive noise adjustment
            print("\nAdjusting for ambient noise... (please be quiet)")
            r.adjust_for_ambient_noise(source, duration=3)  # Increased to 3 seconds
            
            # More lenient recognition settings
            r.energy_threshold = 200  # Even lower threshold for easier detection
            r.dynamic_energy_threshold = True
            r.pause_threshold = 1.2  # Longer pause to ensure full phrase capture
            r.phrase_threshold = 0.5  # More lenient phrase detection
            r.non_speaking_duration = 0.8  # Longer duration to avoid cutting off
            
            print(f"\nRecognition settings configured:")
            print(f"- Energy threshold: {r.energy_threshold}")
            print(f"- Dynamic threshold: {r.dynamic_energy_threshold}")
            print(f"- Pause threshold: {r.pause_threshold}s")
            print(f"- Phrase threshold: {r.phrase_threshold}")
            print(f"- Non-speaking duration: {r.non_speaking_duration}s")
            
            print("\n🎤 Listening... (Speak now)")
            try:
                # More generous timeouts
                audio = r.listen(source, timeout=15, phrase_time_limit=10)
                print("✅ Audio captured! Converting to text...")
                
                try:
                    # First attempt - US English with more variations
                    text = r.recognize_google(audio, language='en-US')
                    print(f"🗣️ You said: '{text}'")
                    print("✅ Recognition successful!")
                    return {"success": True, "text": text}
                    
                except sr.UnknownValueError:
                    # Second attempt - Indian English
                    try:
                        print("Trying Indian English model...")
                        text = r.recognize_google(audio, language='en-IN')
                        print(f"🗣️ You said: '{text}'")
                        print("✅ Recognition successful with Indian English!")
                        return {"success": True, "text": text}
                    except:
                        # Third attempt - General recognition
                        try:
                            print("Trying general recognition model...")
                            text = r.recognize_google(audio)
                            print(f"🗣️ You said: '{text}'")
                            print("✅ Recognition successful with fallback!")
                            return {"success": True, "text": text}
                        except:
                            raise sr.UnknownValueError
                    
            except sr.WaitTimeoutError:
                print("❌ Error: Listening timeout - No speech detected")
                return {"success": False, "error": "No speech detected. Please speak when you see 'Listening...' and ensure your microphone is working."}
            except sr.UnknownValueError:
                print("❌ Error: Could not understand audio - Speech unclear")
                return {"success": False, "error": "Could not understand audio. Please try:\n1. Speaking more slowly and clearly\n2. Moving closer to the microphone\n3. Reducing background noise"}
            except sr.RequestError as e:
                print(f"❌ Error: Recognition service error - {str(e)}")
                return {"success": False, "error": "Speech recognition service error. Please check your internet connection."}
            except Exception as e:
                print(f"❌ Error: Unexpected error - {str(e)}")
                return {"success": False, "error": "An unexpected error occurred. Please try again."}
    except Exception as e:
        print(f"❌ Error: Microphone initialization failed - {str(e)}")
        return {"success": False, "error": "Could not initialize microphone. Please check your microphone settings and try again."}
    finally:
        print("=== Speech Recognition Ended ===\n")

def text_to_speech(text):
    try:
        engine = get_engine()
        # Initialize new speech
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Error in text to speech: {str(e)}")
        try:
            # Try to reinitialize the engine
            thread_local.engine = create_engine()
            engine = thread_local.engine
            engine.say(text)
            engine.runAndWait()
        except Exception as e2:
            print(f"Failed to reinitialize TTS engine: {str(e2)}")

def validate_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def record_audio(duration=5):
    """Record audio for the specified duration"""
    try:
        CHUNK = 1024
        FORMAT = pyaudio.paInt16
        CHANNELS = 1
        RATE = 44100
        
        p = pyaudio.PyAudio()
        
        stream = p.open(format=FORMAT,
                       channels=CHANNELS,
                       rate=RATE,
                       input=True,
                       frames_per_buffer=CHUNK)
        
        print("🎤 Recording audio...")
        frames = []
        
        for i in range(0, int(RATE / CHUNK * duration)):
            data = stream.read(CHUNK)
            frames.append(data)
            
        print("✅ Finished recording")
        
        stream.stop_stream()
        stream.close()
        p.terminate()
        
        # Save the recorded audio
        filename = "attachment.wav"
        wf = wave.open(filename, 'wb')
        wf.setnchannels(CHANNELS)
        wf.setsampwidth(p.get_sample_size(FORMAT))
        wf.setframerate(RATE)
        wf.writeframes(b''.join(frames))
        wf.close()
        
        return {"success": True, "filename": filename}
    except Exception as e:
        print(f"Error recording audio: {e}")
        return {"success": False, "error": str(e)}

def send_email(recipients, subject, body, attachment_path=None):
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = ", ".join(recipients)  # Join multiple recipients with commas
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        # Handle attachment if provided
        if attachment_path:
            # Guess the content type of the attachment
            content_type, encoding = mimetypes.guess_type(attachment_path)
            
            if content_type is None:
                content_type = 'application/octet-stream'
                
            main_type, sub_type = content_type.split('/', 1)
            
            with open(attachment_path, 'rb') as f:
                if main_type == 'audio':
                    # Handle audio files
                    attachment = MIMEAudio(f.read(), _subtype=sub_type)
                else:
                    # Handle other file types
                    attachment = MIMEBase(main_type, sub_type)
                    attachment.set_payload(f.read())
                    email.encoders.encode_base64(attachment)
                
                # Add header
                attachment.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=os.path.basename(attachment_path)
                )
                msg.attach(attachment)

        # Create SMTP session
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL, PASSWORD)
        
        # Send email
        server.send_message(msg)
        server.quit()
        
        # Clean up attachment file if it was a recorded audio
        if attachment_path and attachment_path == "attachment.wav":
            try:
                os.remove(attachment_path)
            except:
                pass
                
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def decode_header_str(header_str):
    """Helper function to decode email headers properly"""
    if not header_str:
        return ''
    try:
        # Decode the header parts
        parts = email.header.decode_header(header_str)
        # Combine all parts into a single string
        decoded_parts = []
        for part, charset in parts:
            if isinstance(part, bytes):
                try:
                    # Try with the specified charset first
                    if charset:
                        decoded_parts.append(part.decode(charset))
                    else:
                        # If no charset specified, try utf-8 first, then fallback to others
                        try:
                            decoded_parts.append(part.decode('utf-8'))
                        except UnicodeDecodeError:
                            try:
                                decoded_parts.append(part.decode('iso-8859-1'))
                            except UnicodeDecodeError:
                                decoded_parts.append(part.decode('ascii', 'ignore'))
                except Exception:
                    # If all decoding fails, use ascii with ignore
                    decoded_parts.append(part.decode('ascii', 'ignore'))
            else:
                decoded_parts.append(str(part))
        return ' '.join(decoded_parts)
    except Exception as e:
        print(f"Error decoding header: {e}")
        return header_str

def read_emails(num_emails=5):
    try:
        # Connect to IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        
        all_emails = []
        
        # Read Inbox emails
        mail.select('inbox')
        _, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        
        # Get the last n inbox emails
        for i in range(min(num_emails, len(email_ids))):
            email_id = email_ids[-(i+1)]
            _, msg = mail.fetch(email_id, '(RFC822)')
            email_body = msg[0][1]
            msg_obj = email.message_from_bytes(email_body)
            
            # Decode headers properly
            subject = decode_header_str(msg_obj['subject']) or '(No subject)'
            sender = decode_header_str(msg_obj['from']) or '(No sender)'
            date = msg_obj['date'] or '(No date)'
            
            all_emails.append({
                'folder': 'Inbox',
                'subject': subject,
                'sender': sender,
                'date': date
            })
        
        # Read Sent emails
        mail.select('"[Gmail]/Sent Mail"')
        _, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        
        # Get the last n sent emails
        for i in range(min(num_emails, len(email_ids))):
            email_id = email_ids[-(i+1)]
            _, msg = mail.fetch(email_id, '(RFC822)')
            email_body = msg[0][1]
            msg_obj = email.message_from_bytes(email_body)
            
            # Decode headers properly
            subject = decode_header_str(msg_obj['subject']) or '(No subject)'
            recipient = decode_header_str(msg_obj['to']) or '(No recipient)'
            date = msg_obj['date'] or '(No date)'
            
            all_emails.append({
                'folder': 'Sent',
                'subject': subject,
                'recipient': recipient,
                'date': date
            })
        
        # Sort all emails by date
        all_emails.sort(key=lambda x: email_utils.parsedate_to_datetime(x['date']), reverse=True)
        
        mail.close()
        mail.logout()
        return all_emails
    except Exception as e:
        print(f"Error reading emails: {e}")
        return []

def read_unread_emails(num_emails=5):
    try:
        # Connect to IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        
        unread_emails = []
        
        # List all available folders to find Primary
        print("Listing available folders...")
        folders = mail.list()[1]
        primary_folder = None
        
        # Try to find Primary category folder
        for folder in folders:
            folder_name = folder.decode().split('"/')[-1].replace('"', '')
            if 'Primary' in folder_name:
                primary_folder = folder_name
                break
        
        if primary_folder:
            print(f"Found Primary folder: {primary_folder}")
            status = mail.select(f'"{primary_folder}"')
        else:
            print("Primary folder not found, using INBOX")
            status = mail.select('INBOX')
            
        if status[0] != 'OK':
            print(f"Error selecting folder: {status}")
            raise Exception("Could not select mail folder")
            
        print("Searching for unread messages...")
        # Search for unread messages
        _, messages = mail.search(None, 'UNSEEN')
        email_ids = messages[0].split()
        
        if not email_ids:
            print("No unread messages found")
            return []
            
        # Reverse the list to get newest first
        email_ids = email_ids[::-1]
        print(f"Found {len(email_ids)} unread emails")
        
        # Get the first n unread emails (newest first)
        for i in range(min(num_emails, len(email_ids))):
            email_id = email_ids[i]
            print(f"Fetching unread email {i+1}/{min(num_emails, len(email_ids))}...")
            
            try:
                # Fetch the email message
                _, msg_data = mail.fetch(email_id, '(RFC822)')
                email_body = msg_data[0][1]
                msg_obj = email.message_from_bytes(email_body)
                
                # Extract email details with better error handling and proper decoding
                subject = decode_header_str(msg_obj['subject'])
                if not subject:
                    subject = '(No subject)'
                
                sender = decode_header_str(msg_obj['from'])
                if not sender:
                    sender = '(No sender)'
                    
                date = msg_obj['date']
                if not date:
                    date = '(No date)'
                
                # Get the message body
                body = ""
                if msg_obj.is_multipart():
                    for part in msg_obj.walk():
                        if part.get_content_type() == "text/plain":
                            try:
                                body = part.get_payload(decode=True).decode()
                                break
                            except:
                                continue
                else:
                    try:
                        body = msg_obj.get_payload(decode=True).decode()
                    except:
                        body = "(Could not decode message body)"
                
                # Clean up the body text
                if body:
                    # Remove extra whitespace and newlines
                    body = ' '.join(body.split())
                    # Truncate if too long
                    if len(body) > 200:
                        body = body[:200] + "..."
                
                # Parse the date for accurate sorting
                try:
                    parsed_date = email_utils.parsedate_to_datetime(date)
                except:
                    parsed_date = None
                
                folder_name = 'Primary' if primary_folder else 'Inbox'
                unread_emails.append({
                    'folder': folder_name,
                    'subject': subject,
                    'sender': sender,
                    'date': date,
                    'parsed_date': parsed_date.isoformat() if parsed_date else date,
                    'body': body,
                    'unread': True,
                    'id': email_id.decode()
                })
                
                print(f"Added unread email: From={sender}, Subject={subject}, Date={date}")
                
                # Mark as read after successfully processing
                mail.store(email_id, '+FLAGS', '\\Seen')
                
            except Exception as e:
                print(f"Error processing email {email_id}: {str(e)}")
                continue
        
        # Sort by date to ensure newest first
        unread_emails.sort(key=lambda x: x.get('parsed_date', ''), reverse=True)
        
        mail.close()
        mail.logout()
        return unread_emails
        
    except Exception as e:
        print(f"Error reading unread emails: {e}")
        return []

def read_trash_emails(num_emails=5):
    try:
        # Connect to IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        
        trash_emails = []
        
        # Try different possible trash folder names
        trash_folder_names = ['[Gmail]/Trash', 'Trash', '"[Gmail]/Trash"', 'Bin', 'Deleted Items', 'Deleted']
        selected_folder = None
        
        print("Searching for trash folder...")
        for folder in trash_folder_names:
            try:
                print(f"Trying to select folder: {folder}")
                status = mail.select(folder)
                if status[0] == 'OK':
                    print(f"Successfully selected trash folder: {folder}")
                    selected_folder = folder
                    break
            except Exception as e:
                print(f"Failed to select folder {folder}: {str(e)}")
                continue
        
        if not selected_folder:
            print("Could not find trash folder, listing available folders...")
            # List all folders and try to find one that looks like trash
            folders = mail.list()[1]
            for folder_data in folders:
                folder_name = folder_data.decode()
                if 'trash' in folder_name.lower() or 'bin' in folder_name.lower() or 'deleted' in folder_name.lower():
                    try:
                        trash_name = folder_name.split('"/')[-1].replace('"', '')
                        print(f"Found possible trash folder: {trash_name}")
                        status = mail.select(f'"{trash_name}"')
                        if status[0] == 'OK':
                            print(f"Successfully selected found trash folder: {trash_name}")
                            selected_folder = trash_name
                            break
                    except Exception as e:
                        print(f"Error selecting found trash folder: {str(e)}")
        
        if not selected_folder:
            raise Exception("Could not find trash folder")
            
        # Read trash emails
        print("Searching for emails in trash...")
        _, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        print(f"Found {len(email_ids)} emails in trash")
        
        if not email_ids:
            print("No emails in trash")
            return []
            
        # Reverse the list to get newest first
        email_ids = email_ids[::-1]
        
        # Get the first n trash emails (newest first)
        for i in range(min(num_emails, len(email_ids))):
            email_id = email_ids[i]
            print(f"Fetching trash email {i+1}/{min(num_emails, len(email_ids))}...")
            
            try:
                # Fetch the email message
                _, msg_data = mail.fetch(email_id, '(RFC822)')
                email_body = msg_data[0][1]
                msg_obj = email.message_from_bytes(email_body)
                
                # Extract email details with error handling and proper decoding
                subject = decode_header_str(msg_obj['subject'])
                if not subject:
                    subject = '(No subject)'
                
                # Determine if it's a sent email (look for recipient) or received (look for sender)
                sender = decode_header_str(msg_obj['from'])
                recipient = decode_header_str(msg_obj['to'])
                
                if not sender:
                    sender = '(No sender)'
                    
                if not recipient:
                    recipient = '(No recipient)'
                    
                date = msg_obj['date']
                if not date:
                    date = '(No date)'
                
                # Get the message body
                body = ""
                if msg_obj.is_multipart():
                    for part in msg_obj.walk():
                        if part.get_content_type() == "text/plain":
                            try:
                                body = part.get_payload(decode=True).decode()
                                break
                            except:
                                continue
                else:
                    try:
                        body = msg_obj.get_payload(decode=True).decode()
                    except:
                        body = "(Could not decode message body)"
                
                # Clean up the body text
                if body:
                    # Remove extra whitespace and newlines
                    body = ' '.join(body.split())
                    # Truncate if too long
                    if len(body) > 200:
                        body = body[:200] + "..."
                
                # Parse the date for accurate sorting
                try:
                    parsed_date = email_utils.parsedate_to_datetime(date)
                except:
                    parsed_date = None
                
                # Determine if this was a sent email or received email
                is_sent = (EMAIL.lower() in sender.lower() if sender else False)
                
                trash_emails.append({
                    'folder': 'Trash',
                    'subject': subject,
                    'sender': sender,
                    'recipient': recipient,
                    'date': date,
                    'parsed_date': parsed_date.isoformat() if parsed_date else date,
                    'body': body,
                    'is_sent': is_sent,
                    'id': email_id.decode() if isinstance(email_id, bytes) else str(email_id)
                })
                
                print(f"Added trash email: Subject={subject}, Date={date}")
                
            except Exception as e:
                print(f"Error processing email {email_id}: {str(e)}")
                continue
        
        # Sort by date to ensure newest first
        trash_emails.sort(key=lambda x: x.get('parsed_date', ''), reverse=True)
        
        mail.close()
        mail.logout()
        return trash_emails
        
    except Exception as e:
        print(f"Error reading trash emails: {e}")
        return []

@app.route('/')
def home():
    if 'logged_in' not in session:
        return render_template('index.html', logged_in=False)
    return render_template('index.html', logged_in=True)

def retry_speech_recognition(max_retries=3):
    """Helper function to retry speech recognition with a maximum number of attempts"""
    for attempt in range(max_retries):
        if attempt > 0:
            text_to_speech(f"No response detected. Attempt {attempt + 1} of {max_retries}. Please speak again.")
            time.sleep(1)  # Brief pause before retry
            
        result = speech_to_text()
        if result["success"]:
            return result
            
    return {"success": False, "error": "Maximum retry attempts reached. Please try again later."}

@app.route('/login', methods=['POST'])
def login():
    # Get email with retries
    text_to_speech("Please speak your email keyword")
    email_result = retry_speech_recognition()
    
    if not email_result["success"]:
        return jsonify({"status": "error", "message": email_result["error"]})
    
    # Map the spoken email to full email address
    spoken_email = map_email(email_result["text"].lower().replace(" ", ""))
    print(f"Mapped email: {spoken_email}")  # Debug print
    
    # Get password with retries
    text_to_speech("Please speak the password keyword")
    password_result = retry_speech_recognition()
    
    if not password_result["success"]:
        return jsonify({"status": "error", "message": password_result["error"]})
    
    # Map the spoken password to actual password
    spoken_password = map_password(password_result["text"].lower().replace(" ", ""))
    print(f"Comparing credentials:")  # Debug print
    print(f"Spoken email ({spoken_email}) == EMAIL ({EMAIL})")  # Debug print
    print(f"Spoken password matches PASSWORD: {spoken_password == PASSWORD}")  # Debug print
    
    # Compare exact strings without case conversion
    if spoken_email == EMAIL and spoken_password == PASSWORD:
        session['logged_in'] = True
        return jsonify({"status": "success", "message": "Login successful"})
    else:
        return jsonify({"status": "error", "message": "Invalid credentials"})

@app.route('/logout', methods=['POST'])
def logout():
    session.pop('logged_in', None)
    return jsonify({"status": "success", "message": "Logged out successfully"})

@app.route('/compose', methods=['POST'])
def compose_email():
    if 'logged_in' not in session:
        return jsonify({"status": "error", "message": "Please login first"})
    
    try:
        # Initialize recipients list and attachment path
        recipients = []
        attachment_path = None
        
        while True:
            # Get recipient with retries
            text_to_speech("Please speak the recipient's email keyword or full address")
            recipient_result = retry_speech_recognition()
            if not recipient_result["success"]:
                raise Exception(recipient_result["error"])
            
            # Map the spoken recipient to full email address
            recipient = map_email(recipient_result["text"].lower().replace(" ", ""))
            
            if not validate_email(recipient):
                text_to_speech("Invalid email address format. Please try again.")
                continue
            
            # Add valid recipient to the list
            recipients.append(recipient)
            
            # Ask if user wants to add more recipients
            text_to_speech("Do you want to add another recipient? Say yes or no")
            more_result = retry_speech_recognition()
            
            if not more_result["success"]:
                raise Exception(more_result["error"])
            
            if "no" in more_result["text"].lower():
                break
            elif "yes" in more_result["text"].lower():
                text_to_speech("Adding another recipient.")
                continue
            else:
                text_to_speech("I didn't understand. Proceeding with current recipients.")
                break
        
        # Confirm recipients
        recipient_list = ", ".join(recipients)
        text_to_speech(f"Sending email to: {recipient_list}")
        
        # Get subject with retries
        text_to_speech("Please speak the subject of the email")
        subject_result = retry_speech_recognition()
        if not subject_result["success"]:
            raise Exception(subject_result["error"])
        subject = subject_result["text"]
        
        # Get content with retries
        text_to_speech("Please speak the content of the email")
        content_result = retry_speech_recognition()
        if not content_result["success"]:
            raise Exception(content_result["error"])
        content = content_result["text"]
        
        # Ask about attachment
        text_to_speech("Would you like to add an attachment? Say 'yes' for audio attachment, 'no' for no attachment")
        attachment_result = retry_speech_recognition()
        
        if attachment_result["success"] and "yes" in attachment_result["text"].lower():
            text_to_speech("I will record a 10-second audio attachment. Please speak your message after the beep")
            time.sleep(1)
            # Play a beep sound
            text_to_speech("beep")
            
            # Record audio
            record_result = record_audio(10)  # 10 seconds recording
            if record_result["success"]:
                attachment_path = record_result["filename"]
                text_to_speech("Audio recorded successfully")
            else:
                text_to_speech("Failed to record audio. Proceeding without attachment")
        
        # Confirm before sending
        confirmation_message = f"You are about to send an email to {recipient_list} with subject {subject}"
        if attachment_path:
            confirmation_message += " and an audio attachment"
        confirmation_message += ". Say 'yes' to confirm or 'no' to cancel."
        
        text_to_speech(confirmation_message)
        confirm_result = retry_speech_recognition()
        
        if not confirm_result["success"]:
            raise Exception(confirm_result["error"])
        
        if "yes" in confirm_result["text"].lower():
            # Try to send email
            try:
                send_result = send_email(recipients, subject, content, attachment_path)
                if not send_result:
                    raise Exception("Failed to send email. Please check your internet connection and try again.")
            except Exception as send_error:
                raise Exception(f"Error sending email: {str(send_error)}")
            finally:
                # Clean up attachment regardless of send result
                if attachment_path and os.path.exists(attachment_path):
                    try:
                        os.remove(attachment_path)
                    except Exception as e:
                        print(f"Error cleaning up attachment: {e}")
            
            # Email sent successfully, provide feedback and start listening for commands
            text_to_speech("Email sent successfully. Returning to command mode.")
            return jsonify({
                "status": "success",
                "message": "Email sent successfully",
                "action": "listen_for_commands",
                "reload_page": True,
                "force_redirect": True,
                "available_commands": [
                    "Compose email",
                    "Read inbox",
                    "Read unread",
                    "Read sent",
                    "Read trash",
                    "Logout",
                    "Help"
                ]
            })
        else:
            # Email cancelled, return to command mode
            text_to_speech("Email cancelled. Returning to command mode.")
            return jsonify({
                "status": "cancelled",
                "message": "Email sending cancelled",
                "action": "listen_for_commands",
                "reload_page": True,
                "force_redirect": True,
                "available_commands": [
                    "Compose email",
                    "Read inbox",
                    "Read unread",
                    "Read sent",
                    "Read trash",
                    "Logout",
                    "Help"
                ]
            })
            
    except Exception as e:
        # Clean up attachment if exists
        if attachment_path and os.path.exists(attachment_path):
            try:
                os.remove(attachment_path)
            except:
                pass
        
        error_message = str(e)
        print(f"Error in compose_email: {error_message}")
        
        # Return to command mode with error message
        text_to_speech(f"An error occurred while sending the email. {error_message}. Returning to command mode.")
        return jsonify({
            "status": "error",
            "message": f"An error occurred while composing the email: {error_message}",
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True,
            "available_commands": [
                "Compose email",
                "Read inbox",
                "Read unread",
                "Read sent",
                "Read trash",
                "Logout",
                "Help"
            ]
        })

@app.route('/read_inbox', methods=['GET'])
def read_inbox():
    try:
        emails = read_emails()
        # Filter only inbox emails
        inbox_emails = [email for email in emails if email['folder'] == 'Inbox']
        
        # Don't read emails automatically, just return them
        return jsonify({
            "status": "success", 
            "emails": emails,
            "total": len(inbox_emails),
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True,
            "available_commands": [
                "Compose email",
                "Read inbox",
                "Read unread",
                "Read sent",
                "Read trash",
                "Logout",
                "Help"
            ]
        })
    except Exception as e:
        error_msg = f"Error reading inbox: {str(e)}"
        print(error_msg)
        return jsonify({"status": "error", "message": error_msg})

@app.route('/read_sent', methods=['GET'])
def read_sent():
    try:
        # Connect to IMAP server
        print("Connecting to IMAP server...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        print("Successfully logged in to IMAP server")
        
        sent_emails = []
        
        # List all available folders
        print("Available folders:", mail.list())
        
        # Try different possible sent folder names
        sent_folder_names = ['[Gmail]/Sent Mail', '"[Gmail]/Sent Mail"', 'Sent', '[Gmail]/Sent']
        selected_folder = None
        
        for folder in sent_folder_names:
            try:
                print(f"Trying to select folder: {folder}")
                status = mail.select(folder)
                if status[0] == 'OK':
                    print(f"Successfully selected folder: {folder}")
                    selected_folder = folder
                    break
            except Exception as e:
                print(f"Failed to select folder {folder}: {str(e)}")
                continue
        
        if not selected_folder:
            raise Exception("Could not find sent mail folder")
        
        # Read Sent emails
        print("Searching for emails...")
        _, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        print(f"Found {len(email_ids)} emails")
        
        # Get the last 5 sent emails
        for i in range(min(5, len(email_ids))):
            email_id = email_ids[-(i+1)]
            print(f"Fetching email {i+1}/5...")
            _, msg = mail.fetch(email_id, '(RFC822)')
            email_body = msg[0][1]
            msg_obj = email.message_from_bytes(email_body)
            
            # Use decode_header_str for proper header decoding
            subject = decode_header_str(msg_obj['subject']) or '(No subject)'
            recipient = decode_header_str(msg_obj['to']) or '(No recipient)'
            date = msg_obj['date'] or '(No date)'
            
            # Get the message body
            body = ""
            if msg_obj.is_multipart():
                for part in msg_obj.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            body = part.get_payload(decode=True).decode()
                            break
                        except:
                            continue
            else:
                try:
                    body = msg_obj.get_payload(decode=True).decode()
                except:
                    body = "(Could not decode message body)"
            
            # Clean up the body text
            if body:
                # Remove extra whitespace and newlines
                body = ' '.join(body.split())
                # Truncate if too long
                if len(body) > 200:
                    body = body[:200] + "..."
            
            # Try to parse the date
            try:
                parsed_date = email_utils.parsedate_to_datetime(date)
            except:
                parsed_date = None
            
            sent_emails.append({
                'folder': 'Sent',
                'subject': subject,
                'recipient': recipient,
                'date': date,
                'parsed_date': parsed_date.isoformat() if parsed_date else date,
                'body': body,
                'id': str(i)  # Using index as ID since we don't have actual email IDs for sent mails
            })
            print(f"Added email: To={recipient}, Subject={subject}")
        
        # Sort sent emails by date
        sent_emails.sort(key=lambda x: x.get('parsed_date', x['date']), reverse=True)
        
        mail.close()
        mail.logout()
        print("Successfully closed IMAP connection")
        
        # Don't read out the sent emails, just return them
        return jsonify({
            "status": "success", 
            "emails": sent_emails,
            "total": len(sent_emails),
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True,
            "available_commands": [
                "Compose email",
                "Read inbox",
                "Read unread",
                "Read sent",
                "Read trash",
                "Logout",
                "Help"
            ]
        })
    except Exception as e:
        error_msg = f"Error reading sent emails: {str(e)}"
        print(error_msg)
        return jsonify({"status": "error", "message": error_msg})

@app.route('/test_microphone', methods=['POST'])
def test_microphone():
    text_to_speech("Please speak a test phrase")
    result = speech_to_text()
    
    if result["success"]:
        return jsonify({
            "status": "success",
            "text": result["text"],
            "message": "Microphone test successful!"
        })
    else:
        return jsonify({
            "status": "error",
            "message": result["error"]
        })

@app.route('/read_unread', methods=['GET'])
def read_unread():
    try:
        unread_emails = read_unread_emails()
        if not unread_emails:
            text_to_speech("You have no unread messages")
            return jsonify({
                "status": "success", 
                "message": "No unread messages", 
                "emails": [],
                "action": "listen_for_commands",
                "reload_page": True,
                "force_redirect": True,
                "available_commands": [
                    "Compose email",
                    "Read inbox",
                    "Read unread",
                    "Read sent",
                    "Read trash",
                    "Logout",
                    "Help"
                ]
            })
            
        # Return emails without reading them out - let frontend handle reading sequence
        return jsonify({
            "status": "success", 
            "emails": unread_emails,
            "total": len(unread_emails),
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True,
            "available_commands": [
                "Compose email",
                "Read inbox",
                "Read unread",
                "Read sent",
                "Read trash",
                "Logout",
                "Help"
            ]
        })
    except Exception as e:
        error_msg = f"Error reading unread emails: {str(e)}"
        print(error_msg)
        return jsonify({"status": "error", "message": error_msg})

@app.route('/read_email_aloud', methods=['POST'])
def read_email_aloud():
    try:
        data = request.get_json()
        email_data = data.get('email')
        index = data.get('index', 0)  # Get the email index
        
        if not email_data:
            return jsonify({"status": "error", "message": "No email data provided"})
        
        try:
            # Wait for voice command
            text_to_speech(f"Email {index + 1}. Say 'read' to hear this email or give a command like 'read first', 'read second', etc. You can also say 'return to main menu' to go back.")
            result = speech_to_text()
            
            if not result["success"]:
                return jsonify({"status": "error", "message": result["error"]})
                
            spoken_command = result["text"].lower()
            
            # Check for "return to main menu" command
            if ('return' in spoken_command and ('menu' in spoken_command or 'main' in spoken_command)) or 'main menu' in spoken_command:
                return jsonify({
                    "status": "success", 
                    "action": "return_to_menu"
                })
            
            # Map number words to digits
            number_map = {
                'first': 1, 'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5,
                'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
                '1st': 1, '2nd': 2, '3rd': 3, '4th': 4, '5th': 5,
                '1': 1, '2': 2, '3': 3, '4': 4, '5': 5
            }
            
            # Check if command contains a number
            requested_number = None
            for word, num in number_map.items():
                if word in spoken_command:
                    requested_number = num
                    break
            
            # If a specific number was requested
            if requested_number is not None:
                # Convert to 0-based index
                requested_index = requested_number - 1
                if requested_index == index:
                    # Read the current email
                    read_email_content(email_data, index)
                    return jsonify({"status": "success", "action": "read"})
                else:
                    # User wants to read a different email
                    return jsonify({
                        "status": "success", 
                        "action": "jump",
                        "target_index": requested_index
                    })
            
            # Handle simple read/skip commands
            elif 'read' in spoken_command:
                read_email_content(email_data, index)
                return jsonify({"status": "success", "action": "read"})
            elif 'skip' in spoken_command:
                return jsonify({"status": "success", "action": "skip"})
            else:
                return jsonify({"status": "error", "message": "Unknown command. Please say 'read', 'skip', 'return to main menu', or specify which email to read (e.g., 'read first')"})
                
        except Exception as e:
            print(f"Error during voice interaction: {str(e)}")
            return jsonify({"status": "error", "message": "Error during voice interaction"})
            
    except Exception as e:
        error_msg = f"Error reading email aloud: {str(e)}"
        print(error_msg)
        return jsonify({"status": "error", "message": error_msg})

def read_email_content(email_data, index):
    """Helper function to read email content with separate TTS calls"""
    try:
        text_to_speech(f"Reading email {index + 1}")
        time.sleep(0.2)  # Small pauses between speech segments
        
        # Check what type of email it is
        folder = email_data.get('folder', '').lower()
        
        if folder == 'sent':
            # For sent emails
            text_to_speech(f"Sent to {email_data['recipient']}")
        elif folder == 'trash':
            # For trash emails, check if it was a sent or received email
            if email_data.get('is_sent', False):
                text_to_speech(f"Deleted email sent to {email_data['recipient']}")
            else:
                text_to_speech(f"Deleted email from {email_data['sender']}")
        else:
            # For inbox or other emails
            text_to_speech(f"From {email_data['sender']}")
        
        time.sleep(0.2)
        
        text_to_speech(f"Subject: {email_data['subject']}")
        time.sleep(0.2)
        
        if 'body' in email_data and email_data['body']:
            body = email_data['body']
            if len(body) > 500:
                body = body[:500] + "... Message truncated."
            text_to_speech(f"Message: {body}")
    except Exception as e:
        print(f"Error reading email content: {str(e)}")
        # Continue even if there's an error - don't break the flow

@app.route('/read_trash', methods=['GET'])
def read_trash():
    try:
        trash_emails = read_trash_emails()
        
        if not trash_emails:
            text_to_speech("You have no emails in trash")
            return jsonify({
                "status": "success", 
                "message": "No emails in trash", 
                "emails": [],
                "action": "listen_for_commands",
                "reload_page": True,
                "force_redirect": True,
                "available_commands": [
                    "Compose email",
                    "Read inbox",
                    "Read unread",
                    "Read sent",
                    "Read trash",
                    "Logout",
                    "Help"
                ]
            })
            
        # Return trash emails without reading them aloud
        return jsonify({
            "status": "success", 
            "emails": trash_emails,
            "total": len(trash_emails),
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True,
            "available_commands": [
                "Compose email",
                "Read inbox",
                "Read unread",
                "Read sent",
                "Read trash",
                "Logout",
                "Help"
            ]
        })
    except Exception as e:
        error_msg = f"Error reading trash: {str(e)}"
        print(error_msg)
        return jsonify({"status": "error", "message": error_msg})

@app.route('/voice_command', methods=['POST'])
def voice_command():
    """Handle voice commands for app navigation"""
    # First check if the user is logged in
    if 'logged_in' not in session:
        text_to_speech("Please log in first. Say login to proceed.")
        result = speech_to_text()
        if result["success"] and "login" in result["text"].lower():
            return jsonify({"status": "success", "action": "login"})
        else:
            return jsonify({"status": "error", "message": "You need to log in first."})
    
    # If logged in, listen for navigation commands
    text_to_speech("What would you like to do? Say compose email, read inbox, read unread, read sent, or read trash. For sent mail, you can also say read send, read saint, or read scent. Say 'return to main menu' anytime to go back.")
    result = speech_to_text()
    
    if not result["success"]:
        return jsonify({"status": "error", "message": result["error"]})
    
    command = result["text"].lower()
    print(f"Voice command received: {command}")
    
    # Special handling for "return to main menu" command
    if 'return' in command and ('menu' in command or 'main' in command):
        text_to_speech("Returning to main menu.")
        return jsonify({
            "status": "success", 
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True
        })
    
    # Map of command keywords to actions
    command_map = {
        'compose': ['compose', 'write', 'new', 'create', 'send'],
        'read_inbox': ['inbox', 'check mail', 'check email', 'check inbox'],
        'read_unread': ['unread', 'new email', 'new mail', 'new message'],
        'read_sent': ['sent', 'send', 'saint', 'scent', 'cent', 'sand', 'spend', 'outbox', 'sent mail', 'sent email', 'outgoing', 'my emails'],
        'read_trash': ['trash', 'bin', 'deleted', 'deleted mail', 'deleted email'],
        'logout': ['logout', 'sign out', 'exit', 'bye', 'goodbye'],
        'stop': ['stop', 'cancel', 'end'],
        'main_menu': ['main menu', 'return', 'go back', 'home']
    }
    
    # Look for command keywords
    detected_action = None
    for action, keywords in command_map.items():
        for keyword in keywords:
            if keyword in command:
                detected_action = action
                break
        if detected_action:
            break
    
    # Special handling for "stop listening" command
    if detected_action == 'stop' and ('listen' in command or 'voice' in command):
        return jsonify({"status": "success", "action": "stop_listening"})
    
    # Special handling for "main menu" command
    if detected_action == 'main_menu':
        text_to_speech("Returning to main menu.")
        return jsonify({
            "status": "success", 
            "action": "listen_for_commands",
            "reload_page": True,
            "force_redirect": True
        })
    
    # Special handling for combinations of "read" with poorly pronounced "sent"
    if 'read' in command:
        for sent_variant in ['sent', 'send', 'saint', 'scent', 'cent', 'sand', 'spend']:
            if sent_variant in command:
                detected_action = 'read_sent'
                break
    
    # Return the detected action
    if detected_action:
        print(f"Detected action: {detected_action}")
        return jsonify({"status": "success", "action": detected_action})
    else:
        text_to_speech("I didn't understand that command. Please try again.")
        print(f"No action detected for command: {command}")
        return jsonify({"status": "error", "message": "Unknown command"})

@app.route('/listen_for_commands', methods=['POST'])
def listen_for_commands():
    """Continuously listen for commands"""
    try:
        text_to_speech("Listening for commands. Say help for options.")
        result = speech_to_text()
        
        if not result["success"]:
            return jsonify({"status": "error", "message": result["error"]})
        
        command = result["text"].lower()
        print(f"Received command: {command}")
        
        # Help command
        if "help" in command:
            text_to_speech("Available commands: compose email, read inbox, read unread, read sent, read trash, stop listening, return to main menu, logout, or help.")
            return jsonify({"status": "success", "action": "help"})
        
        # Stop listening command
        if ("stop" in command and ("listen" in command or "voice" in command)) or "stop listening" in command:
            text_to_speech("Stopping voice command mode.")
            return jsonify({"status": "success", "action": "stop_listening"})
        
        # Return to main menu command
        if ('return' in command and ('menu' in command or 'main' in command)) or ('main menu' in command):
            text_to_speech("Returning to main menu.")
            return jsonify({
                "status": "success", 
                "action": "listen_for_commands",
                "reload_page": True,
                "force_redirect": True
            })
            
        # Check for main commands - using the existing endpoint
        return voice_command()
        
    except Exception as e:
        error_msg = f"Error listening for commands: {str(e)}"
        print(error_msg)
        return jsonify({"status": "error", "message": error_msg})

if __name__ == '__main__':
    app.run(debug=True) 