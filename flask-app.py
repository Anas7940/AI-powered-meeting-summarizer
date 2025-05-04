from flask import Flask, render_template, request, jsonify 
import os
import threading
from appv4 import join_meeting, record_audio_live, transcribe_audio, meeting_minutes, save_as_docx, send_email

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/start_meeting", methods=["POST"])
def start_meeting():
    data = request.json
    meet_link = data.get("meetLink")
    email_list = data.get("emails", [])
    print(f"Received email list: {email_list}")

    if not meet_link:
        app.logger.error("No Google Meet link provided.")
        return jsonify({"error": "Google Meet link is required!"}), 400

    # Configuration parameters
    bot_name = "MeetReed.AI"
    audio_file_path = "audio/meeting_audio_live911.wav"
    output_doc_path = "meeting_minutes911.docx"
    driver = None

    try:
        # Step 1: Join Google Meeting
        app.logger.info("Joining meeting...")
        driver = join_meeting(meet_link, bot_name=bot_name)
        
        # Step 2: Record audio in background
        app.logger.info("Starting audio recording...")
        recording_thread = threading.Thread(
            target=record_audio_live, 
            args=(audio_file_path,),
            kwargs={"duration": 120}
        )
        recording_thread.start()

        # Let recording proceed while handling other tasks
        recording_thread.join()

        # Step 3: Process audio
        app.logger.info("Transcribing audio...")
        transcription = transcribe_audio(audio_file_path)
        
        # Step 4: Generate meeting minutes
        app.logger.info("Generating meeting minutes...")
        minutes = meeting_minutes(transcription)
        
        # Step 5: Save document
        app.logger.info("Saving document...")
        save_as_docx(minutes, output_doc_path)
        
        # Step 6: Email distribution
        if email_list:
            app.logger.info("Sending emails to %d recipients...", len(email_list))
            for email in email_list:
                try:
                    send_email(email.strip(), output_doc_path)
                    app.logger.info("Email sent successfully to: %s", email.strip())
                except Exception as e:
                    app.logger.error("Error sending to %s: %s", email.strip(), str(e))


        return jsonify({
            "transcription": transcription,
            "summary": minutes["abstract_summary"],
            "key_points": minutes["key_points"],
            "action_items": minutes["action_items"],
            "sentiment": minutes["sentiment"],
            "document": output_doc_path
        })

    except Exception as e:
        app.logger.error(f"Main workflow error: {str(e)}", exc_info=True)
        return jsonify({"error": str(e)}), 500

    finally:
        if driver:
            driver.quit()
            app.logger.info("Browser instance closed")

if __name__ == "__main__":
    os.makedirs("audio", exist_ok=True)
    app.run(debug=True)