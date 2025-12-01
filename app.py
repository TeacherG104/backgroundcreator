from flask import Flask, render_template_string, request, Response
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import requests

app = Flask(__name__)

SERVICE_ACCOUNT_FILE = "background-builder-479821-c88533cc7183.json"
SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive"
]

HTML_FORM = """
<!doctype html>
<html>
<head>
  <title>Slides Background Builder</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2em; }
    h1 { color: #2c3e50; }
    .log { background: #f4f4f4; border: 1px solid #ccc; padding: 1em; height: 300px; overflow-y: auto; }
    .line { margin: 0.2em 0; }
    .complete { color: green; font-weight: bold; }
    .running { color: #2980b9; }
    .progress-container { width: 100%; background: #ddd; margin-top: 1em; }
    .progress-bar { width: 0%; height: 25px; background: #27ae60; text-align: center; color: white; }
  </style>
</head>
<body>
  <h1>Update Slide Backgrounds</h1>
  <form method="post">
    Deck ID: <input type="text" name="deck"><br>
    Folder ID: <input type="text" name="folder"><br><br>
    <input type="submit" value="Run">
  </form>
  <div class="progress-container">
    <div class="progress-bar" id="progress-bar">0%</div>
  </div>
  <div class="log" id="log"></div>

  <script>
    function runStream(totalSlides) {
      const log = document.getElementById("log");
      const bar = document.getElementById("progress-bar");
      let completed = 0;

      const evtSource = new EventSource("/progress");
      evtSource.onmessage = function(e) {
        const div = document.createElement("div");
        div.className = "line";
        if (e.data.includes("Slide")) {
          div.classList.add("complete");
          completed++;
          let percent = Math.round((completed / totalSlides) * 100);
          bar.style.width = percent + "%";
          bar.textContent = percent + "%";
        } else {
          div.classList.add("running");
        }
        div.textContent = e.data;
        log.appendChild(div);
        log.scrollTop = log.scrollHeight;
      };
    }
  </script>
</body>
</html>
"""

def get_services():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    slides_service = build("slides", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return slides_service, drive_service

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        deck_id = request.form["deck"]
        folder_id = request.form["folder"]

        def generate():
            slides_service, drive_service = get_services()
            presentation = slides_service.presentations().get(presentationId=deck_id).execute()
            slides = presentation["slides"]
            total_slides = len(slides)

            yield f"data: Starting update for {total_slides} slides...\n\n"

            for i, slide in enumerate(slides, start=1):
                slide_id = slide["objectId"]

                # Export thumbnail
                thumb = slides_service.presentations().pages().getThumbnail(
                    presentationId=deck_id,
                    pageObjectId=slide_id,
                    thumbnailProperties_thumbnailSize="LARGE"
                ).execute()
                image_url = thumb["contentUrl"]

                # Download image
                filename = f"slide_{i}.png"
                resp = requests.get(image_url)
                with open(filename, "wb") as f:
                    f.write(resp.content)

                # Upload to folder
                file_metadata = {"name": filename, "mimeType": "image/png", "parents": [folder_id]}
                media = MediaFileUpload(filename, mimetype="image/png")
                uploaded = drive_service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields="id, webContentLink",
                    supportsAllDrives=True
                ).execute()

                file_link = uploaded["webContentLink"]

                # Share file
                drive_service.permissions().create(
                    fileId=uploaded["id"],
                    body={"role": "reader", "type": "anyone"},
                    supportsAllDrives=True
                ).execute()

                # Update background
                requests_body = {
                    "requests": [
                        {
                            "updatePageProperties": {
                                "objectId": slide_id,
                                "pageProperties": {
                                    "pageBackgroundFill": {
                                        "stretchedPictureFill": {"contentUrl": file_link}
                                    }
                                },
                                "fields": "pageBackgroundFill"
                            }
                        }
                    ]
                }
                slides_service.presentations().batchUpdate(
                    presentationId=deck_id, body=requests_body
                ).execute()

                yield f"data: Slide {i}... complete\n\n"

            yield "data: ✅ All slides updated.\n\n"

        return Response(generate(), mimetype="text/event-stream")

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    app.run(debug=True, threaded=True)