from pathlib import Path
from uuid import uuid4
import logging
import shutil
import traceback

from flask import Flask, request, render_template, send_file, url_for
from werkzeug.utils import secure_filename

from process_data import process_portfolio



# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

logging.basicConfig(level=logging.INFO)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"

UPLOAD_DIR = BASE_DIR / "runtime" / "uploads"
OUTPUT_DIR = BASE_DIR / "runtime" / "outputs"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(TEMPLATE_DIR))

app.config["ALLOWED_EXTENSIONS"] = {"xls", "xlsx"}
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

DEFAULT_SHEET_NAME = "Sheet 1"


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def allowed_file(filename: str) -> bool:
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]
    )


def render_home(success=None, error=None, download_url=None):
    return render_template(
        "index.html",
        success=success,
        error=error,
        download_url=download_url,
    )


def friendly_error_message(exc: Exception) -> str:
    """Translate technical exceptions into plain-English messages for users."""
    message = str(exc).strip()

    if isinstance(exc, ValueError) and message.startswith("Missing required columns:"):
        missing = message.removeprefix("Missing required columns:").strip()
        return (
            "Your file is missing required information. "
            f"Please add these columns and try again: {missing}."
        )

    if isinstance(exc, FileNotFoundError):
        return "We could not create the processed file. Please try uploading the file again."

    if "sheet" in message.lower() and "not found" in message.lower():
        return (
            "The sheet name you entered was not found in the Excel file. "
            "Please check the sheet name and try again."
        )

    if message:
        return f"Something went wrong while processing your file: {message}"

    return "Something went wrong while processing your file. Please try again."


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    return render_home()

@app.route("/upload", methods=["POST"])
def upload_file():
    uploaded_file = request.files.get("file")

    if uploaded_file is None or uploaded_file.filename == "":
        return render_home(
            error="No file selected. Please upload an Excel file."
        )

    if not allowed_file(uploaded_file.filename):
        return render_home(
            error="Invalid file type. Please upload only .xls or .xlsx files."
        )

    sheet_name = request.form.get("sheet_name", DEFAULT_SHEET_NAME).strip()
    if not sheet_name:
        sheet_name = DEFAULT_SHEET_NAME

    original_filename = secure_filename(uploaded_file.filename)
    file_stem = Path(original_filename).stem
    file_suffix = Path(original_filename).suffix

    run_id = uuid4().hex

    run_upload_dir = UPLOAD_DIR / run_id
    run_output_dir = OUTPUT_DIR / run_id

    run_upload_dir.mkdir(parents=True, exist_ok=True)
    run_output_dir.mkdir(parents=True, exist_ok=True)

    input_path = run_upload_dir / f"{file_stem}{file_suffix}"
    output_filename = f"processed_{file_stem}.xlsx"
    output_path = run_output_dir / output_filename

    try:
        uploaded_file.save(input_path)

        process_portfolio(
            input_path=input_path,
            output_path=output_path,
            sheet_name=sheet_name,
        )

        if not output_path.exists():
            raise FileNotFoundError("Processing finished, but output file was not created.")

        download_url = url_for(
            "download_file",
            run_id=run_id,
            filename=output_filename,
        )

        return render_home(
            success="File processed successfully.",
            error=None,
            download_url=download_url,
        )

    except Exception as exc:
        logging.error("Error during file processing")
        logging.error(traceback.format_exc())

        shutil.rmtree(run_upload_dir, ignore_errors=True)
        shutil.rmtree(run_output_dir, ignore_errors=True)

        return render_home(
            success=None,
            error=friendly_error_message(exc),
            download_url=None,
        )


@app.route("/download/<run_id>/<filename>", methods=["GET"])
def download_file(run_id: str, filename: str):
    safe_filename = secure_filename(filename)
    file_path = OUTPUT_DIR / run_id / safe_filename

    if not file_path.exists():
        return render_home(
            error="Processed file was not found. Please upload and process the file again."
        )

    return send_file(
        file_path,
        as_attachment=True,
        download_name=safe_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/health", methods=["GET"])
def health():
    return "App is running"


@app.errorhandler(413)
def file_too_large(error):
    return render_home(
        error="Uploaded file is too large. Maximum allowed file size is 100 MB."
    ), 413


# ---------------------------------------------------------------------------
# Local run
# ---------------------------------------------------------------------------
@app.route("/download-demo", methods=["GET"])
def download_demo_file():
    demo_path = BASE_DIR / "static" / "demo" / "ProjectX_raw_input_demo_for_upload.xlsx"

    if not demo_path.exists():
        return render_home(
            error="Demo input file is not available on the server."
        )

    return send_file(
        demo_path,
        as_attachment=True,
        download_name="ProjectX_raw_input_demo_for_upload.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True, use_reloader=False)
