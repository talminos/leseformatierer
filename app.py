"""
app.py
======

Flask-Anwendung für den Leseformatierer.

Endpunkte
---------
* ``GET  /``        – Startseite mit Upload-Formular.
* ``POST /format``  – Verarbeitet das hochgeladene .docx, sendet die fertige
  Datei als Download zurück und räumt anschließend auf.
* ``GET  /health``  – Einfache Lebendprüfung für Render & Co.

Sicherheit / Hygiene
--------------------
* Maximal 10 MB Upload-Größe (über ``MAX_CONTENT_LENGTH`` erzwungen).
* Nur ``.docx`` wird akzeptiert.
* ``secure_filename`` säubert den Originalnamen.
* Jede Anfrage erhält eine eindeutige Job-ID.
* Hochgeladene und erzeugte Dateien werden nach dem Versand gelöscht;
  zusätzlich räumt eine Cleanup-Funktion verwaiste Dateien aus dem
  ``uploads/``- und ``outputs/``-Ordner.
"""

from __future__ import annotations

import logging
import os
import time
import uuid
from pathlib import Path
from typing import Optional

from flask import (
    Flask,
    after_this_request,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

from formatter import format_document


# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MB
ALLOWED_EXTENSIONS = {".docx"}
CLEANUP_AGE_SECONDS = 60 * 60        # Verwaiste Dateien älter als 1 h löschen

logger = logging.getLogger("leseformatierer")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def create_app() -> Flask:
    """Erzeugt die Flask-App. Nützlich für Tests und Gunicorn."""
    app = Flask(__name__, template_folder="templates")
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-bitte-ersetzen")

    # ----------------------------------------------------------------- Routes
    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/health")
    def health():
        return {"status": "ok"}, 200

    @app.post("/format")
    def format_endpoint():
        # Aufräumen: alte Dateien entfernen.
        cleanup_old_files()

        upload = request.files.get("document")
        if upload is None or upload.filename == "":
            flash("Bitte wählen Sie eine .docx-Datei aus.", "error")
            return redirect(url_for("index"))

        original_name = secure_filename(upload.filename or "")
        if not original_name:
            flash("Der Dateiname ist ungültig.", "error")
            return redirect(url_for("index"))

        ext = Path(original_name).suffix.lower()
        if ext not in ALLOWED_EXTENSIONS:
            flash("Nur .docx-Dateien sind erlaubt.", "error")
            return redirect(url_for("index"))

        # Optionen aus dem Formular
        mode_raw = (request.form.get("mode") or "loose").strip().lower()
        mode = "strict" if mode_raw == "strict" else "loose"
        keep_existing_red = request.form.get("keep_existing_red") == "on"
        only_trigger_paragraphs = request.form.get("only_trigger_paragraphs") == "on"
        speech_units = request.form.get("speech_units") == "on"
        manuscript_layout = request.form.get("manuscript_layout") == "on"

        # Eindeutige Job-ID
        job_id = uuid.uuid4().hex
        stem = Path(original_name).stem
        upload_path = UPLOAD_DIR / f"{job_id}__{original_name}"
        output_filename = f"{stem}_formatiert.docx"
        output_path = OUTPUT_DIR / f"{job_id}__{output_filename}"

        try:
            upload.save(upload_path)
        except Exception:  # noqa: BLE001 – wir wollen wirklich alles fangen
            logger.exception("Speichern des Uploads fehlgeschlagen (Job %s)", job_id)
            flash("Datei konnte nicht gespeichert werden. Bitte erneut versuchen.", "error")
            return redirect(url_for("index"))

        # Größe nochmal prüfen (Schutzgürtel & Hosenträger – Flask prüft bereits).
        try:
            size = upload_path.stat().st_size
        except OSError:
            size = 0
        if size > MAX_UPLOAD_SIZE:
            _safe_unlink(upload_path)
            flash("Die Datei ist größer als 10 MB.", "error")
            return redirect(url_for("index"))

        try:
            format_document(
                str(upload_path),
                str(output_path),
                mode=mode,
                keep_existing_red=keep_existing_red,
                only_trigger_paragraphs=only_trigger_paragraphs,
                speech_units=speech_units,
                manuscript_layout=manuscript_layout,
            )
        except Exception:  # noqa: BLE001
            logger.exception("Formatierung fehlgeschlagen (Job %s)", job_id)
            _safe_unlink(upload_path)
            _safe_unlink(output_path)
            flash(
                "Die Datei konnte nicht verarbeitet werden. "
                "Bitte prüfen Sie, ob es sich um ein gültiges Word-Dokument handelt.",
                "error",
            )
            return redirect(url_for("index"))

        # Upload-Datei wird sofort gelöscht – der Inhalt steckt nun im Output.
        _safe_unlink(upload_path)

        # Ergebnisdatei nach erfolgreichem Versand löschen.
        @after_this_request
        def _cleanup_response(response):  # type: ignore[unused-ignore]
            _safe_unlink(output_path)
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

    # --------------------------------------------------------------- Errors
    @app.errorhandler(RequestEntityTooLarge)
    def too_large(_e):
        flash("Die Datei ist größer als 10 MB.", "error")
        return redirect(url_for("index"))

    @app.errorhandler(404)
    def not_found(_e):
        return ("Seite nicht gefunden.", 404)

    @app.errorhandler(500)
    def server_error(_e):
        return ("Interner Serverfehler. Bitte später erneut versuchen.", 500)

    return app


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------


def _safe_unlink(path: Path) -> None:
    """Löscht eine Datei, ohne Fehler nach außen zu reichen."""
    try:
        if path.exists():
            path.unlink()
    except OSError:
        logger.warning("Datei konnte nicht gelöscht werden: %s", path)


def cleanup_old_files(max_age_seconds: Optional[int] = None) -> int:
    """Räumt verwaiste Uploads/Outputs auf. Gibt die Anzahl gelöschter Dateien zurück."""
    if max_age_seconds is None:
        max_age_seconds = CLEANUP_AGE_SECONDS
    now = time.time()
    removed = 0
    for folder in (UPLOAD_DIR, OUTPUT_DIR):
        try:
            for entry in folder.iterdir():
                if not entry.is_file():
                    continue
                try:
                    age = now - entry.stat().st_mtime
                except OSError:
                    continue
                if age > max_age_seconds:
                    _safe_unlink(entry)
                    removed += 1
        except FileNotFoundError:
            continue
    return removed


# WSGI-Einstiegspunkt für gunicorn (``gunicorn app:app``).
app = create_app()


if __name__ == "__main__":
    # Lokaler Entwicklungsserver. Für die Produktion ``gunicorn app:app``.
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5000")), debug=True)
