"""
AWS Lambda handler – compila un template PPTX sostituendo i placeholder
con i valori ricevuti nel JSON di input, salva il risultato su S3
e restituisce un pre-signed URL per il download.
"""

import json
import os
import copy
import re
import tempfile

import boto3
from pptx import Presentation

S3_BUCKET = os.environ["S3_BUCKET"]
PRESIGNED_URL_EXPIRATION = int(os.environ.get("PRESIGNED_URL_EXPIRATION", "3600"))

s3 = boto3.client("s3")

PLACEHOLDER_RE = re.compile(r"\{\{(\w+)\}\}")


def replace_text_in_paragraph(paragraph, placeholders: dict):
    """Sostituisce i placeholder {{key}} nel testo di un paragrafo.

    python-pptx spezza il testo in più *run* per motivi di formattazione,
    quindi ricostruiamo il testo completo, facciamo la sostituzione e poi
    lo rimappiamo nei run originali preservando lo stile.
    """
    full_text = "".join(run.text for run in paragraph.runs)
    if not PLACEHOLDER_RE.search(full_text):
        return

    for key, value in placeholders.items():
        full_text = full_text.replace("{{" + key + "}}", str(value))

    # Redistribuisci il testo nei run esistenti
    if len(paragraph.runs) == 1:
        paragraph.runs[0].text = full_text
    elif paragraph.runs:
        paragraph.runs[0].text = full_text
        for run in paragraph.runs[1:]:
            run.text = ""


def replace_placeholders(prs: Presentation, placeholders: dict):
    """Scorre tutte le slide e sostituisce i placeholder."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                replace_text_in_paragraph(paragraph, placeholders)


def lambda_handler(event, context):
    # --- Parsing input ---
    if isinstance(event.get("body"), str):
        body = json.loads(event["body"])
    else:
        body = event

    template_key = body["template_s3_key"]
    output_key = body.get("output_s3_key", f"output/{context.aws_request_id}.pptx")
    placeholders = body.get("placeholders", {})

    # --- Download template da S3 ---
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_in:
        s3.download_fileobj(S3_BUCKET, template_key, tmp_in)
        tmp_in_path = tmp_in.name

    # --- Compila il template ---
    prs = Presentation(tmp_in_path)
    replace_placeholders(prs, placeholders)

    # --- Salva il risultato ---
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_out:
        prs.save(tmp_out.name)
        tmp_out_path = tmp_out.name

    # --- Upload su S3 ---
    with open(tmp_out_path, "rb") as f:
        s3.upload_fileobj(f, S3_BUCKET, output_key, ExtraArgs={
            "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        })

    # --- Cleanup temp files ---
    os.unlink(tmp_in_path)
    os.unlink(tmp_out_path)

    # --- Pre-signed URL ---
    presigned_url = s3.generate_presigned_url(
        "get_object",
        Params={"Bucket": S3_BUCKET, "Key": output_key},
        ExpiresIn=PRESIGNED_URL_EXPIRATION,
    )

    return {
        "statusCode": 200,
        "body": json.dumps({
            "message": "PPT compilato con successo",
            "output_s3_key": output_key,
            "download_url": presigned_url,
        }),
    }
