from flask import Flask, request, send_file
from docx import Document
from docx.shared import Inches
import os
import tempfile
import uuid

app = Flask(__name__)

@app.route("/generate-report", methods=["POST"])
def generate_report():
    try:
        data = request.form.to_dict()
        template = request.files.get("template")
        if not template:
            return "Template required", 400

        doc = Document(template)

        field_map = {
            "Replica #": "ReplicaNo",
            "Component / Location": "ComponentLocation",
            "Material of Construction": "Material",
            "Hardness In HB": "Hardness",
            "Observations": "PhotoLocation",
            "Etchant": "Etchant",
            "Microstructure": "Microstructure",
            "Structural Damage Rating": "DamageRating",
            "Life Exhaustion": "LifeExhaustion",
            "Inspection Interval": "InspectionInterval",
            "Result / Remarks": "ResultRemarks"
        }

        image_placeholders = {
            "location_photo": "{{PhotoLocation}}",
            "magnification_100x": "{{Magnification100x}}",
            "magnification_500x": "{{Magnification500x}}"
        }

        def replace_text_placeholders(paragraph):
            full_text = ''.join(run.text for run in paragraph.runs)
            replaced_text = full_text
            for form_key, value in data.items():
                if form_key in field_map:
                    placeholder = f"{{{{{field_map[form_key]}}}}}"
                    replaced_text = replaced_text.replace(placeholder, value)

            if full_text != replaced_text:
                for run in paragraph.runs:
                    run.text = ''
                paragraph.runs[0].text = replaced_text

        for para in doc.paragraphs:
            replace_text_placeholders(para)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_text_placeholders(para)

        # Replace image placeholders
        for para in doc.paragraphs:
            for field, placeholder in image_placeholders.items():
                if placeholder in para.text and field in request.files:
                    para.clear()
                    para.add_run().add_picture(request.files[field], width=Inches(4.0))

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for field, placeholder in image_placeholders.items():
                            if placeholder in para.text and field in request.files:
                                para.clear()
                                para.add_run().add_picture(request.files[field], width=Inches(4.0))

        # 🔁 Replace last 3 real images (not placeholders) at the end of doc
        images_to_replace = [
            request.files.get("location_photo"),
            request.files.get("magnification_100x"),
            request.files.get("magnification_500x")
        ]
        image_idx = 0

        for i, para in enumerate(reversed(doc.paragraphs)):
            if image_idx >= 3:
                break
            if para.runs and any(run._element.xpath(".//w:drawing") for run in para.runs):
                para.clear()
                if images_to_replace[2 - image_idx]:
                    para.add_run().add_picture(images_to_replace[2 - image_idx], width=Inches(4.0))
                image_idx += 1

        # Save
        temp_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex}.docx")
        doc.save(temp_path)
        return send_file(temp_path, as_attachment=True, download_name="generated_report.docx")

    except Exception as e:
        app.logger.error(f"Server error: {e}")
        return f"Server error: {e}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
