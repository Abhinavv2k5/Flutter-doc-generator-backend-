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

        # Mapping from Flutter field labels → Template placeholders
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

        # 🔁 Replace text placeholders
        for para in doc.paragraphs:
            for form_key, value in data.items():
                if form_key in field_map:
                    placeholder = f"{{{{{field_map[form_key]}}}}}"
                    if placeholder in para.text:
                        para.text = para.text.replace(placeholder, value)

        # 🖼 Replace image placeholders inline
        image_placeholders = {
            "location_photo": "{{PhotoLocation}}",
            "magnification_100x": "{{Magnification100x}}",
            "magnification_500x": "{{Magnification500x}}"
        }

        for para in doc.paragraphs:
            for field, placeholder in image_placeholders.items():
                if placeholder in para.text and field in request.files:
                    para.clear()  # Remove placeholder text
                    para.add_run().add_picture(request.files[field], width=Inches(4.0))

        # Save final output
        temp_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex}.docx")
        doc.save(temp_path)

        return send_file(temp_path, as_attachment=True, download_name="generated_report.docx")

    except Exception as e:
        app.logger.error(f"Server error: {e}")
        return f"Server error: {e}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
