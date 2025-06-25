from flask import Flask, request, send_file
from docx import Document
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
        for para in doc.paragraphs:
            for key, value in data.items():
                if f"{{{{{key}}}}}" in para.text:
                    para.text = para.text.replace(f"{{{{{key}}}}}", value)

        for image_field in ["location_photo", "magnification_100x", "magnification_500x"]:
            if image_field in request.files:
                doc.add_paragraph(image_field.replace("_", " ").title())
                doc.add_picture(request.files[image_field], width=doc.sections[0].page_width * 0.4)

        temp_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex}.docx")
        doc.save(temp_path)

        # ✅ DO NOT delete the file in this request
        return send_file(temp_path, as_attachment=True, download_name="generated_report.docx")

    except Exception as e:
        app.logger.error(f"Server error: {e}")
        return f"Server error: {e}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
