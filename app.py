from flask import Flask, render_template, request, send_file
from docx import Document
import os
import io

app = Flask(__name__)

def replace_text_preserving_formatting(doc, full_name, firm_name):
    first_name = full_name.strip().split()[0]
    for para in doc.paragraphs:
        for run in para.runs:
            if "Finley Bond" in run.text:
                run.text = run.text.replace("Finley Bond", full_name)
            if "Burlington Street Partners" in run.text:
                run.text = run.text.replace("Burlington Street Partners", firm_name)
            if "Dear Finley," in run.text:
                run.text = run.text.replace("Dear Finley", f"Dear {first_name}")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if "Finley Bond" in run.text:
                            run.text = run.text.replace("Finley Bond", full_name)
                        if "Burlington Street Partners" in run.text:
                            run.text = run.text.replace("Burlington Street Partners", firm_name)
                        if "Dear Finley," in run.text:
                            run.text = run.text.replace("Dear Finley", f"Dear {first_name}")
    return doc

@app.route("/", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        full_name = request.form["full_name"]
        firm_name = request.form["firm_name"]

        doc = Document("Project Slab_NDA_Burlington Street Partners.docx")
        updated_doc = replace_text_preserving_formatting(doc, full_name, firm_name)

        output_stream = io.BytesIO()
        updated_doc.save(output_stream)
        output_stream.seek(0)
        return send_file(output_stream, as_attachment=True, download_name=f"NDA_{firm_name.replace(' ', '_')}.docx")

    return render_template("form.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
