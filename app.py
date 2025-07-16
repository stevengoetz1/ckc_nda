from flask import Flask, render_template, request, send_file
from docx import Document
import io

app = Flask(__name__)

def generate_nda(full_name, firm_name):
    first_name = full_name.strip().split()[0]
    template_path = "Project Slab_NDA_Burlington Street Partners.docx"
    doc = Document(template_path)

    for para in doc.paragraphs:
        if "Dear" in para.text:
            para.text = f"Dear {first_name},"
        para.text = para.text.replace("Burlington Street Partners", firm_name)
        para.text = para.text.replace("Finley Bond", full_name)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = cell.text.replace("Burlington Street Partners", firm_name)
                cell.text = cell.text.replace("Finley Bond", full_name)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

@app.route("/", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        full_name = request.form["full_name"]
        firm_name = request.form["firm_name"]
        docx_stream = generate_nda(full_name, firm_name)
        return send_file(docx_stream, as_attachment=True, download_name=f"NDA_{firm_name.replace(' ', '_')}.docx")
    return render_template("form.html")

if __name__ == "__main__":
    app.run(debug=True)
