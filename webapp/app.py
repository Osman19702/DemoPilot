import io
import tempfile
from pathlib import Path
import sys
import inspect

from flask import Flask, render_template, request, send_file

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
	sys.path.insert(0, str(BASE_DIR))

from make_deck import build_deck


SEND_FILE_SUPPORTS_DOWNLOAD_NAME = "download_name" in inspect.signature(send_file).parameters


app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def index():
	errors = []
	form_data = {
		"out_name": request.form.get("out_name", "demo_deck.pptx"),
		"top": request.form.get("top", "8"),
		"title": request.form.get("title", ""),
	}

	if request.method == "POST":
		csv_file = request.files.get("csv_file")
		out_name = (form_data["out_name"] or "").strip()
		top_raw = (form_data["top"] or "").strip()
		title = (form_data["title"] or "").strip()

		if not csv_file or not csv_file.filename:
			errors.append("Please choose a Jira CSV export.")

		if not out_name:
			errors.append("Please provide an output filename.")
		else:
			if not out_name.lower().endswith(".pptx"):
				out_name += ".pptx"
			out_name = Path(out_name).name

		try:
			top_n = int(top_raw) if top_raw else 8
			if top_n <= 0:
				raise ValueError
		except ValueError:
			errors.append("Top-N value must be a positive integer.")
			top_n = 8

		if not errors:
			try:
				with tempfile.TemporaryDirectory() as tmpdir:
					tmp_path = Path(tmpdir)
					upload_path = tmp_path / "upload.csv"
					csv_file.save(upload_path)

					output_path = tmp_path / Path(out_name).name

					build_deck(
						csv_path=upload_path,
						out_path=output_path,
						top_n=top_n,
						title=title or None,
						pi_filter=None,
						template_path=None,
						include_appendix=False,
					)

					pptx_bytes = output_path.read_bytes()
					buffer = io.BytesIO(pptx_bytes)
					buffer.seek(0)

					kwargs = {
						"as_attachment": True,
						"mimetype": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
					}
					if SEND_FILE_SUPPORTS_DOWNLOAD_NAME:
						kwargs["download_name"] = out_name
					else:
						kwargs["attachment_filename"] = out_name

					return send_file(buffer, **kwargs)
			except Exception as exc:
				errors.append(f"Failed to generate deck: {exc}")

	return render_template("index.html", errors=errors, form_data=form_data)


if __name__ == "__main__":
	app.run(debug=True)
