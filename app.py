import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from mnrega_scraper import run_mnrega_scraper

app = Flask(__name__)
app.secret_key = "super-secret-key-change-this"  # for flash messages



@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        work_code = request.form.get("work_code", "").strip()

        if not work_code:
            flash("Please enter a Work Code.", "error")
            return redirect(url_for("index"))

        try:
            # Run scraper (headless)
            output_path = run_mnrega_scraper(work_code, output_dir=".")

            if not os.path.exists(output_path):
                flash("Something went wrong, file not generated.", "error")
                return redirect(url_for("index"))

            # Send file to user
            return send_file(
                output_path,
                as_attachment=True,
                download_name=os.path.basename(output_path),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            print("ERROR:", e)
            flash(f"Error while processing: {e}", "error")
            return redirect(url_for("index"))

    # GET request
    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)