# app.py
from flask import Flask
from excel_merger.routes import excel_merger_bp
import config
import os

app = Flask(__name__)
app.config.from_object(config)

def clear_uploads_on_start():
    folder = config.UPLOAD_FOLDER
    if os.path.exists(folder):
        for f in os.listdir(folder):
            try:
                os.remove(os.path.join(folder, f))
            except:
                pass

clear_uploads_on_start()

app.register_blueprint(excel_merger_bp)

if __name__ == "__main__":
    app.run(debug=True)
