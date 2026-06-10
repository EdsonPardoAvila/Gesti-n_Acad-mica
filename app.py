from flask import Flask, render_template
from flask_cors import CORS
import os
from modules.database import init_db
from modules.docente_routes import init_docente_routes
from modules.estudiante_routes import init_estudiante_routes

app = Flask(__name__)
CORS(app)

# Configuración
UPLOAD_FOLDER = 'uploads'
REPORT_FOLDER = 'reportes'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Inicializar rutas
init_docente_routes(app)
init_estudiante_routes(app)

# Rutas principales
@app.route('/')
def index():
    return render_template('home.html')

@app.route('/docentes')
def docentes():
    return render_template('docentes.html')

@app.route('/estudiantes')
def estudiantes():
    return render_template('estudiantes.html')

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=5000)