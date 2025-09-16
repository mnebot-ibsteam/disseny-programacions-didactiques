from flask import Flask, render_template, request, redirect, jsonify
from load_data import carregar_estructura, carregar_sabers
from utils import generar_docx
import os
import json
import uuid
import time
import pandas as pd  # <- Per poder carregar el CSV a les noves rutes


app = Flask(__name__)

estructura = carregar_estructura('data/competencies.csv')
sabers_data = carregar_sabers('data/sabers_concrecio.csv')
sabers_df = pd.read_csv("data/sabers_concrecio.csv", sep=':', encoding="utf-8", quotechar='"')  # <- Carregam el DataFrame per a l'API

TEMP_DIR = 'static/'
FILE_TTL = 3600  # Temps de vida dels fitxers en segons (ex: 1 hora)

def netejar_fitxers_temporals():
    now = time.time()
    for fitxer in os.listdir(TEMP_DIR):
        if fitxer.startswith('informe_') and fitxer.endswith('.docx'):
            ruta_fitxer = os.path.join(TEMP_DIR, fitxer)
            if os.path.isfile(ruta_fitxer) and now - os.path.getmtime(ruta_fitxer) > FILE_TTL:
                try:
                    os.remove(ruta_fitxer)
                except Exception as e:
                    print(f"Error eliminant fitxer temporal {fitxer}: {e}")

@app.route('/', methods=['GET', 'POST'])
def index():
    now = time.time()
    carpeta = 'static'
    for nom in os.listdir(carpeta):
        if nom.endswith('.docx'):
            ruta = os.path.join(carpeta, nom)
            if os.path.isfile(ruta) and now - os.path.getmtime(ruta) > 3600:
                os.remove(ruta)
                
    nivells = list(estructura.keys())
    if request.method == 'POST':
        nivell = request.form['nivell']
        return redirect(f'/formulari?nivell={nivell}')
    return render_template('index.html', nivells=nivells)

@app.route('/formulari', methods=['GET', 'POST'])
def formulari():
    global estructura, sabers_data
    nivell = request.args.get('nivell')
    cursos = list(estructura[nivell].keys())
    estructura_json = estructura[nivell]
    sabers_json = sabers_data[nivell]

    if request.method == 'POST':
        dades = {
            'nivell': request.form['nivell'],
            'curs': request.form['curs'],
            'materia': request.form['materia'],
            'competencies': request.form.getlist('competencies'),
            'criteris': request.form.getlist('criteris'),
            'sabers': request.form.getlist('sabers'),
            'concrecions': request.form.getlist('concrecions')
        }
        nom_fitxer = f"static/informe_{uuid.uuid4()}.docx"
        generar_docx(dades, nom_fitxer)
        return render_template('resultat.html', dades=dades, doc=nom_fitxer)

    return render_template('formulari.html',
                           nivell=nivell,
                           cursos=cursos,
                           estructura=estructura[nivell],
                           estructura_json=json.dumps(estructura[nivell], ensure_ascii=False),
                           sabers_json=json.dumps(sabers_data[nivell], ensure_ascii=False))

# 🔽 NOVES RUTES API 🔽

@app.route('/get_sabers')
def get_sabers():
    nivell = (request.args.get('nivell') or '').strip()
    curs = (request.args.get('curs') or '').strip()
    materia = (request.args.get('materia') or '').strip()
    if not curs or not materia:
        return jsonify([])

    df = sabers_df.copy()
    # Normalitza espais i majúscules
    df['Nivell'] = df['Nivell'].astype(str).str.strip()
    df['Curs'] = df['Curs'].astype(str).str.strip()
    df['Matèria'] = df['Matèria'].astype(str).str.strip()

    sabers = (
        df[
            (df['Nivell'].str.casefold() == nivell.casefold()) &
            (df['Curs'] == curs) &
            (df['Matèria'] == materia)
        ]['Saber bàsic']
        .dropna()
        .drop_duplicates()
        .tolist()
    )
    return jsonify(sabers)


@app.route('/get_concrecions')
def get_concrecions():
    saber = request.args.get('saber')
    if not saber:
        return jsonify([])

    concrecions = sabers_df[
        sabers_df['Saber bàsic'] == saber
    ]['Concreció'].dropna().tolist()

    return jsonify(concrecions)

if __name__ == '__main__':
    app.run(debug=True)
