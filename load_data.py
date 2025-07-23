
import csv
from collections import defaultdict

def carregar_estructura(ruta_csv):
    estructura = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))

    with open(ruta_csv, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=':')
        for fila in reader:
            nivell = fila['Nivell']
            curs = fila['Curs']
            materia = fila['Matèria']
            compet = fila['Competència']
            criteri = fila['Criteri']
            estructura[nivell][curs][materia][compet].append(criteri)

    return estructura

def carregar_sabers(ruta_csv):
    sabers = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))

    with open(ruta_csv, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=':')
        for fila in reader:
            nivell = fila['Nivell']
            curs = fila['Curs']
            materia = fila['Matèria']
            saber = fila['Saber bàsic']
            sabers[nivell][curs][materia].append(saber)

    return sabers
