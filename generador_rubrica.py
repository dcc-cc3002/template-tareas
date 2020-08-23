"""
Modulo que transforma la rúbrica en excel en un formato fácil de pasar a u-cursos
"""
from typing import Tuple
from pathlib import Path

import os
import pandas as pd

INDICE_NOMBRE_ALUMNO = 0
SUBSEC_CODIGO_FUENTE = ("Funcionalidad", "Diseño")
SECCIONES = ("Código Fuente", "Coverage", "Javadoc", "Resumen")
NOTA = "Nota"
COMENTARIOS = ("Comentarios", "Corrector")
SEC_ADICIONALES = "Adicionales"
COVERAGE = "Porcentaje de coverage"
ROOT = Path(os.path.dirname(os.path.realpath(__file__)))


def get_total(puntaje: str):
    """ Borra el substring `Total: ` del puntaje """
    return puntaje.replace("Total: ", "").replace(",", ".")


def excel_a_string(excel_filename: str) -> Tuple[str, str]:
    """ Convierte la rúbrica a una tupla fácil de pasar a un archivo .txt

    :param excel_filename: el nombre del excel con la rúbrica
    :return: una tupla con el nombre del alumno y los comentarios de revisión
    """
    revision = ""
    nombre_alumno = ""
    nota = ""
    a = pd.read_excel(excel_filename, header=None)
    for index, row in a.iterrows():
        if index == INDICE_NOMBRE_ALUMNO:
            nombre_alumno = f"{row[1]}"
        item = row[0]

        # Puntajes totales de las subsecciones
        if item in SUBSEC_CODIGO_FUENTE:
            revision += "\n" + "=" * 80 + f"\n{item}: {row[2]} / {get_total(row[3])}\n" \
                        + "=" * 80 + "\n"
        # Puntajes totales de las secciones
        elif item in SECCIONES:
            revision += "\n" + "#" * 80 + f"\n{item}: {row[2]} / {get_total(row[3])}\n" \
                        + "#" * 80 + "\n"
        # Nota final
        elif item == NOTA:
            nota = f"{row[3]}"
        # Notas del corrector
        elif item in COMENTARIOS:
            revision += f"\n{item}: {row[1]}"
        # Descuentos adicionales
        elif item == SEC_ADICIONALES:
            revision += "\n" + "#" * 80 + f"\n{item}: {row[2]}\n" + "#" * 80 + "\n"
        # Detalle de los descuentos
        elif index > 1 and row[2] != 0:
            if item == COVERAGE:
                if row[3] != 0:
                    revision += f"\n{item}: {row[2] * 100}% = {row[3]}"
            else:
                revision += f"\n{row[0]}: {row[1]}x{row[2]} = {row[3]}"
    if not nombre_alumno:
        raise Exception("Falta nombre del alumno!!")
    return nombre_alumno, f"Alumno: {nombre_alumno}\nNota: {nota}\n\n{revision}"


if __name__ == '__main__':
    NOMBRE_ALUMNO, REVISION = excel_a_string(f"Rubrica_T2.xlsx")

    with open(f"Comentarios {NOMBRE_ALUMNO}.txt", "w+",
              encoding='utf-8') as comentarios_alumno:
        comentarios_alumno.write(REVISION)
