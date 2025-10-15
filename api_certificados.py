import os
import sys
import io
import datetime
import fitz  # PyMuPDF
import shutil
import uuid
from flask import Flask, request, jsonify, send_file
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import reportlab.rl_config

app = Flask(__name__)

def resource_path(relative_path):
    """ Obtiene la ruta absoluta del recurso, funciona para desarrollo y para PyInstaller. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def crear_certificados_desde_json(data, output_folder):
    """
    Función principal que contiene tu lógica de generación de certificados.
    Toma los datos de un diccionario JSON y una carpeta de salida.
    """
    CursoPansatiaOCorto = data['tipo_curso_id']
    estudiantes = data['estudiantes']

    # Lógica para manejar cursos nuevos, ahora con ID 0
    if CursoPansatiaOCorto == 0:
        Curso = data['curso_nombre']
        TituloCurso = data['curso_nombre']
        Duracion = str(data['duracion_horas'])
        InstructorFirma = 1 if data.get('incluir_instructor_seccion', False) else 2
    else: # Lógica para cursos existentes (Pasantía o Corto)
        Cursovar = data['curso_id']
        if Cursovar == 1:
            Curso = "Farmacia"
            TituloCurso = "Auxiliar de Farmacia"
        elif Cursovar == 2:
            Curso = "Enfermeria"
            TituloCurso = "Auxiliar de Enfermería"
        elif Cursovar == 3:
            Curso = "Bionalisis"
            TituloCurso = "Asistente de Laboratorio Clínico"
        elif Cursovar == 4:
            Curso = "Administracion"
            TituloCurso = "Asistente Administrativo Contable"
        elif Cursovar == 5:
            Curso = "Computacion"
            TituloCurso = Curso
        elif Cursovar == 6:
            Curso = "Office"
            TituloCurso = Curso
        elif Cursovar == 7:
            Curso = "Electronica"
            TituloCurso = Curso
        elif Cursovar == 8:
            Curso = "Barberia"
            TituloCurso = "Barbería"
        elif Cursovar == 9:
            Curso = "Sistema de Uñas"
            TituloCurso = Curso
        elif Cursovar == 10:
            Curso = "Depilacion"
            TituloCurso = "Depilación Facial"
        else:
            raise ValueError(f"El ID de curso '{Cursovar}' no es válido.")

        Duracion = "280" if Cursovar in (1, 2, 3, 4) else "36"
        InstructorFirma = 1 if CursoPansatiaOCorto == 1 else (1 if data.get('incluir_instructor_seccion', False) else 2)

    CursoMayus = Curso.upper()

    for estudiante_data in estudiantes:
        Estudiante = estudiante_data['nombre']
        Cedula = str(estudiante_data['cedula'])

        FechaHoy = datetime.datetime.now()
        MesNum = FechaHoy.month
        Año = str(FechaHoy.year)
        MesesEspañol = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        NombreMes = MesesEspañol[MesNum-1]
        FinaCur = f"Edo. Carabobo, {NombreMes.lower()} {Año}"
        FechaTotalMayus = f"{NombreMes.upper()} {Año}"

        # Lógica del instructor para curso nuevo (ID 0)
        if CursoPansatiaOCorto == 0:
             Instructor = data.get("nombre_instructor", " ") if InstructorFirma == 1 else ""
        elif CursoPansatiaOCorto == 1: # Pasantía
            curso_id_temp = data['curso_id']
            if curso_id_temp == 1: Instructor = "Dr. Lee Nuñez"
            elif curso_id_temp == 2: Instructor = "Lic. Lisbeth De Logan"
            elif curso_id_temp == 3: Instructor = "Lic. Luis Marcano"
            else: Instructor = ""
        else: # Curso Corto
            Instructor = data.get("nombre_instructor", " ") if InstructorFirma == 1 else ""

        packet = io.BytesIO()
        width, height = letter
        c = canvas.Canvas(packet, pagesize=(width*5.25, height*3.14))
        reportlab.rl_config.warnOnMissingFontGlyphs = 0
        pdfmetrics.registerFont(TTFont('Edward', resource_path('Edwardian.ttf')))
        pdfmetrics.registerFont(TTFont('MyriadBold', resource_path('MYRIADPRO-BOLD.ttf')))
        pdfmetrics.registerFont(TTFont('Myriad', resource_path('MYRIADPRO-REGULAR.ttf')))
        pdfmetrics.registerFont(TTFont('TimeNewBold', resource_path('times new roman bold.ttf')))
        c.setFillColorRGB(0,0,0)
        c.setFont('Edward', 72)
        c.drawCentredString(410, 300, Estudiante)
        c.setFont('MyriadBold', 14.5)
        c.drawCentredString(375, 230, Cedula)
        c.setFont('TimeNewBold', 16)
        c.drawCentredString(390, 35, FinaCur)
        c.setFont('Myriad', 14.1)
        c.drawCentredString(390, 87, Instructor)
        c.setFont('Myriad', 12.1)
        c.drawCentredString(673, 163, Duracion)
        c.setFont('MyriadBold', 17.5)
        c.drawCentredString(410, 211, TituloCurso)
        c.save()

        if InstructorFirma == 1:
            plantilla_path = resource_path("Plantilla certificados General.pdf")
        else:
            plantilla_path = resource_path("Plantilla certificados General sin instructor.pdf")

        existing_pdf = PdfReader(open(plantilla_path, "rb"))
        page = existing_pdf.pages[0]
        packet.seek(0)
        new_pdf = PdfReader(packet)
        page.merge_page(new_pdf.pages[0])

        certificado_path = os.path.join(output_folder, f"{Estudiante} {Curso}.pdf")
        with open(certificado_path, "wb") as outputStream:
            output = PdfWriter()
            output.add_page(page)
            output.write(outputStream)

    if data.get('unir_pdfs', False):
        archivo_salida_unido = os.path.join(output_folder, f"{CursoMayus} {FechaTotalMayus}.pdf")
        pdf_final = fitz.open()

        for nombre_archivo in sorted(os.listdir(output_folder)):
            if nombre_archivo.endswith('.pdf'):
                ruta_completa = os.path.join(output_folder, nombre_archivo)
                with fitz.open(ruta_completa) as pdf_individual:
                    pdf_final.insert_pdf(pdf_individual)

        pdf_final.save(archivo_salida_unido)
        pdf_final.close()

        if data.get('eliminar_individuales', False):
            for nombre_archivo in os.listdir(output_folder):
                if nombre_archivo.endswith(f'{Curso}.pdf'):
                    os.remove(os.path.join(output_folder, nombre_archivo))

        return archivo_salida_unido

    raise ValueError("La opción de no unir PDFs no está soportada. Por favor, establece 'unir_pdfs' a true.")


@app.route('/api/crear_certificados', methods=['POST'])
def manejar_solicitud_certificados():
    """Endpoint principal para recibir solicitudes de creación de certificados."""
    temp_dir = os.path.join(os.path.abspath("."), "temp_output", str(uuid.uuid4()))
    os.makedirs(temp_dir, exist_ok=True)

    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No se recibió un cuerpo JSON."}), 400

        # Validación para Cursos Nuevos, ahora con ID 0
        tipo_curso = data.get('tipo_curso_id')
        if tipo_curso == 0:
            if 'curso_nombre' not in data or 'duracion_horas' not in data:
                return jsonify({"error": "Para tipo_curso_id 0, se requieren 'curso_nombre' y 'duracion_horas'."}), 400
        elif tipo_curso in [1, 2]:
            if 'curso_id' not in data:
                 return jsonify({"error": "Para tipo_curso_id 1 o 2, se requiere 'curso_id'."}), 400
        else:
            return jsonify({"error": f"El tipo_curso_id '{tipo_curso}' no es válido."}), 400

        ruta_pdf_final = crear_certificados_desde_json(data, temp_dir)

        if not os.path.exists(ruta_pdf_final):
            return jsonify({"error": "El archivo final no pudo ser generado."}), 500

        return send_file(
            ruta_pdf_final,
            as_attachment=True,
            download_name=os.path.basename(ruta_pdf_final)
        )

    except ValueError as ve:
        return jsonify({"error": f"Dato inválido: {str(ve)}"}), 400
    except Exception as e:
        return jsonify({"error": f"Error interno del servidor: {str(e)}"}), 500
    finally:
        # Limpiar la carpeta temporal después de enviar la respuesta
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == '__main__':
    # Asegurarse de que la carpeta temporal principal exista
    if not os.path.exists("temp_output"):
        os.makedirs("temp_output")
    app.run(debug=True, port=5000)