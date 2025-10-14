import os
import sys
import io
import datetime
import PyPDF2
import shutil
import uuid
from flask import Flask, request, jsonify, send_file
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import reportlab.rl_config

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- Lógica de Creación de Certificados (Tu script original, ahora como una función) ---

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
    # 1. Extraer instrucciones del JSON
    CursoPansatiaOCorto = data['tipo_curso_id']
    Cursovar = data['curso_id']
    estudiantes = data['estudiantes']
    
    # Lógica para determinar si se muestra la firma del instructor
    if CursoPansatiaOCorto == 1: # Pasantía
        InstructorFirma = 1 # Por defecto se muestra la firma (con nombre específico)
    else: # Curso Corto
        # Si se incluye la sección, es 1, si no, es 2.
        InstructorFirma = 1 if data.get('incluir_instructor_seccion', False) else 2

    # 2. Determinar nombre del curso y título
    # (Esta sección es idéntica a tu script original)
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
        # Manejo de error si el ID del curso no es válido
        raise ValueError(f"El ID de curso '{Cursovar}' no es válido.")

    CursoMayus = Curso.upper()

    # 3. Iterar sobre la lista de estudiantes del JSON
    for estudiante_data in estudiantes:
        Estudiante = estudiante_data['nombre']
        Cedula = str(estudiante_data['cedula'])

        # Lógica de fecha (idéntica a la original)
        FechaHoy = datetime.datetime.now()
        MesNum = FechaHoy.month
        Año = str(FechaHoy.year)
        MesesEspañol = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        NombreMes = MesesEspañol[MesNum-1]
        FinaCur = f"Edo. Carabobo, {NombreMes.lower()} {Año}"
        FechaTotalMayus = f"{NombreMes.upper()} {Año}"

        # Lógica para determinar nombre del instructor (idéntica a la original)
        if Cursovar == 1:
            Instructor = "Dr. Lee Nuñez"
        elif Cursovar == 2:
            Instructor = "Lic. Lisbeth De Logan"
        elif Cursovar == 3:
            Instructor = "Lic. Luis Marcano"
        elif Cursovar == 4:
            Instructor = ""
        else: # Cursos Cortos
            if InstructorFirma == 1:
                Instructor = data.get("nombre_instructor", " ") # Obtiene del JSON o deja en blanco
            else:
                Instructor = ""

        # Lógica de duración (idéntica a la original)
        Duracion = "280" if Cursovar in (1, 2, 3, 4) else "36"

        # Lógica de creación de PDF con ReportLab (idéntica a la original)
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

        # Lógica de merge con PyPDF2 (idéntica a la original)
        if InstructorFirma == 1:
            plantilla_path = resource_path("Plantilla certificados General.pdf")
        else:
            plantilla_path = resource_path("Plantilla certificados General sin instructor.pdf")
        
        existing_pdf = PdfReader(open(plantilla_path, "rb"))
        page = existing_pdf.pages[0]
        packet.seek(0)
        new_pdf = PdfReader(packet)
        page.merge_page(new_pdf.pages[0])
        
        # Guardar certificado individual
        certificado_path = os.path.join(output_folder, f"{Estudiante} {Curso}.pdf")
        with open(certificado_path, "wb") as outputStream:
            output = PdfWriter()
            output.add_page(page)
            output.write(outputStream)

    # 4. Unir PDFs si se solicita en el JSON
    if data.get('unir_pdfs', False):
        archivo_salida_unido = os.path.join(output_folder, f"{CursoMayus} {FechaTotalMayus}.pdf")
        pdf_final = PyPDF2.PdfWriter()
        
        for nombre_archivo in sorted(os.listdir(output_folder)):
            if nombre_archivo.endswith('.pdf'):
                ruta_completa = os.path.join(output_folder, nombre_archivo)
                with open(ruta_completa, 'rb') as pdf_individual:
                    pdf_reader = PyPDF2.PdfReader(pdf_individual)
                    for pagina in range(len(pdf_reader.pages)):
                        pdf_final.add_page(pdf_reader.pages[pagina])
        
        with open(archivo_salida_unido, 'wb') as archivo_salida:
            pdf_final.write(archivo_salida)
            
        return archivo_salida_unido
    
    # Si no se unen, no podemos devolver todos los archivos. Este es un caso no manejado.
    # Por defecto, devolvemos la ruta de la carpeta, la API deberá manejar el zippeo si es necesario.
    # O, para simplificar, siempre se devolverá un único archivo unido si hay más de un estudiante.
    raise ValueError("La opción de no unir PDFs no está soportada en esta versión de la API. Por favor, establece 'unir_pdfs' a true.")


# --- Endpoint de la API ---
@app.route('/api/crear_certificados', methods=['POST'])
def manejar_solicitud_certificados():
    # Crear una carpeta temporal única para esta solicitud
    temp_dir = os.path.join(os.path.abspath("."), "temp_output", str(uuid.uuid4()))
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No se recibió un cuerpo JSON."}), 400
        
        # Llamar a la función de lógica principal
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

# --- Punto de Entrada para Ejecutar el Servidor ---
if __name__ == '__main__':
    # Asegurarse de que la carpeta temporal principal exista
    if not os.path.exists("temp_output"):
        os.makedirs("temp_output")
    app.run(debug=True, port=5000)