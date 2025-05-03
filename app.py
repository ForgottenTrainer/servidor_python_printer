
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_from_directory
import os
import sys
import subprocess
import platform
import tempfile
import win32print
import win32api
import win32ui
from win32con import LOGPIXELSX, LOGPIXELSY
from werkzeug.utils import secure_filename
import pythoncom
import time
import json
import datetime
import shutil
from pathlib import Path
from PIL import Image
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.secret_key = 'clave_secreta_para_flask'

# Archivo para almacenar el historial de impresión
HISTORY_FILE = 'print_history.json'

# Asegurar que existe el directorio de uploads
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Extensiones permitidas
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc', 'txt', 'jpg', 'jpeg', 'png'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
           
# Asegurar que existen los directorios necesarios
def setup_directories():
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs('static', exist_ok=True)
    
    # Si no existe el directorio de templates, crearlo
    if not os.path.exists('templates'):
        os.makedirs('templates')
        
    # Si no existe el archivo index.html en templates, crear uno básico
    if not os.path.exists(os.path.join('templates', 'index.html')):
        with open(os.path.join('templates', 'index.html'), 'w', encoding='utf-8') as f:
            f.write("""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="refresh" content="0; url=/static/index.html">
    <title>Redireccionando...</title>
</head>
<body>
    <p>Redireccionando a la interfaz...</p>
</body>
</html>""")
    
    # Si no existe el archivo de historial, crearlo vacío
    if not os.path.exists(HISTORY_FILE):
        save_print_history([])

def convert_to_pdf(input_file):
    """Convierte diferentes formatos a PDF si es necesario"""
    filename, file_extension = os.path.splitext(input_file)
    output_pdf = f"{filename}.pdf"
    
    # Si ya es PDF, devolvemos la misma ruta
    if file_extension.lower() == '.pdf':
        return input_file
    
    # Convertir DOCX a PDF con alternativas que no usan comtypes
    elif file_extension.lower() in ['.docx', '.doc']:
        try:
            # Intentamos usar LibreOffice si está disponible
            for libreoffice_path in [
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
            ]:
                if os.path.exists(libreoffice_path):
                    cmd = [
                        libreoffice_path,
                        '--headless',
                        '--convert-to', 'pdf',
                        '--outdir', os.path.dirname(input_file),
                        input_file
                    ]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    return output_pdf
            
            # Alternativa: intentar usar soffice en PATH
            if shutil.which('soffice'):
                cmd = [
                    'soffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(input_file),
                    input_file
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                return output_pdf
                
            # Si llegamos aquí, informamos que no podemos convertir
            print(f"No se pudo convertir el documento Word: LibreOffice no está instalado")
            return input_file
            
        except Exception as e:
            print(f"Error al convertir documento: {e}")
            return input_file
    
    # Para archivos de texto plano, convertimos a PDF
    elif file_extension.lower() == '.txt':
        try:
            # Crear un PDF desde el archivo de texto
            c = canvas.Canvas(output_pdf, pagesize=letter)
            text_file = open(input_file, 'r', encoding='utf-8', errors='ignore')
            y = 750  # posición inicial en Y
            for line in text_file:
                if y < 50:  # Si alcanzamos el final de la página
                    c.showPage()  # Nueva página
                    y = 750
                c.drawString(50, y, line.strip())
                y -= 12  # Espacio entre líneas
            text_file.close()
            c.save()
            return output_pdf
        except Exception as e:
            print(f"Error al convertir TXT a PDF: {e}")
            return input_file
    
    # Para imágenes, convertimos a PDF
    elif file_extension.lower() in ['.jpg', '.jpeg', '.png']:
        try:
            img = Image.open(input_file)
            img_rgb = img.convert('RGB')
            img_rgb.save(output_pdf, 'PDF')
            return output_pdf
        except Exception as e:
            print(f"Error al convertir imagen a PDF: {e}")
            return input_file
    
    return input_file

def print_file_direct(file_path, printer_name=None):
    """Imprime un archivo directamente sin mostrar diálogos"""
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()
    
    try:
        file_path = os.path.abspath(file_path)
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # Para PDF, usamos una impresión más directa
        if file_ext == '.pdf':
            # MÉTODO 1: Imprimir usando GhostScript (si está instalado)
            gs_paths = [
                shutil.which('gsprint.exe'),
                r'C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe',
                r'C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe', 
                r'C:\Program Files (x86)\gs\gs9.56.1\bin\gswin32c.exe'
            ]
            
            for gs_path in gs_paths:
                if gs_path and os.path.exists(gs_path):
                    try:
                        if 'gsprint' in gs_path:
                            cmd = [gs_path, '-printer', printer_name, file_path]
                            print(f"Ejecutando: {' '.join(cmd)}")
                            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            return True
                        else:
                            cmd = [
                                gs_path, 
                                '-dPrinted', '-dBATCH', '-dNOPAUSE', '-dNOPROMPT',
                                f'-sOutputFile=%printer%{printer_name}', 
                                '-sDEVICE=mswinpr2',
                                file_path
                            ]
                            print(f"Ejecutando: {' '.join(cmd)}")
                            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            return True
                    except Exception as e:
                        print(f"Error al imprimir usando GhostScript: {e}")
                        # Continuamos con otros métodos si este falla
            
            # MÉTODO 2: Usar SumatraPDF (si está disponible)
            sumatra_paths = [
                r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
                r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe'
            ]
            
            for sumatra_path in sumatra_paths:
                if os.path.exists(sumatra_path):
                    try:
                        cmd = [sumatra_path, '-print-to', printer_name, '-silent', file_path]
                        print(f"Ejecutando: {' '.join(cmd)}")
                        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        return True
                    except Exception as e:
                        print(f"Error al imprimir con SumatraPDF: {e}")
                        # Continuamos con otros métodos si este falla
            
            # MÉTODO 3: Usar PDFtoPrinter (si está disponible)
            pdftoprinterpaths = [
                r'C:\Program Files\PDFtoPrinter\PDFtoPrinter.exe',
                r'C:\Program Files (x86)\PDFtoPrinter\PDFtoPrinter.exe',
                'PDFtoPrinter.exe'  # Si está en PATH
            ]
            
            for pdftoprinter in pdftoprinterpaths:
                if os.path.exists(pdftoprinter) or shutil.which(pdftoprinter):
                    try:
                        cmd = [pdftoprinter, file_path, printer_name]
                        print(f"Ejecutando: {' '.join(cmd)}")
                        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        return True
                    except Exception as e:
                        print(f"Error al imprimir con PDFtoPrinter: {e}")
            
            # MÉTODO 4: Como alternativa final, usar ShellExecute con flags específicos
            # Este método puede mostrar un diálogo, pero es más confiable
            print(f"Usando ShellExecute para imprimir {file_path} en {printer_name}")
            win32api.ShellExecute(
                0, 
                "print", 
                file_path,
                f'/d:"{printer_name}"', 
                ".", 
                0
            )
            # Damos tiempo para que se procese la impresión
            time.sleep(3)
            return True
        
        elif file_ext in ['.txt']:
            try:
                # Convertir a PDF primero (más confiable)
                pdf_path = convert_to_pdf(file_path)
                if pdf_path != file_path:
                    return print_file_direct(pdf_path, printer_name)
                else:
                    # Imprimir texto directamente sólo si la conversión falla
                    hprinter = win32print.OpenPrinter(printer_name)
                    try:
                        hdc = win32ui.CreateDC()
                        hdc.CreatePrinterDC(printer_name)
                        
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.readlines()
                        
                        # Iniciar documento
                        hdc.StartDoc(os.path.basename(file_path))
                        hdc.StartPage()
                        
                        # Configurar fuente
                        font = win32ui.CreateFont({
                            'name': 'Courier New',
                            'height': 12,
                            'weight': 400,
                        })
                        hdc.SelectObject(font)
                        
                        # Imprimir líneas de texto
                        y = 100  # Posición inicial en Y
                        for line in content:
                            hdc.TextOut(100, y, line.strip())
                            y += 20  # Espacio entre líneas
                        
                        # Finalizar impresión
                        hdc.EndPage()
                        hdc.EndDoc()
                        hdc.DeleteDC()
                    finally:
                        win32print.ClosePrinter(hprinter)
                    return True
            except Exception as e:
                print(f"Error al imprimir archivo de texto: {e}")
                # Fallback a método alternativo
                win32api.ShellExecute(
                    0, 
                    "print", 
                    file_path,
                    f'/d:"{printer_name}"', 
                    ".", 
                    0
                )
                time.sleep(2)
                return True
        
        elif file_ext in ['.doc', '.docx', '.jpg', '.jpeg', '.png']:
            # Para estos formatos, usamos el método de conversión a PDF
            pdf_path = convert_to_pdf(file_path)
            if pdf_path != file_path:  # Si se convirtió exitosamente
                return print_file_direct(pdf_path, printer_name)
            else:
                # Si la conversión falló, usamos shellexecute directamente
                print(f"Usando ShellExecute para imprimir {file_path} en {printer_name}")
                win32api.ShellExecute(
                    0, 
                    "print", 
                    file_path,
                    f'/d:"{printer_name}"', 
                    ".", 
                    0
                )
                time.sleep(2)
                return True
        
        else:
            # Para otros tipos de archivos
            print(f"Usando ShellExecute para imprimir {file_path} en {printer_name}")
            win32api.ShellExecute(
                0, 
                "print", 
                file_path,
                f'/d:"{printer_name}"', 
                ".", 
                0
            )
            time.sleep(2)
            return True
    
    except Exception as e:
        print(f"Error al imprimir: {e}")
        return False

# Cargar historial de impresión
def load_print_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

# Guardar historial de impresión
def save_print_history(history):
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

# Agregar un trabajo al historial
def add_print_job(filename, status="Pendiente"):
    history = load_print_history()
    
    # Generar un ID nuevo
    job_id = 1
    if history:
        job_id = max(item["id"] for item in history) + 1
    
    # Fecha actual formateada
    now = datetime.datetime.now()
    formatted_date = now.strftime("%d/%m/%Y %H:%M:%S")
    
    # Crear registro
    job = {
        "id": job_id,
        "nombre": filename,
        "fecha": formatted_date,
        "estado": status
    }
    
    history.append(job)
    save_print_history(history)
    return job_id

# Actualizar estado de un trabajo
def update_print_job_status(job_id, status):
    history = load_print_history()
    
    for job in history:
        if job["id"] == job_id:
            job["estado"] = status
            break
    
    save_print_history(history)

@app.route('/')
def index():
    # Servir el archivo HTML estático
    return render_template('index.html')

@app.route('/static/<path:path>')
def serve_static(path):
    return send_from_directory('static', path)

@app.route('/imprimir', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No se seleccionó ningún archivo"}), 400
    
    file = request.files['file']
    printer_name = request.form.get('printer', None)  # Obtener impresora seleccionada
    
    if file.filename == '':
        return jsonify({"error": "No se seleccionó ningún archivo"}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Registrar en el historial
        job_id = add_print_job(filename)
        
        try:
            # Convertir a PDF si es necesario
            print_ready_file = convert_to_pdf(file_path)
            
            # Imprimir directamente sin diálogos
            success = print_file_direct(print_ready_file, printer_name)
            
            if success:
                update_print_job_status(job_id, "Impreso")
                return jsonify({"success": True, "message": f"Archivo '{filename}' enviado a imprimir correctamente"})
            else:
                update_print_job_status(job_id, "Error: No se pudo imprimir")
                return jsonify({"error": f"Error al imprimir '{filename}'"}), 500
        except Exception as e:
            update_print_job_status(job_id, f"Error: {str(e)[:50]}")
            return jsonify({"error": f"Error en el procesamiento: {str(e)}"}), 500
    
    return jsonify({"error": "Tipo de archivo no permitido"}), 400

@app.route('/impresoras', methods=['GET'])
def get_printers():
    """Devuelve la lista de impresoras disponibles"""
    try:
        printers = []
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
            printers.append({
                'name': printer[2],
                'default': printer[2] == win32print.GetDefaultPrinter()
            })
        return jsonify(printers)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/historial', methods=['GET'])
def get_history():
    history = load_print_history()
    return jsonify(history)

# Punto de entrada principal
if __name__ == '__main__':
    setup_directories()
    print("Servidor de impresión iniciado en http://0.0.0.0:5000")

from flask import Flask, render_template, request, redirect, url_for, jsonify, send_from_directory
import os
import sys
import subprocess
import platform
import tempfile
import win32print
import win32api
import win32ui
from win32con import LOGPIXELSX, LOGPIXELSY
from werkzeug.utils import secure_filename
import pythoncom
import time
import json
import datetime
import shutil
from pathlib import Path
from PIL import Image
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.secret_key = 'clave_secreta_para_flask'

# Archivo para almacenar el historial de impresión
HISTORY_FILE = 'print_history.json'

# Asegurar que existe el directorio de uploads
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Extensiones permitidas
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc', 'txt', 'jpg', 'jpeg', 'png'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
           
# Asegurar que existen los directorios necesarios
def setup_directories():
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs('static', exist_ok=True)
    
    # Si no existe el directorio de templates, crearlo
    if not os.path.exists('templates'):
        os.makedirs('templates')
        
    # Si no existe el archivo index.html en templates, crear uno básico
    if not os.path.exists(os.path.join('templates', 'index.html')):
        with open(os.path.join('templates', 'index.html'), 'w', encoding='utf-8') as f:
            f.write("""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="refresh" content="0; url=/static/index.html">
    <title>Redireccionando...</title>
</head>
<body>
    <p>Redireccionando a la interfaz...</p>
</body>
</html>""")
    
    # Si no existe el archivo de historial, crearlo vacío
    if not os.path.exists(HISTORY_FILE):
        save_print_history([])

def convert_to_pdf(input_file):
    """Convierte diferentes formatos a PDF si es necesario"""
    filename, file_extension = os.path.splitext(input_file)
    output_pdf = f"{filename}.pdf"
    
    # Si ya es PDF, devolvemos la misma ruta
    if file_extension.lower() == '.pdf':
        return input_file
    
    # Convertir DOCX a PDF con alternativas que no usan comtypes
    elif file_extension.lower() in ['.docx', '.doc']:
        try:
            # Intentamos usar LibreOffice si está disponible
            for libreoffice_path in [
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
            ]:
                if os.path.exists(libreoffice_path):
                    cmd = [
                        libreoffice_path,
                        '--headless',
                        '--convert-to', 'pdf',
                        '--outdir', os.path.dirname(input_file),
                        input_file
                    ]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    return output_pdf
            
            # Alternativa: intentar usar soffice en PATH
            if shutil.which('soffice'):
                cmd = [
                    'soffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(input_file),
                    input_file
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                return output_pdf
                
            # Si llegamos aquí, informamos que no podemos convertir
            print(f"No se pudo convertir el documento Word: LibreOffice no está instalado")
            return input_file
            
        except Exception as e:
            print(f"Error al convertir documento: {e}")
            return input_file
    
    # Para archivos de texto plano, convertimos a PDF
    elif file_extension.lower() == '.txt':
        try:
            # Crear un PDF desde el archivo de texto
            c = canvas.Canvas(output_pdf, pagesize=letter)
            text_file = open(input_file, 'r', encoding='utf-8', errors='ignore')
            y = 750  # posición inicial en Y
            for line in text_file:
                if y < 50:  # Si alcanzamos el final de la página
                    c.showPage()  # Nueva página
                    y = 750
                c.drawString(50, y, line.strip())
                y -= 12  # Espacio entre líneas
            text_file.close()
            c.save()
            return output_pdf
        except Exception as e:
            print(f"Error al convertir TXT a PDF: {e}")
            return input_file
    
    # Para imágenes, convertimos a PDF
    elif file_extension.lower() in ['.jpg', '.jpeg', '.png']:
        try:
            img = Image.open(input_file)
            img_rgb = img.convert('RGB')
            img_rgb.save(output_pdf, 'PDF')
            return output_pdf
        except Exception as e:
            print(f"Error al convertir imagen a PDF: {e}")
            return input_file
    
    return input_file

def print_file_direct(file_path, printer_name=None):
    """Imprime un archivo directamente sin mostrar diálogos"""
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()
    
    try:
        file_path = os.path.abspath(file_path)
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # Para PDF, usamos una impresión más directa
        if file_ext == '.pdf':
            # MÉTODO 1: Imprimir usando GhostScript (si está instalado)
            gs_paths = [
                shutil.which('gsprint.exe'),
                r'C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe',
                r'C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe', 
                r'C:\Program Files (x86)\gs\gs9.56.1\bin\gswin32c.exe'
            ]
            
            for gs_path in gs_paths:
                if gs_path and os.path.exists(gs_path):
                    try:
                        if 'gsprint' in gs_path:
                            cmd = [gs_path, '-printer', printer_name, file_path]
                            print(f"Ejecutando: {' '.join(cmd)}")
                            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            return True
                        else:
                            cmd = [
                                gs_path, 
                                '-dPrinted', '-dBATCH', '-dNOPAUSE', '-dNOPROMPT',
                                f'-sOutputFile=%printer%{printer_name}', 
                                '-sDEVICE=mswinpr2',
                                file_path
                            ]
                            print(f"Ejecutando: {' '.join(cmd)}")
                            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            return True
                    except Exception as e:
                        print(f"Error al imprimir usando GhostScript: {e}")
                        # Continuamos con otros métodos si este falla
            
            # MÉTODO 2: Usar SumatraPDF (si está disponible)
            sumatra_paths = [
                r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
                r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe'
            ]
            
            for sumatra_path in sumatra_paths:
                if os.path.exists(sumatra_path):
                    try:
                        cmd = [sumatra_path, '-print-to', printer_name, '-silent', file_path]
                        print(f"Ejecutando: {' '.join(cmd)}")
                        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        return True
                    except Exception as e:
                        print(f"Error al imprimir con SumatraPDF: {e}")
                        # Continuamos con otros métodos si este falla
            
            # MÉTODO 3: Usar PDFtoPrinter (si está disponible)
            pdftoprinterpaths = [
                r'C:\Program Files\PDFtoPrinter\PDFtoPrinter.exe',
                r'C:\Program Files (x86)\PDFtoPrinter\PDFtoPrinter.exe',
                'PDFtoPrinter.exe'  # Si está en PATH
            ]
            
            for pdftoprinter in pdftoprinterpaths:
                if os.path.exists(pdftoprinter) or shutil.which(pdftoprinter):
                    try:
                        cmd = [pdftoprinter, file_path, printer_name]
                        print(f"Ejecutando: {' '.join(cmd)}")
                        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        return True
                    except Exception as e:
                        print(f"Error al imprimir con PDFtoPrinter: {e}")
            
            # MÉTODO 4: Como alternativa final, usar ShellExecute con flags específicos
            # Este método puede mostrar un diálogo, pero es más confiable
            print(f"Usando ShellExecute para imprimir {file_path} en {printer_name}")
            win32api.ShellExecute(
                0, 
                "print", 
                file_path,
                f'/d:"{printer_name}"', 
                ".", 
                0
            )
            # Damos tiempo para que se procese la impresión
            time.sleep(3)
            return True
        
        elif file_ext in ['.txt']:
            try:
                # Convertir a PDF primero (más confiable)
                pdf_path = convert_to_pdf(file_path)
                if pdf_path != file_path:
                    return print_file_direct(pdf_path, printer_name)
                else:
                    # Imprimir texto directamente sólo si la conversión falla
                    hprinter = win32print.OpenPrinter(printer_name)
                    try:
                        hdc = win32ui.CreateDC()
                        hdc.CreatePrinterDC(printer_name)
                        
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.readlines()
                        
                        # Iniciar documento
                        hdc.StartDoc(os.path.basename(file_path))
                        hdc.StartPage()
                        
                        # Configurar fuente
                        font = win32ui.CreateFont({
                            'name': 'Courier New',
                            'height': 12,
                            'weight': 400,
                        })
                        hdc.SelectObject(font)
                        
                        # Imprimir líneas de texto
                        y = 100  # Posición inicial en Y
                        for line in content:
                            hdc.TextOut(100, y, line.strip())
                            y += 20  # Espacio entre líneas
                        
                        # Finalizar impresión
                        hdc.EndPage()
                        hdc.EndDoc()
                        hdc.DeleteDC()
                    finally:
                        win32print.ClosePrinter(hprinter)
                    return True
            except Exception as e:
                print(f"Error al imprimir archivo de texto: {e}")
                # Fallback a método alternativo
                win32api.ShellExecute(
                    0, 
                    "print", 
                    file_path,
                    f'/d:"{printer_name}"', 
                    ".", 
                    0
                )
                time.sleep(2)
                return True
        
        elif file_ext in ['.doc', '.docx', '.jpg', '.jpeg', '.png']:
            # Para estos formatos, usamos el método de conversión a PDF
            pdf_path = convert_to_pdf(file_path)
            if pdf_path != file_path:  # Si se convirtió exitosamente
                return print_file_direct(pdf_path, printer_name)
            else:
                # Si la conversión falló, usamos shellexecute directamente
                print(f"Usando ShellExecute para imprimir {file_path} en {printer_name}")
                win32api.ShellExecute(
                    0, 
                    "print", 
                    file_path,
                    f'/d:"{printer_name}"', 
                    ".", 
                    0
                )
                time.sleep(2)
                return True
        
        else:
            # Para otros tipos de archivos
            print(f"Usando ShellExecute para imprimir {file_path} en {printer_name}")
            win32api.ShellExecute(
                0, 
                "print", 
                file_path,
                f'/d:"{printer_name}"', 
                ".", 
                0
            )
            time.sleep(2)
            return True
    
    except Exception as e:
        print(f"Error al imprimir: {e}")
        return False

# Cargar historial de impresión
def load_print_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

# Guardar historial de impresión
def save_print_history(history):
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

# Agregar un trabajo al historial
def add_print_job(filename, status="Pendiente"):
    history = load_print_history()
    
    # Generar un ID nuevo
    job_id = 1
    if history:
        job_id = max(item["id"] for item in history) + 1
    
    # Fecha actual formateada
    now = datetime.datetime.now()
    formatted_date = now.strftime("%d/%m/%Y %H:%M:%S")
    
    # Crear registro
    job = {
        "id": job_id,
        "nombre": filename,
        "fecha": formatted_date,
        "estado": status
    }
    
    history.append(job)
    save_print_history(history)
    return job_id

# Actualizar estado de un trabajo
def update_print_job_status(job_id, status):
    history = load_print_history()
    
    for job in history:
        if job["id"] == job_id:
            job["estado"] = status
            break
    
    save_print_history(history)

@app.route('/')
def index():
    # Servir el archivo HTML estático
    return render_template('index.html')

@app.route('/static/<path:path>')
def serve_static(path):
    return send_from_directory('static', path)

@app.route('/imprimir', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No se seleccionó ningún archivo"}), 400
    
    file = request.files['file']
    printer_name = request.form.get('printer', None)  # Obtener impresora seleccionada
    
    if file.filename == '':
        return jsonify({"error": "No se seleccionó ningún archivo"}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Registrar en el historial
        job_id = add_print_job(filename)
        
        try:
            # Convertir a PDF si es necesario
            print_ready_file = convert_to_pdf(file_path)
            
            # Imprimir directamente sin diálogos
            success = print_file_direct(print_ready_file, printer_name)
            
            if success:
                update_print_job_status(job_id, "Impreso")
                return jsonify({"success": True, "message": f"Archivo '{filename}' enviado a imprimir correctamente"})
            else:
                update_print_job_status(job_id, "Error: No se pudo imprimir")
                return jsonify({"error": f"Error al imprimir '{filename}'"}), 500
        except Exception as e:
            update_print_job_status(job_id, f"Error: {str(e)[:50]}")
            return jsonify({"error": f"Error en el procesamiento: {str(e)}"}), 500
    
    return jsonify({"error": "Tipo de archivo no permitido"}), 400

@app.route('/impresoras', methods=['GET'])
def get_printers():
    """Devuelve la lista de impresoras disponibles"""
    try:
        printers = []
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
            printers.append({
                'name': printer[2],
                'default': printer[2] == win32print.GetDefaultPrinter()
            })
        return jsonify(printers)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/historial', methods=['GET'])
def get_history():
    history = load_print_history()
    return jsonify(history)

# Punto de entrada principal
if __name__ == '__main__':
    setup_directories()
    print("Servidor de impresión iniciado en http://0.0.0.0:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)