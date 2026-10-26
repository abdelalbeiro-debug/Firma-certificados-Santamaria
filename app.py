import io
import os
import zipfile
import tempfile
import subprocess
import fitz  # PyMuPDF para manipular el PDF e inyectar la firma
import qrcode
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

st.set_page_config(page_title="Firma Masiva de Auditoría - Formato Intacto", layout="wide")

st.title("🔏 Firma Masiva de Certificados (Formato 100% Intacto)")
st.markdown("""
Esta versión preserva exactamente las tablas, alineaciones y tipografías originales del documento Word 
e inyecta el sello digital con código QR directamente en el espacio de la firma.
""")

# --- Configuración del Contador ---
st.sidebar.header("📋 Datos del Contador")
nombre_contador = st.sidebar.text_input("Nombre del Contador", "Abdel Areiza")
tarjeta_profesional = st.sidebar.text_input("Tarjeta Profesional", "TP-123456-T")
entidad_auditoria = st.sidebar.text_input("Entidad / Auditoría", "Auditoría Financiera 24 Años")

uploaded_zip = st.file_uploader("Sube el archivo ZIP con las carpetas de los terceros", type=["zip"])

def crear_imagen_sello(nombre_tercero, contador, tp, entidad):
    """Crea una imagen de sello limpia, con mejor resolución y texto espaciado."""
    # 1. Generar Código QR con mayor resolución
    texto_qr = f"Tercero: {nombre_tercero} | Contador: {contador} | TP: {tp} | {entidad}"
    qr = qrcode.QRCode(version=1, box_size=4, border=1)
    qr.add_data(texto_qr)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
    
    # 2. Lienzo más grande para dar espacio y nitidez
    ancho, alto = 500, 110
    sello_bg = Image.new('RGB', (ancho, alto), color='#F8FAFC')
    draw = ImageDraw.Draw(sello_bg)
    
    # Borde exterior elegante
    draw.rectangle([0, 0, ancho - 1, alto - 1], outline='#1A365D', width=2)
    
    # Pegar QR escalado y proporcionado a la izquierda
    qr_size = 90
    qr_img = qr_img.resize((qr_size, qr_size), Image.Resampling.LANCZOS)
    sello_bg.paste(qr_img, (10, 10))
    
    # Intentar cargar una fuente del sistema o usar la predeterminada con espaciado correcto
    try:
        # Intenta usar Arial si está disponible en Windows
        font_titulo = ImageFont.truetype("arial.ttf", 13)
        font_texto = ImageFont.truetype("arial.ttf", 11)
    except:
        font_titulo = ImageFont.load_default()
        font_texto = ImageFont.load_default()

    # Textos organizados
    hash_val = abs(hash(texto_qr))
    
    x_texto = 115
    draw.text((x_texto, 12), "CERTIFICADO DIGITAL DE AUDITORÍA", fill='#1A365D', font=font_titulo)
    draw.text((x_texto, 32), f"Tercero / Cliente: {nombre_tercero}", fill='#333333', font=font_texto)
    draw.text((x_texto, 50), f"Contador Firmante: {contador}  |  T.P. N°: {tp}", fill='#333333', font=font_texto)
    draw.text((x_texto, 68), f"Entidad: {entidad}", fill='#333333', font=font_texto)
    draw.text((x_texto, 86), f"Verificación Hash: {hash_val}", fill='#666666', font=font_texto)
    
    sello_path = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
    sello_bg.save(sello_path, quality=95)
    return sello_path

def convertir_docx_a_pdf_nativo(docx_path, output_dir):
    """Convierte Word a PDF usando LibreOffice o Word COM para mantener formato 100% exacto."""
    # Intentar conversión con LibreOffice (ideal para Streamlit Cloud / Linux)
    try:
        subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf',
            '--outdir', output_dir, docx_path
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        pdf_nombre = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        return os.path.join(output_dir, pdf_nombre)
    except Exception as e:
        # Si falla (ej. Windows local sin soffice), intentar con docx2pdf
        try:
            from docx2pdf import convert
            pdf_nombre = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            pdf_path = os.path.join(output_dir, pdf_nombre)
            convert(docx_path, pdf_path)
            return pdf_path
        except Exception as ex:
            raise RuntimeError(f"No se pudo convertir el Word a PDF manteniendo el formato: {ex}")

def estampar_sello_en_pdf(pdf_path, sello_img_path, output_pdf_path):
    """Estampa la imagen del sello en la parte inferior de TODAS las páginas del PDF."""
    doc = fitz.open(pdf_path)
    
    ancho_sello = 320
    alto_sello = 70
    
    # Recorrer todas las páginas del documento
    for pagina in doc:
        rect_pagina = pagina.rect
        x1 = (rect_pagina.width - ancho_sello) / 2
        y1 = rect_pagina.height - alto_sello - 40  # 40 puntos desde el borde inferior
        x2 = x1 + ancho_sello
        y2 = y1 + alto_sello
        
        rect_sello = fitz.Rect(x1, y1, x2, y2)
        pagina.insert_image(rect_sello, filename=sello_img_path)
    
    doc.save(output_pdf_path)
    doc.close()

if uploaded_zip is not None:
    st.info("Archivo ZIP cargado correctamente. Presiona el botón para procesar con formato nativo.")
    
    if st.button("🚀 Procesar y Estampar Firma Preservando Formato"):
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, "input.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.getbuffer())
            
            extract_dir = os.path.join(temp_dir, "extracted")
            os.makedirs(extract_dir, exist_ok=True)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            # Buscar archivos Word
            docx_files = []
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.lower().endswith(".docx") and not file.startswith("~$"):
                        docx_files.append((root, file))
            
            total = len(docx_files)
            if total == 0:
                st.warning("No se encontraron archivos Word (.docx) dentro del ZIP.")
            else:
                barra = st.progress(0)
                procesados = 0
                
                for i, (root, file) in enumerate(docx_files):
                    docx_path = os.path.join(root, file)
                    nombre_tercero = os.path.basename(root)
                    
                    try:
                        # 1. Convertir Word a PDF idéntico
                        pdf_temp = convertir_docx_a_pdf_nativo(docx_path, root)
                        
                        # 2. Generar imagen del sello
                        sello_img = crear_imagen_sello(
                            nombre_tercero, nombre_contador, tarjeta_profesional, entidad_auditoria
                        )
                        
                        # 3. Estampar sello en el PDF
                        pdf_final_path = os.path.join(root, f"{os.path.splitext(file)[0]}_FIRMADO.pdf")
                        estampar_sello_en_pdf(pdf_temp, sello_img, pdf_final_path)
                        
                        # Limpiar PDF intermedio
                        if os.path.exists(pdf_temp) and pdf_temp != pdf_final_path:
                            os.remove(pdf_temp)
                        if os.path.exists(sello_img):
                            os.remove(sello_img)
                            
                        procesados += 1
                    except Exception as e:
                        st.error(f"Error procesando {file}: {e}")
                    
                    barra.progress((i + 1) / total)
                
                # Crear ZIP final
                output_zip_path = os.path.join(temp_dir, "certificados_firmados_formato_ok.zip")
                with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for root, dirs, files in os.walk(extract_dir):
                        for file in files:
                            if file.endswith("_FIRMADO.pdf"):
                                full_path = os.path.join(root, file)
                                rel_path = os.path.relpath(full_path, extract_dir)
                                zip_out.write(full_path, rel_path)
                
                with open(output_zip_path, "rb") as f:
                    zip_bytes = f.read()
                    
                st.success(f"¡Proceso completado! Se firmaron {procesados} certificados conservando el formato exacto del Word.")
                
                st.download_button(
                    label="📦 Descargar ZIP con PDF firmados (Formato Intacto)",
                    data=zip_bytes,
                    file_name="certificados_firmados_formato_intacto.zip",
                    mime="application/zip"
                )