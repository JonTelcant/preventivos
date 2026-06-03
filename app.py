from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
import openpyxl
import os
from datetime import datetime
from werkzeug.utils import secure_filename
from difflib import SequenceMatcher as IndiceCoincidencia
import re
import csv
import google.auth
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

app = Flask(__name__)
app.secret_key = 'tu_clave_secreta_aqui'
app.config['SESSION_TYPE'] = 'filesystem'

# Configuración
UPLOAD_FOLDER = 'uploads'
MAPA_FILE = 'mapa_preventivos.xlsx'
ALLOWED_EXTENSIONS = {'xlsx'}
DOWNLOAD_FOLDER = 'downloads'
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# Google Drive API scopes
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'

def obtener_servicio_drive():
    """Obtiene el servicio de Google Drive autenticado usando OAuth con token pre-generado."""
    try:
        creds = None
        if os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        else:
            print(f"Error: No se encontró el archivo {TOKEN_FILE}")
            return None
        
        if creds and creds.valid:
            return build('drive', 'v3', credentials=creds)
        
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                with open(TOKEN_FILE, 'w') as token:
                    token.write(creds.to_json())
                return build('drive', 'v3', credentials=creds)
            except Exception as e:
                print(f"Error al renovar token: {str(e)}")
                return None
        
        print("Error: Token inválido o expirado sin refresh token")
        return None
        
    except Exception as e:
        print(f"Error al autenticar con Google Drive: {str(e)}")
        return None

def subir_archivo_drive(ruta_archivo, nombre_archivo, carpeta_destino):
    """Sube un archivo a Google Drive en la carpeta especificada (puede ser ruta anidada)."""
    try:
        service = obtener_servicio_drive()
        if not service:
            return False, "No se pudo autenticar con Google Drive"
        
        # Dividir la ruta en carpetas anidadas
        carpetas = carpeta_destino.split('/')
        parent_id = None
        
        for carpeta in carpetas:
            # Buscar o crear la carpeta
            folder_id = None
            query = f"name='{carpeta}' and mimeType='application/vnd.google-apps.folder'"
            if parent_id:
                query += f" and '{parent_id}' in parents"
            
            results = service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)'
            ).execute()
            
            folders = results.get('files', [])
            if folders:
                folder_id = folders[0]['id']
            else:
                # Crear la carpeta
                folder_metadata = {
                    'name': carpeta,
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                if parent_id:
                    folder_metadata['parents'] = [parent_id]
                folder = service.files().create(body=folder_metadata, fields='id').execute()
                folder_id = folder.get('id')
            
            parent_id = folder_id
        
        # Subir el archivo
        file_metadata = {
            'name': nombre_archivo,
            'parents': [parent_id]
        }
        media = MediaFileUpload(ruta_archivo, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        
        return True, f"Archivo subido correctamente con ID: {file.get('id')}"
    except Exception as e:
        return False, f"Error al subir a Google Drive: {str(e)}"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Funciones para comparación de preventivos y emplazamientos
def formatear_emplazamiento(emplazamiento):
    """Formatea el nombre del emplazamiento eliminando patrones comunes."""
    emplazamiento_formateado = re.sub('(2G)', '', emplazamiento)
    emplazamiento_formateado = re.sub('(3G)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(4G)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(C.T.)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(CT)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(REP-INT-GSM)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(--)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(ATW-T)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(ATW-V)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(ATW)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('[-]', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(TSM)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(T.S.M.)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('(E.B.)', '', emplazamiento_formateado)
    emplazamiento_formateado = re.sub('\s?\d', '', emplazamiento_formateado)
    emplazamiento_formateado = emplazamiento_formateado.lstrip()
    return emplazamiento_formateado

def limpiar_title(title):
    """Limpia el title para usar en URL: quita /, dobles espacios y codifica espacios."""
    # Quitar barras
    title = title.replace('/', '')
    # Quitar dobles espacios
    title = re.sub(r'\s+', ' ', title)
    # Codificar espacios como %20
    title = title.replace(' ', '%20')
    return title

def accion_a_tipos(accion):
    """Convierte el texto de acción a tipos abreviados de preventivos."""
    tipos = set()
    accion_upper = accion.upper()
    
    # Mapeo exacto de nombres de preventivos a tipos
    mapeo_tipos = {
        'MP EB BT ANUAL': 'BT',
        'MP EB OC ANUAL': 'OC',
        'MP EB ANTENA ANUAL': 'SA',
        'MP EB ALIM EQ RADIO': 'AE',
        'MP CABLE DE VIDA ANUAL': 'GS',
        'MP EB CF ANUAL': 'CF',
        'MP EB AA ANUAL': 'AA',
        'MP EB ARMARIO INTEMPERIE': 'AI'
    }
    
    # Buscar coincidencias exactas en el texto de acción
    for nombre, tipo in mapeo_tipos.items():
        if nombre in accion_upper:
            tipos.add(tipo)
    
    # Si no se encontraron tipos, devolver lista vacía
    return sorted(list(tipos))

def leer_excel_emplazamientos(archivo):
    """Lee el archivo de emplazamientos y retorna una lista."""
    hoja_excel = archivo.worksheets[0]
    lista = []
    for fila in hoja_excel.rows:
        emplazamiento = fila[4].value
        if emplazamiento:
            emplazamiento_formateado = formatear_emplazamiento(str(emplazamiento))
            latitud = fila[2].value
            altitud = fila[3].value
            lista.append([emplazamiento, emplazamiento_formateado, latitud, altitud])
    return lista

def leer_excel_preventivos(archivo):
    """Lee el archivo de preventivos y retorna una lista."""
    hoja_excel = archivo.worksheets[0]
    lista = []
    cont = 0
    emplazamiento_a = "SALA ELEMENTO"
    accion = ""
    for fila in hoja_excel.rows:
        cont += 1
        if emplazamiento_a != fila[5].value:
            if cont != 2:
                lista.append([emplazamiento_a, emplazamiento_formateado_a, emplazamiento_b,
                             emplazamiento_formateado_b, accion])
            emplazamiento_a = fila[5].value
            emplazamiento_b = fila[6].value
            emplazamiento_formateado_a = formatear_emplazamiento(str(emplazamiento_a)) if emplazamiento_a else ""
            emplazamiento_formateado_b = formatear_emplazamiento(str(emplazamiento_b)) if emplazamiento_b else ""
            accion = str(fila[8].value) if fila[8].value else ""
        else:
            if cont != 1:
                accion = accion + "\n" + str(fila[8].value)
    lista.append([emplazamiento_a, emplazamiento_formateado_a, emplazamiento_b,
                 emplazamiento_formateado_b, accion])
    return lista

def buscar_coincidencia(lista_preventivos, lista_emplazamientos, indice):
    """Busca coincidencias entre preventivos y emplazamientos."""
    lista_coincidencia = []
    lista_sin_coincidencia = []
    lista_multiple = []  # Para guardar preventivos con múltiples coincidencias
    for preventivo in lista_preventivos:
        contador_coincidencias = 0
        lista = []
        nombre_preventivo = preventivo[0]
        accion = preventivo[4]
        emplazamiento_coincidencia = []
        for emplazamiento in lista_emplazamientos:
            if IndiceCoincidencia(None, preventivo[indice], emplazamiento[1]).ratio() > 0.7:
                latitud = emplazamiento[2]
                altitud = emplazamiento[3]
                # Limpiar title y obtener tipos
                title_limpio = limpiar_title(nombre_preventivo)
                tipos = accion_a_tipos(accion)
                # Generar URL con parámetros de tipos individuales
                tipos_params = '&'.join([f"{tipo}=1" for tipo in tipos])
                url = f"https://preventivos-rgkk.onrender.com/rellenar?title={title_limpio}"
                if tipos_params:
                    url += f"&{tipos_params}"
                lista.append(["Preventivos", "#FF0000", latitud, altitud, nombre_preventivo, url, "#71b300", emplazamiento[0]])
                contador_coincidencias += 1
                emplazamiento_coincidencia.append(emplazamiento)
        if contador_coincidencias == 0:
            lista_sin_coincidencia.append(preventivo)
        elif contador_coincidencias == 1:
            lista_coincidencia.append(lista[0])
        else:
            # Guardar preventivo con múltiples coincidencias para revisión
            lista_multiple.append({
                'preventivo': preventivo,
                'coincidencias': lista,
                'indice': indice
            })
    lista_resultado = []
    lista_resultado.append(lista_coincidencia)
    lista_resultado.append(lista_sin_coincidencia)
    lista_resultado.append(lista_multiple)
    return lista_resultado

def coincidencias(lista_preventivos, lista_emplazamientos):
    """Realiza la búsqueda de coincidencias en dos vueltas."""
    lista_preventivos = buscar_coincidencia(lista_preventivos, lista_emplazamientos, 1)
    lista_coincidencia = lista_preventivos[0]
    lista_sin_coincidencia = lista_preventivos[1]
    lista_multiple_vuelta1 = lista_preventivos[2]
    
    lista_preventivos = buscar_coincidencia(lista_sin_coincidencia, lista_emplazamientos, 3)
    
    for item in lista_preventivos[0]:
        lista_coincidencia.append(item)
    for item in lista_preventivos[1]:
        # Generar URL para preventivos sin coincidencia
        nombre_preventivo = item[0]
        accion = item[4]
        # Limpiar title y obtener tipos
        title_limpio = limpiar_title(nombre_preventivo)
        tipos = accion_a_tipos(accion)
        # Generar URL con parámetros de tipos individuales
        tipos_params = '&'.join([f"{tipo}=1" for tipo in tipos])
        url = f"https://preventivos-rgkk.onrender.com/rellenar?title={title_limpio}"
        if tipos_params:
            url += f"&{tipos_params}"
        lista_coincidencia.append(["Preventivos", "#FF0000", 0, 0, nombre_preventivo, url, "#FF0000"])
    
    # Combinar las múltiples coincidencias de ambas vueltas
    lista_multiple = lista_multiple_vuelta1 + lista_preventivos[2]
    
    return lista_coincidencia, lista_multiple

def obtener_tipo_preventivo(nombre_archivo):
    """Determina el tipo de preventivo basándose en el nombre del archivo."""
    nombre_upper = nombre_archivo.upper()
    
    if 'ALIMENTACIÓN' in nombre_upper or 'ALIMENTACION' in nombre_upper:
        return 'AE'
    elif 'ARMARIOS' in nombre_upper and 'INTEMPERIE' in nombre_upper:
        return 'AI'
    elif '_EBAA_' in nombre_upper or 'EBAA' in nombre_upper:
        return 'AA'
    elif '_SA_' in nombre_upper:
        return 'SA'
    elif '_GS' in nombre_upper:
        return 'GS'
    elif '_OC' in nombre_upper:
        return 'OC'
    elif 'BT' in nombre_upper:
        return 'BT'
    elif '_CF' in nombre_upper:
        return 'CF'
    else:
        # Fallback
        if 'SA' in nombre_upper:
            return 'SA'
        elif 'GS' in nombre_upper:
            return 'GS'
        elif 'OC' in nombre_upper:
            return 'OC'
        elif 'BT' in nombre_upper:
            return 'BT'
        elif 'CF' in nombre_upper:
            return 'CF'
        return 'DESCONOCIDO'

def validar_archivo_preventivo(ruta_archivo):
    """
    Valida si un archivo es un preventivo válido.
    Un preventivo válido debe tener:
    - Una hoja llamada 'Maestra'
    - La hoja Maestra debe estar protegida
    - Tener celdas combinadas
    """
    try:
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        
        # Verificar que tiene hoja Maestra
        if 'Maestra' not in wb.sheetnames:
            return False, "El archivo no contiene una hoja llamada 'Maestra'"
        
        ws_maestra = wb['Maestra']
        
        # Verificar que la hoja Maestra está protegida
        if not ws_maestra.protection.sheet:
            return False, "La hoja 'Maestra' no está protegida"
        
        # Verificar que tiene celdas combinadas
        if len(ws_maestra.merged_cells.ranges) == 0:
            return False, "La hoja 'Maestra' no tiene celdas combinadas"
        
        wb.close()
        return True, "Archivo válido"
        
    except Exception as e:
        return False, f"Error al leer el archivo: {str(e)}"

def buscar_celda_por_texto(ws, texto_buscar):
    """
    Busca una celda que contenga el texto especificado.
    Retorna la información de la celda encontrada o None.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and texto_buscar in cell.value:
                # Obtener la celda a la derecha
                col_derecha = cell.column + 1
                celda_derecha = ws.cell(row=cell.row, column=col_derecha)
                
                return {
                    'texto': texto_buscar,
                    'coord': cell.coordinate,
                    'coord_derecha': celda_derecha.coordinate,
                    'valor_derecha': str(celda_derecha.value) if celda_derecha.value else ''
                }
    return None

def analizar_archivo_excel(ruta_archivo):
    """Analiza un archivo Excel y extrae las celdas editables con sus preguntas."""
    nombre_archivo = os.path.basename(ruta_archivo)
    tipo_preventivo = obtener_tipo_preventivo(nombre_archivo)
    
    celdas_info = []
    info_especial = {
        'elemento': None,
        'hoja_validada': []
    }
    
    # Preguntas especiales que siempre son defecto
    preguntas_especiales = ['DNI/MATRICULA:', 'EMPRESA REALIZACIÓN:', 'TELÉFONO:']
    
    try:
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        
        # Determinar qué hojas analizar
        hojas_a_analizar = ['Maestra']
        
        # Si existe Hija1, agregarla
        if 'Hija1' in wb.sheetnames:
            hojas_a_analizar.append('Hija1')
        
        # Buscar ELEMENTO: siempre en Maestra
        ws_maestra = wb['Maestra']
        info_especial['elemento'] = buscar_celda_por_texto(ws_maestra, 'ELEMENTO:')
        
        # Buscar HOJA VALIDADA en Maestra y Hija1
        for nombre_hoja in hojas_a_analizar:
            ws = wb[nombre_hoja]
            hoja_validada_info = buscar_celda_por_texto(ws, 'HOJA VALIDADA')
            if hoja_validada_info:
                hoja_validada_info['hoja'] = nombre_hoja
                info_especial['hoja_validada'].append(hoja_validada_info)
        
        # Analizar celdas editables en las hojas correspondientes
        for nombre_hoja in hojas_a_analizar:
            ws = wb[nombre_hoja]
            
            for row in ws.iter_rows():
                for cell in row:
                    # Verificar si la celda no está bloqueada (editable)
                    if not cell.protection.locked:
                        # Evitar celdas combinadas que no son la esquina superior izquierda
                        es_merge = type(cell).__name__ == 'MergedCell'
                        if es_merge:
                            continue
                        
                        # Verificar si pertenece a un rango de celdas combinadas
                        rango_combinado = None
                        esquina_1_1_coord = None
                        for merged_range in ws.merged_cells.ranges:
                            if cell.coordinate in merged_range:
                                rango_combinado = str(merged_range)
                                min_col, min_row, max_col, max_row = merged_range.bounds
                                esquina_1_1_coord = ws.cell(row=min_row, column=min_col).coordinate
                                break
                        
                        # Determinar la coordenada de la celda de respuesta
                        coord_respuesta = esquina_1_1_coord if esquina_1_1_coord else cell.coordinate
                        
                        # Obtener la pregunta a la izquierda
                        col_respuesta = openpyxl.utils.column_index_from_string(coord_respuesta[0])
                        fila_respuesta = int(coord_respuesta[1:])
                        
                        pregunta_texto = ""
                        pregunta_coord = ""
                        
                        if col_respuesta > 1:
                            celda_izquierda = ws.cell(row=fila_respuesta, column=col_respuesta-1)
                            pregunta_coord = celda_izquierda.coordinate
                            
                            # Verificar si la celda izquierda está en un rango combinado
                            pregunta_esquina_1_1 = None
                            for preg_merged in ws.merged_cells.ranges:
                                if celda_izquierda.coordinate in preg_merged:
                                    preg_min_col, preg_min_row, _, _ = preg_merged.bounds
                                    pregunta_esquina_1_1 = ws.cell(row=preg_min_row, column=preg_min_col)
                                    pregunta_coord = pregunta_esquina_1_1.coordinate
                                    pregunta_texto = str(pregunta_esquina_1_1.value) if pregunta_esquina_1_1.value else ""
                                    break
                            
                            if not pregunta_esquina_1_1:
                                pregunta_texto = str(celda_izquierda.value) if celda_izquierda.value else ""
                        
                        # Verificar si tiene validación de datos
                        es_lista = False
                        lista_valores = ""
                        for dv in ws.data_validations.dataValidation:
                            if cell.coordinate in dv.sqref:
                                if getattr(dv, 'type', '') == 'list':
                                    es_lista = True
                                    formula1 = getattr(dv, 'formula1', '')
                                    if formula1 and formula1 != 'N/A':
                                        # Extraer valores de la fórmula (ej: "C,IG,IR,NP")
                                        lista_valores = formula1.strip('"')
                                break
                        
                        # Determinar si es una pregunta especial
                        config_default = None
                        valor_default = ''
                        for preg_especial in preguntas_especiales:
                            if preg_especial in pregunta_texto:
                                config_default = 'defecto'
                                valor_default = ''  # Se llenará con valores globales
                                break
                        
                        celdas_info.append({
                            'pregunta': pregunta_texto,
                            'coord_celda': cell.coordinate,
                            'coord_respuesta': coord_respuesta,
                            'es_lista': es_lista,
                            'lista_valores': lista_valores,
                            'hoja': nombre_hoja,
                            'config_default': config_default,
                            'valor_default': valor_default
                        })
        
        wb.close()
        
    except Exception as e:
        return None, str(e), None
    
    return celdas_info, None, info_especial

def asegurar_mapa_preventivos(tipo_preventivo):
    """
    Asegura que existe el archivo mapa_preventivos.xlsx y la hoja del tipo de preventivo.
    Si no existen, los crea.
    """
    try:
        # Verificar si existe el archivo
        if not os.path.exists(MAPA_FILE):
            # Crear archivo nuevo
            wb = openpyxl.Workbook()
            # Eliminar hoja por defecto
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
        else:
            # Abrir archivo existente
            wb = openpyxl.load_workbook(MAPA_FILE)
        
        # Verificar si existe la hoja del tipo de preventivo
        if tipo_preventivo not in wb.sheetnames:
            ws = wb.create_sheet(title=tipo_preventivo)
            # Crear encabezados (sin columna Es Lista, con Hoja)
            headers = ['Pregunta', 'Posición Celda', 'Posición Respuesta', 'Hoja', 'Configuración', 'Valor']
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=header)
            
            # Ajustar ancho de columnas
            ws.column_dimensions['A'].width = 50
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 10
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 30
        
        wb.save(MAPA_FILE)
        wb.close()
        return True, None
    except PermissionError:
        return False, "El archivo mapa_preventivos.xlsx está abierto en otro programa. Por favor, ciérralo antes de continuar."
    except Exception as e:
        return False, f"Error al gestionar el archivo mapa_preventivos.xlsx: {str(e)}"

def guardar_configuracion_mapa(tipo_preventivo, configuraciones, info_especial):
    """
    Guarda las configuraciones en el archivo mapa_preventivos.xlsx.
    configuraciones es una lista de diccionarios con:
    - pregunta
    - coord_celda
    - coord_respuesta
    - hoja
    - config (ignorar, defecto, equipo, lista)
    - valor
    info_especial contiene información de ELEMENTO y HOJA VALIDADA
    """
    try:
        wb = openpyxl.load_workbook(MAPA_FILE)
        ws = wb[tipo_preventivo]
        
        # Limpiar datos existentes (mantener encabezados)
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col, value=None)
        
        fila = 2
        
        # Primero agregar información especial (ELEMENTO y HOJA VALIDADA)
        if info_especial and info_especial.get('elemento'):
            elem = info_especial['elemento']
            ws.cell(row=fila, column=1, value=elem['texto'])
            ws.cell(row=fila, column=2, value=elem['coord'])
            ws.cell(row=fila, column=3, value=elem['coord_derecha'])
            ws.cell(row=fila, column=4, value='Maestra')
            ws.cell(row=fila, column=5, value='INFO')
            ws.cell(row=fila, column=6, value='')  # No guardar valor
            fila += 1
        
        if info_especial and info_especial.get('hoja_validada'):
            for hv in info_especial['hoja_validada']:
                ws.cell(row=fila, column=1, value=hv['texto'])
                ws.cell(row=fila, column=2, value=hv['coord'])
                ws.cell(row=fila, column=3, value=hv['coord_derecha'])
                ws.cell(row=fila, column=4, value=hv['hoja'])
                ws.cell(row=fila, column=5, value='INFO')
                ws.cell(row=fila, column=6, value='')  # No guardar valor
                fila += 1
        
        # Escribir configuraciones de preguntas
        for config in configuraciones:
            ws.cell(row=fila, column=1, value=config['pregunta'])
            ws.cell(row=fila, column=2, value=config['coord_celda'])
            ws.cell(row=fila, column=3, value=config['coord_respuesta'])
            ws.cell(row=fila, column=4, value=config['hoja'])
            ws.cell(row=fila, column=5, value=config['config'])
            ws.cell(row=fila, column=6, value=config.get('valor', ''))
            fila += 1
        
        wb.save(MAPA_FILE)
        wb.close()
        return True, None
    except PermissionError:
        return False, "El archivo mapa_preventivos.xlsx está abierto en otro programa. Por favor, ciérralo antes de continuar."
    except Exception as e:
        return False, f"Error al guardar en el archivo mapa_preventivos.xlsx: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/rellenar')
def rellenar():
    """Ruta para rellenar preventivos basándose en parámetros de URL."""
    title = request.args.get('title', '')
    
    # Obtener tipos de preventivos de los parámetros de URL
    tipos = []
    for key in request.args.keys():
        if key.upper() in ['BT', 'OC', 'SA', 'AE', 'GS', 'CF', 'AA', 'AI']:
            if request.args.get(key) == '1':
                tipos.append(key.upper())
    
    # Leer el archivo mapa_preventivos.xlsx
    try:
        wb = openpyxl.load_workbook(MAPA_FILE)
    except Exception as e:
        flash(f'Error al cargar el archivo mapa_preventivos.xlsx: {str(e)}')
        return redirect(url_for('index'))
    
    # Para cada tipo, buscar la hoja correspondiente y leer las preguntas
    items_por_tipo = {}  # Para mantener el orden original
    equipos_por_tipo = {}  # Para guardar información de equipos por tipo
    
    for tipo in tipos:
        if tipo in wb.sheetnames:
            hoja = wb[tipo]
            items = []  # Para mantener el orden original
            equipos_info = {}  # Para agrupar preguntas por equipo
            equipos_procesados = set()  # Para evitar duplicados
            
            # Leer filas desde la fila 2 (asumiendo que la fila 1 es el encabezado)
            for fila in hoja.iter_rows(min_row=2, values_only=True):
                pregunta = fila[0]  # Columna A
                if pregunta:  # Si hay pregunta
                    config = fila[4] if len(fila) > 4 else None  # Columna E
                    valor = fila[5] if len(fila) > 5 else None  # Columna F
                    
                    # Detectar si es una configuración de equipo
                    es_equipo = False
                    nombre_equipo = None
                    if config and config.lower().startswith('equipo'):
                        es_equipo = True
                        nombre_equipo = config
                    
                    # Solo incluir preguntas con configuración 'lista', 'rellenar' o 'equipoXX'
                    if config and (config.lower() in ['lista', 'rellenar'] or es_equipo):
                        # Si la configuración es 'lista', leer las siguientes columnas hasta encontrar una vacía
                        lista_valores = []
                        if config and config.lower() == 'lista':
                            col_index = 5  # Empezar desde la columna F
                            while col_index < len(fila) and fila[col_index]:
                                lista_valores.append(fila[col_index])
                                col_index += 1
                        
                        pregunta_info = {
                            'pregunta': pregunta,
                            'config': config,
                            'valor': valor,
                            'lista_valores': lista_valores,
                            'es_equipo': es_equipo,
                            'nombre_equipo': nombre_equipo
                        }
                        
                        # Si es equipo, agrupar por nombre de equipo
                        if es_equipo:
                            if nombre_equipo not in equipos_info:
                                equipos_info[nombre_equipo] = []
                            equipos_info[nombre_equipo].append(pregunta_info)
                            
                            # Si es la primera vez que encontramos este equipo, agregar el desplegable
                            if nombre_equipo not in equipos_procesados:
                                items.append({
                                    'tipo': 'equipo_selector',
                                    'nombre_equipo': nombre_equipo
                                })
                                equipos_procesados.add(nombre_equipo)
                            
                            # Agregar la pregunta
                            items.append({
                                'tipo': 'pregunta',
                                'data': pregunta_info
                            })
                        else:
                            items.append({
                                'tipo': 'pregunta',
                                'data': pregunta_info
                            })
            
            # Para cada equipo, leer la hoja correspondiente y obtener los equipos disponibles
            for nombre_equipo, preguntas_equipo in equipos_info.items():
                if nombre_equipo in wb.sheetnames:
                    hoja_equipo = wb[nombre_equipo]
                    equipos_disponibles = []
                    
                    # Leer filas desde la fila 2 (la fila 1 son los títulos)
                    for fila in hoja_equipo.iter_rows(min_row=2, values_only=True):
                        if fila[0]:  # Si hay datos en la primera columna
                            equipos_disponibles.append(fila)
                    
                    # Obtener los títulos de las columnas (fila 1)
                    titulos = []
                    fila_titulos = list(hoja_equipo.iter_rows(min_row=1, max_row=1, values_only=True))
                    if fila_titulos:
                        for titulo in fila_titulos[0]:
                            if titulo:
                                titulos.append(str(titulo))
                    
                    # Si no hay títulos, usar los valores de la columna F de las preguntas como títulos
                    if not titulos:
                        for pregunta in preguntas_equipo:
                            if pregunta['valor']:
                                titulos.append(str(pregunta['valor']))
                    
                    equipos_por_tipo[nombre_equipo] = {
                        'preguntas': preguntas_equipo,
                        'equipos_disponibles': equipos_disponibles,
                        'titulos': titulos
                    }
            
            items_por_tipo[tipo] = items
    
    wb.close()
    
    return render_template('rellenar.html', 
                         title=title,
                         tipos=tipos,
                         items_por_tipo=items_por_tipo,
                         equipos_por_tipo=equipos_por_tipo)

@app.route('/comparar')
def comparar():
    return render_template('comparar.html')

@app.route('/comparar_upload', methods=['POST'])
def comparar_upload():
    if 'file_preventivos' not in request.files or 'file_emplazamientos' not in request.files:
        flash('Debes subir ambos archivos')
        return redirect(url_for('comparar'))
    
    file_preventivos = request.files['file_preventivos']
    file_emplazamientos = request.files['file_emplazamientos']
    
    if file_preventivos.filename == '' or file_emplazamientos.filename == '':
        flash('Debes subir ambos archivos')
        return redirect(url_for('comparar'))
    
    if file_preventivos and allowed_file(file_preventivos.filename) and file_emplazamientos and allowed_file(file_emplazamientos.filename):
        filename_preventivos = secure_filename(file_preventivos.filename)
        filename_emplazamientos = secure_filename(file_emplazamientos.filename)
        
        filepath_preventivos = os.path.join(app.config['UPLOAD_FOLDER'], filename_preventivos)
        filepath_emplazamientos = os.path.join(app.config['UPLOAD_FOLDER'], filename_emplazamientos)
        
        file_preventivos.save(filepath_preventivos)
        file_emplazamientos.save(filepath_emplazamientos)
        
        try:
            # Leer archivos
            wb_preventivos = openpyxl.load_workbook(filepath_preventivos)
            wb_emplazamientos = openpyxl.load_workbook(filepath_emplazamientos)
            
            # Procesar datos
            lista_emplazamientos = leer_excel_emplazamientos(wb_emplazamientos)
            lista_preventivos = leer_excel_preventivos(wb_preventivos)
            
            # Buscar coincidencias
            lista_resultado, lista_multiple = coincidencias(lista_preventivos, lista_emplazamientos)
            
            # Cerrar archivos
            wb_preventivos.close()
            wb_emplazamientos.close()
            
            # Si hay múltiples coincidencias, mostrar página de selección
            if lista_multiple:
                # Guardar datos en sesión para usarlos después
                session['lista_resultado'] = lista_resultado
                session['lista_multiple'] = lista_multiple
                
                # Limpiar archivos temporales
                os.remove(filepath_preventivos)
                os.remove(filepath_emplazamientos)
                
                return render_template('seleccionar_coincidencias.html', 
                                     lista_multiple=lista_multiple,
                                     total_coincidencias=len(lista_resultado),
                                     total_multiple=len(lista_multiple))
            
            # Si no hay múltiples coincidencias, generar CSV directamente
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_filename = f"preventivos_{timestamp}.csv"
            csv_path = os.path.join(DOWNLOAD_FOLDER, csv_filename)
            
            with open(csv_path, mode="w", newline="", encoding="utf-8") as archivo:
                escritor = csv.writer(archivo)
                escritor.writerow(["Folder name", "Folder color", "Latitude", "Longitude", "Title", "Description", "Color"])
                escritor.writerows(lista_resultado)
            
            # Limpiar archivos temporales
            os.remove(filepath_preventivos)
            os.remove(filepath_emplazamientos)
            
            return send_file(csv_path, as_attachment=True, download_name=csv_filename)
            
        except Exception as e:
            # Limpiar archivos temporales en caso de error
            if os.path.exists(filepath_preventivos):
                os.remove(filepath_preventivos)
            if os.path.exists(filepath_emplazamientos):
                os.remove(filepath_emplazamientos)
            flash(f'Error al procesar los archivos: {str(e)}')
            return redirect(url_for('comparar'))
    
    flash('Tipo de archivo no permitido. Solo se permiten archivos .xlsx')
    return redirect(url_for('comparar'))

@app.route('/procesar_selecciones', methods=['POST'])
def procesar_selecciones():
    data = request.json
    selecciones = data.get('selecciones', [])
    
    # Recuperar datos de la sesión
    lista_resultado = session.get('lista_resultado', [])
    lista_multiple = session.get('lista_multiple', [])
    
    # Procesar selecciones del usuario
    for sel in selecciones:
        index = sel['index']
        seleccion = sel['seleccion']
        
        if index < len(lista_multiple):
            item = lista_multiple[index]
            coincidencia_seleccionada = item['coincidencias'][seleccion]
            # Remover el último elemento (nombre del emplazamiento) para mantener el formato original
            coincidencia_formateada = coincidencia_seleccionada[:-1]
            lista_resultado.append(coincidencia_formateada)
    
    # Generar CSV
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"preventivos_{timestamp}.csv"
    csv_path = os.path.join(DOWNLOAD_FOLDER, csv_filename)
    
    with open(csv_path, mode="w", newline="", encoding="utf-8") as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(["Folder name", "Folder color", "Latitude", "Longitude", "Title", "Description", "Color"])
        escritor.writerows(lista_resultado)
    
    # Limpiar sesión
    session.pop('lista_resultado', None)
    session.pop('lista_multiple', None)
    
    return jsonify({
        'success': True,
        'download_url': f'/download_csv/{csv_filename}'
    })

@app.route('/download_csv/<filename>')
def download_csv(filename):
    csv_path = os.path.join(DOWNLOAD_FOLDER, filename)
    return send_file(csv_path, as_attachment=True, download_name=filename)

@app.route('/guardar_equipo', methods=['POST'])
def guardar_equipo():
    """Guarda un nuevo equipo en la hoja correspondiente del Excel."""
    data = request.json
    nombre_equipo = data.get('nombre_equipo')
    datos_equipo = data.get('datos_equipo', [])
    
    if not nombre_equipo or not datos_equipo:
        return jsonify({'success': False, 'message': 'Faltan datos del equipo'})
    
    try:
        # Abrir el archivo Excel
        wb = openpyxl.load_workbook(MAPA_FILE)
        
        # Verificar si existe la hoja del equipo
        if nombre_equipo not in wb.sheetnames:
            wb.close()
            return jsonify({'success': False, 'message': f'No existe la hoja {nombre_equipo}'})
        
        hoja = wb[nombre_equipo]
        
        # Encontrar la primera fila vacía
        fila_vacia = None
        for fila in hoja.iter_rows(min_row=2):
            if not fila[0].value:  # Si la primera celda está vacía
                fila_vacia = fila[0].row
                break
        
        # Si no hay fila vacía, agregar al final
        if fila_vacia is None:
            fila_vacia = hoja.max_row + 1
        
        # Escribir los datos del equipo
        for col_index, valor in enumerate(datos_equipo):
            hoja.cell(row=fila_vacia, column=col_index + 1, value=valor)
        
        # Guardar el Excel
        wb.save(MAPA_FILE)
        wb.close()
        
        return jsonify({'success': True, 'message': 'Equipo guardado correctamente'})
        
    except PermissionError:
        return jsonify({'success': False, 'message': 'El archivo mapa_preventivos.xlsx está abierto en otro programa. Por favor, ciérralo antes de continuar.'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al guardar el equipo: {str(e)}'})

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No se ha seleccionado ningún archivo')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No se ha seleccionado ningún archivo')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Validar archivo
        es_valido, mensaje = validar_archivo_preventivo(filepath)
        if not es_valido:
            os.remove(filepath)
            flash(f'Archivo no válido: {mensaje}')
            return redirect(url_for('index'))
        
        # Analizar archivo
        celdas_info, error, info_especial = analizar_archivo_excel(filepath)
        if error:
            os.remove(filepath)
            flash(f'Error al analizar el archivo: {error}')
            return redirect(url_for('index'))
        
        # Obtener tipo de preventivo
        tipo_preventivo = obtener_tipo_preventivo(filename)
        
        # Asegurar mapa_preventivos.xlsx
        exito, error_msg = asegurar_mapa_preventivos(tipo_preventivo)
        if not exito:
            os.remove(filepath)
            flash(error_msg)
            return redirect(url_for('index'))
        
        # Limpiar archivo temporal
        os.remove(filepath)
        
        # Renderizar página de configuración
        return render_template('configurar.html', 
                             tipo_preventivo=tipo_preventivo,
                             celdas_info=celdas_info,
                             info_especial=info_especial,
                             filename=filename)
    
    flash('Tipo de archivo no permitido. Solo se permiten archivos .xlsx')
    return redirect(url_for('index'))

@app.route('/guardar_config', methods=['POST'])
def guardar_config():
    data = request.json
    tipo_preventivo = data.get('tipo_preventivo')
    configuraciones = data.get('configuraciones', [])
    info_especial = data.get('info_especial', {})
    
    exito, error_msg = guardar_configuracion_mapa(tipo_preventivo, configuraciones, info_especial)
    if exito:
        return jsonify({'success': True, 'message': 'Configuración guardada correctamente'})
    else:
        return jsonify({'success': False, 'message': error_msg})

@app.route('/guardar_respuestas', methods=['POST'])
def guardar_respuestas():
    """Guarda las respuestas del formulario en Google Drive."""
    data = request.json
    title = data.get('title', '')
    tipos = data.get('tipos', [])
    respuestas = data.get('respuestas', {})
    
    if not title or not tipos:
        return jsonify({'success': False, 'message': 'Faltan datos: title o tipos'})
    
    try:
        # Leer mapa_preventivos.xlsx para obtener la estructura
        wb_mapa = openpyxl.load_workbook(MAPA_FILE)
        
        # Crear nuevo workbook para las respuestas
        wb_respuestas = openpyxl.Workbook()
        # Eliminar la hoja por defecto
        if 'Sheet' in wb_respuestas.sheetnames:
            del wb_respuestas['Sheet']
        
        # Obtener el año actual
        año_actual = datetime.now().year
        
        # Por cada tipo de preventivo, crear una hoja
        for tipo in tipos:
            if tipo not in wb_mapa.sheetnames:
                continue
            
            # Crear hoja para este tipo
            hoja_mapa = wb_mapa[tipo]
            hoja_respuestas = wb_respuestas.create_sheet(title=tipo)
            
            # Copiar fila 1 (encabezados) de mapa_preventivos.xlsx
            for col_idx, cell in enumerate(hoja_mapa[1], start=1):
                hoja_respuestas.cell(row=1, column=col_idx, value=cell.value)
            
            # Copiar fila 2 (elemento): columnas A-D de mapa_preventivos, columna E = title
            for col_idx in range(1, 5):  # Columnas A-D (1-4)
                hoja_respuestas.cell(row=2, column=col_idx, value=hoja_mapa.cell(row=2, column=col_idx).value)
            hoja_respuestas.cell(row=2, column=5, value=title)  # Columna E = title
            
            # Copiar fila 3 (posición validación): columnas A-D de mapa_preventivos
            for col_idx in range(1, 5):  # Columnas A-D (1-4)
                hoja_respuestas.cell(row=3, column=col_idx, value=hoja_mapa.cell(row=3, column=col_idx).value)
            
            # Añadir fila de encabezados de datos (fila 4)
            hoja_respuestas.cell(row=4, column=1, value='Pregunta')
            hoja_respuestas.cell(row=4, column=2, value='Posicion Celda')
            hoja_respuestas.cell(row=4, column=3, value='Posicion Respuesta')
            hoja_respuestas.cell(row=4, column=4, value='Hoja')
            hoja_respuestas.cell(row=4, column=5, value='Respuesta')
            
            # Llenar datos de respuestas
            fila_actual = 5
            pregunta_index = 0
            for fila in hoja_mapa.iter_rows(min_row=2):
                pregunta = fila[0].value
                if not pregunta:
                    continue
                
                config = fila[4].value if len(fila) > 4 else None
                if not config or config.lower() not in ['lista', 'rellenar'] and not config.lower().startswith('equipo'):
                    pregunta_index += 1
                    continue
                
                # Obtener información de la pregunta
                posicion_celda = fila[1].value if len(fila) > 1 else ''
                posicion_respuesta = fila[2].value if len(fila) > 2 else ''
                hoja_nombre = fila[3].value if len(fila) > 3 else ''
                
                # Obtener la respuesta del formulario
                # El nombre del campo en el formulario es tipo_indice
                campo_nombre = f"{tipo}_{pregunta_index}"
                respuesta = respuestas.get(campo_nombre, '')
                
                # Si es equipo, buscar la respuesta en el campo correspondiente
                if config and config.lower().startswith('equipo'):
                    nombre_equipo = config
                    # Buscar campos de equipo
                    for key, value in respuestas.items():
                        if key.startswith(f'equipo_{nombre_equipo}_'):
                            respuesta = value
                            break
                
                # Añadir fila con los datos
                hoja_respuestas.cell(row=fila_actual, column=1, value=pregunta)
                hoja_respuestas.cell(row=fila_actual, column=2, value=posicion_celda)
                hoja_respuestas.cell(row=fila_actual, column=3, value=posicion_respuesta)
                hoja_respuestas.cell(row=fila_actual, column=4, value=hoja_nombre)
                hoja_respuestas.cell(row=fila_actual, column=5, value=respuesta)
                
                fila_actual += 1
                pregunta_index += 1
        
        wb_mapa.close()
        
        # Guardar el archivo Excel temporalmente
        nombre_archivo = f"{title}.xlsx"
        ruta_temporal = os.path.join(DOWNLOAD_FOLDER, nombre_archivo)
        wb_respuestas.save(ruta_temporal)
        wb_respuestas.close()
        
        # Subir a Google Drive usando Service Account
        carpeta_destino = f"preventivos/gipuzcoa/{año_actual}"
        exito_drive, mensaje_drive = subir_archivo_drive(ruta_temporal, nombre_archivo, carpeta_destino)
        
        # Limpiar archivo temporal
        os.remove(ruta_temporal)
        
        if exito_drive:
            return jsonify({
                'success': True, 
                'message': f'Respuestas guardadas correctamente en Google Drive: {mensaje_drive}'
            })
        else:
            return jsonify({
                'success': False, 
                'message': f'Archivo guardado localmente pero error al subir a Google Drive: {mensaje_drive}'
            })
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al guardar respuestas: {str(e)}'})

if __name__ == '__main__':
    app.run(debug=True)
