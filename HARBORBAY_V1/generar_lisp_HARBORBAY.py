import csv
import json
import os
import logging
from collections import defaultdict
from datetime import datetime

# --- CONFIGURACION GENERAL ---
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
CSV_INPUT_FILE = os.path.join(SCRIPT_DIR, "torres.csv")
LISP_OUTPUT_FILE = os.path.join(SCRIPT_DIR, "dibujo_red.lsp")
BOM_OUTPUT_FILE = os.path.join(SCRIPT_DIR, "bom_proyecto.txt")
LOG_FILE = os.path.join(SCRIPT_DIR, "logs.TXT")
ENCODING = "utf-8"

# --- CONFIGURACION DE LOGS ---
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)

# Log a archivo
file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding=ENCODING)
file_handler.setFormatter(log_formatter)
root_logger.addHandler(file_handler)

# Log a consola
console_handler = logging.StreamHandler()
console_handler.setFormatter(log_formatter)
console_handler.setLevel(logging.INFO)
root_logger.addHandler(console_handler)


def cargar_configuracion():
    """Carga el archivo de configuración JSON."""
    try:
        with open(CONFIG_FILE, 'r', encoding=ENCODING) as f:
            logging.info(f"Cargando configuración desde '{CONFIG_FILE}'...")
            return json.load(f)
    except FileNotFoundError:
        logging.error(f"Error crítico: El archivo de configuración '{CONFIG_FILE}' no fue encontrado.")
        raise
    except json.JSONDecodeError:
        logging.error(f"Error crítico: El archivo '{CONFIG_FILE}' no es un JSON válido.")
        raise

def cargar_datos_csv():
    """Carga y procesa los datos desde el archivo CSV."""
    if not os.path.exists(CSV_INPUT_FILE):
        logging.error(f"Error crítico: El archivo de datos '{CSV_INPUT_FILE}' no fue encontrado.")
        raise FileNotFoundError(f"No se encontró {CSV_INPUT_FILE}")

    torres = defaultdict(lambda: {
        "nombre": "",
        "niveles": defaultdict(dict),
        "switches": {},
        "modelos_sw": {}
    })

    with open(CSV_INPUT_FILE, mode='r', encoding=ENCODING) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            try:
                torre_id = int(row['TORRE'])
                nivel_id = int(row['NIVEL'])

                torres[torre_id]['nombre'] = row['torre_nombre']

                # Almacenar datos del nivel
                for key, value in row.items():
                    # Convertir cantidades a enteros, manejando celdas vacías
                    if 'Qty' in key or 'switch_' in key and '_modelo' not in key:
                         try:
                            torres[torre_id]['niveles'][nivel_id][key] = int(value or 0)
                         except (ValueError, TypeError):
                            torres[torre_id]['niveles'][nivel_id][key] = 0
                    else:
                        torres[torre_id]['niveles'][nivel_id][key] = value

                # Consolidar switches y modelos por torre
                for sw_key in row:
                    if sw_key.startswith('switch_') and '_modelo' not in sw_key and (row[sw_key] or '0') != '0':
                        sw_name = "SW-" + sw_key.split('_')[1].upper()
                        modelo_key = sw_key + "_modelo"
                        torres[torre_id]['switches'][sw_name] = row.get(modelo_key, '')

            except (ValueError, KeyError) as e:
                logging.error(f"Error procesando fila del CSV: {row}. Error: {e}")
                continue

    # Convertir a lista ordenada por ID de torre
    torres_list = [value for key, value in sorted(torres.items())]
    logging.info(f"Cargados datos de {len(torres_list)} torres desde '{CSV_INPUT_FILE}'.")
    return torres_list

# --- FUNCIONES DE DIBUJO LISP ---

def lisp_escribir(f, comando):
    """Escribe una línea en el archivo LISP con salto de línea."""
    f.write(comando + "\n")

def lisp_crear_capa(f, nombre, color):
    """Genera el comando LISP para crear una capa de forma robusta."""
    nombre_escaped = nombre.replace('"', '\\"')
    lisp_escribir(f, f'(command "-LAYER" "N" "{nombre_escaped}" "C" "{color}" "" "")')

def lisp_seleccionar_capa_y_color(f, capa, color):
    """Genera comandos para seleccionar capa y color."""
    lisp_escribir(f, f'(command "-LAYER" "S" "{capa}" "")')
    lisp_escribir(f, f'(command "-COLOR" "{color}")')

def lisp_dibujar_linea(f, p1, p2):
    """Dibuja una línea en LISP."""
    lisp_escribir(f, f'(command "_.LINE" (list {p1[0]} {p1[1]}) (list {p2[0]} {p2[1]}) "")')

def lisp_dibujar_texto(f, punto, altura, texto, capa="Textos", color=7):
    """Dibuja texto en LISP de forma no interactiva."""
    texto_escaped = texto.replace('"', '\\"')
    lisp_seleccionar_capa_y_color(f, capa, color)
    lisp_escribir(f, f'(command "-TEXT" "S" "Standard" "J" "C" (list {punto[0]} {punto[1]}) {altura} 0 "{texto_escaped}")')

def lisp_dibujar_rectangulo(f, p1, p2):
    """Dibuja un rectángulo usando PLINE para máxima compatibilidad."""
    (x1, y1) = p1
    (x2, y2) = p2
    lisp_dibujar_polilinea(f, [(x1, y1), (x2, y1), (x2, y2), (x1, y2)], cerrada=True)

def lisp_dibujar_circulo(f, centro, radio):
    """Dibuja un círculo en LISP."""
    lisp_escribir(f, f'(command "_.CIRCLE" (list {centro[0]} {centro[1]}) {radio})')

def lisp_dibujar_polilinea(f, puntos, cerrada=False):
    """Dibuja una polilínea."""
    comando = '(command "_.PLINE"'
    for p in puntos:
        comando += f' (list {p[0]} {p[1]})'
    if cerrada:
        comando += ' "C")'
    else:
        comando += ' "")'
    lisp_escribir(f, comando)

def lisp_dibujar_hatch(f):
    """Rellena el último objeto creado de forma segura."""
    lisp_escribir(f, '(command "-HATCH" "S" "L" "" "")')

def lisp_dibujar_arco_eliptico(f, centro, eje_x, eje_y, angulo_inicio, angulo_fin):
    """Dibuja un arco elíptico."""
    lisp_escribir(f, f'(command "_.ELLIPSE" "A" (list {centro[0]} {centro[1]}) (list {eje_x[0]} {eje_x[1]}) (list {eje_y[0]} {eje_y[1]}) {angulo_inicio} {angulo_fin})')

# --- FUNCIONES DE DIBUJO DE COMPONENTES ---

def dibujar_icono_ap(f, cfg, x, y):
    """Dibuja el icono de un AP: Triángulo con 3 círculos concéntricos."""
    capa_info = cfg['CAPAS']['APs']
    lisp_seleccionar_capa_y_color(f, "APs", capa_info)

    base = 20
    altura = 25
    p1 = (x - base / 2, y)
    p2 = (x + base / 2, y)
    p3 = (x, y + altura)
    lisp_dibujar_polilinea(f, [p1, p2, p3, p1])

    centro_circulos = (x, y + altura)
    radios = [10, 21.25, 30]
    for radio in radios:
        lisp_dibujar_circulo(f, centro_circulos, radio)

def dibujar_icono_telefono(f, cfg, x, y):
    """Dibuja el icono de un Teléfono: Rectángulo (cuerpo) + círculo (auricular)."""
    capa_info = cfg['CAPAS']['Telefonos']
    lisp_seleccionar_capa_y_color(f, "Telefonos", capa_info)

    cuerpo_ancho, cuerpo_alto = 20, 30
    auricular_radio = 5
    p1 = (x - cuerpo_ancho / 2, y)
    p2 = (x + cuerpo_ancho / 2, y + cuerpo_alto)
    lisp_dibujar_rectangulo(f, p1, p2)
    lisp_dibujar_circulo(f, (x, y + cuerpo_alto + auricular_radio + 2), auricular_radio)

def dibujar_icono_tv(f, cfg, x, y):
    """Dibuja el icono de una TV: Rectángulo (pantalla) + triángulo (soporte)."""
    capa_info = cfg['CAPAS']['TVs']
    lisp_seleccionar_capa_y_color(f, "TVs", capa_info)

    pantalla_ancho, pantalla_alto = 40, 25
    p1_pantalla = (x - pantalla_ancho / 2, y)
    p2_pantalla = (x + pantalla_ancho / 2, y + pantalla_alto)
    lisp_dibujar_rectangulo(f, p1_pantalla, p2_pantalla)

    soporte_base = 20
    p1_soporte = (x - soporte_base / 2, y)
    p2_soporte = (x + soporte_base / 2, y)
    p3_soporte = (x, y - 10)
    lisp_dibujar_polilinea(f, [p1_soporte, p2_soporte, p3_soporte, p1_soporte])

def dibujar_icono_camara(f, cfg, x, y):
    """Dibuja el icono de una Cámara: Caja + arco elíptico."""
    capa_info = cfg['CAPAS']['Camaras']
    lisp_seleccionar_capa_y_color(f, "Camaras", capa_info)

    caja_ancho, caja_alto = 20, 15
    p1 = (x - caja_ancho / 2, y)
    p2 = (x + caja_ancho / 2, y + caja_alto)
    lisp_dibujar_rectangulo(f, p1, p2)
    # Simulación de lente
    lisp_dibujar_circulo(f, (x, y + caja_alto / 2), 3)

def dibujar_icono_dato(f, cfg, x, y):
    """Dibuja el icono de una Toma de Datos: Triángulo rojo relleno."""
    capa_info = cfg['CAPAS']['Datos']
    lisp_seleccionar_capa_y_color(f, "Datos", capa_info)

    base = 20
    altura = 20
    p1 = (x - base / 2, y)
    p2 = (x + base / 2, y)
    p3 = (x, y + altura)
    lisp_dibujar_polilinea(f, [p1, p2, p3], cerrada=True)
    lisp_dibujar_hatch(f)

def dibujar_switch(f, cfg, x, y, nombre, modelo):
    """Dibuja un switch con su etiqueta."""
    ancho, alto = cfg['SWITCH_ANCHO'], cfg['SWITCH_ALTO']
    lisp_seleccionar_capa_y_color(f, "Switches", cfg['CAPAS']['Switches'])
    p1 = (x, y)
    p2 = (x + ancho, y + alto)
    lisp_dibujar_rectangulo(f, p1, p2)

    texto_completo = f"{nombre} ({modelo})" if modelo else nombre
    lisp_dibujar_texto(f, (x + cfg['SWITCH_TEXTO_OFFSET_X'], y + cfg['SWITCH_TEXTO_OFFSET_Y']), cfg['SWITCH_TEXTO_ALTURA'], texto_completo)

def dibujar_ups(f, cfg, x, y):
    """Dibuja el icono de la UPS."""
    ancho, alto = cfg['UPS_ANCHO'], cfg['UPS_ALTO']
    lisp_seleccionar_capa_y_color(f, "UPS", cfg['CAPAS']['UPS'])
    p1 = (x, y)
    p2 = (x + ancho, y + alto)
    lisp_dibujar_rectangulo(f, p1, p2)
    lisp_dibujar_texto(f, (x + 25, y + 25), 15, "UPS")
    lisp_dibujar_texto(f, (x + 5, y - 20), 10, "ALIMENTACION")

# --- LÓGICA PRINCIPAL DE GENERACIÓN ---

def generar_lisp(cfg, torres):
    """Función principal que genera el archivo LISP completo."""
    with open(LISP_OUTPUT_FILE, "w", encoding=ENCODING) as f:
        logging.info(f"Iniciando la generación del archivo LISP: '{LISP_OUTPUT_FILE}'...")

        # --- INICIALIZACIÓN DE AUTOCAD ---
        lisp_escribir(f, '(setq *error* (lambda (msg) (if msg (princ (strcat "\\nError: " msg)))))')
        lisp_escribir(f, '(setvar "OSMODE" 0)')
        lisp_escribir(f, '(command "_.-PURGE" "ALL" "*" "N")')
        lisp_escribir(f, '(command "_.REGEN")')
        lisp_escribir(f, '(command "VISUALSTYLE" "Wireframe")')
        lisp_escribir(f, '(command "_.UNDO" "BEGIN")')
        lisp_escribir(f, '(princ "--- INICIO DE DIBUJO AUTOMATIZADO HARBORBAY ---")')

        # --- CREACIÓN DE CAPAS ---
        lisp_escribir(f, "\n; === CREAR CAPAS ===")
        lisp_escribir(f, '(princ "\\nCreando capas...")')
        for nombre, color in cfg['CAPAS'].items():
            lisp_crear_capa(f, nombre, color)
        lisp_escribir(f, '(princ "DONE.")')

        # --- CÁLCULO DE POSICIONES Y ALMACENAMIENTO DE COORDENADAS ---
        # Este diccionario almacenará las coordenadas exactas de cada elemento para su posterior uso en cableado
        coords = defaultdict(dict)

        # 1. Calcular coordenadas X para cada torre
        x_torre_actual = cfg['X_INICIAL']
        for i, torre in enumerate(torres):
            torre['id'] = i
            torre['x'] = x_torre_actual
            x_torre_actual += cfg['LONGITUD_PISO'] + cfg['SEPARACION_ENTRE_TORRES']

        # 2. Calcular coordenadas Y para cada nivel (bottom-up)
        y_nivel_actual = cfg['Y_INICIAL']
        alturas_niveles = {}
        niveles_ordenados = sorted(list(set(n_id for t in torres for n_id in t['niveles'])))
        for nivel_id in niveles_ordenados:
            alturas_niveles[nivel_id] = y_nivel_actual
            y_nivel_actual += cfg['ESPACIO_ENTRE_NIVELES']

        # --- DIBUJAR NIVELES, ETIQUETAS Y EQUIPOS ---
        lisp_escribir(f, "\n; === DIBUJAR NIVELES Y ESTRUCTURA DE TORRES ===")
        # Dibujar líneas de nivel y etiquetas
        lisp_escribir(f, '(princ "\\nDibujando lineas de Nivel...")')
        for nivel_id, y_nivel in alturas_niveles.items():
            lisp_seleccionar_capa_y_color(f, "Niveles", cfg['CAPAS']['Niveles'])
            p1 = (cfg['X_INICIAL'] - 50, y_nivel)
            p2 = (torres[-1]['x'] + cfg['LONGITUD_PISO'] + 50, y_nivel)
            lisp_dibujar_linea(f, p1, p2)
            nivel_nombre = next((t['niveles'][nivel_id]['nivel_nombre'] for t in torres if nivel_id in t['niveles']), f"NIVEL {nivel_id}")
            lisp_dibujar_texto(f, (p1[0] + 10, y_nivel + 10), 15, nivel_nombre)
        lisp_escribir(f, '(princ "DONE.")')

        # Dibujar Torres, Switches, y Dispositivos
        for torre in torres:
            torre_id = torre['id']
            x_base = torre['x']
            lisp_escribir(f, f'\n; === DIBUJAR TORRE: {torre["nombre"]} ===')
            lisp_escribir(f, f'(princ "\\n>> Dibujando Torre: {torre["nombre"]}...")')

            # Dibujar etiqueta de la torre
            y_etiqueta_torre = alturas_niveles[min(alturas_niveles.keys())] - cfg['TORRE_LABEL_OFFSET_Y']
            lisp_dibujar_texto(f, (x_base, y_etiqueta_torre), cfg['TORRE_LABEL_ALTURA'], torre['nombre'])

            # Dibujar Switches en el Sótano (Nivel 0)
            lisp_escribir(f, '(princ "\\n   - Dibujando switches...")')
            y_sotano = alturas_niveles[0]
            y_switch = y_sotano - cfg['SWITCH_VERTICAL_SPACING'] * (len(torre['switches']))
            coords[torre_id]['switches'] = {}
            for i, (sw_nombre, sw_modelo) in enumerate(sorted(torre['switches'].items())):
                x_switch = x_base + 50
                dibujar_switch(f, cfg, x_switch, y_switch, sw_nombre, sw_modelo)
                coords[torre_id]['switches'][sw_nombre] = (x_switch + cfg['SWITCH_ANCHO'] / 2, y_switch + cfg['SWITCH_ALTO'])
                y_switch += cfg['SWITCH_VERTICAL_SPACING']
            lisp_escribir(f, '(princ "DONE.")')

            # Dibujar UPS solo en el MDF (torre 0)
            if torre_id == 0:
                lisp_escribir(f, '(princ "\\n   - Dibujando UPS...")')
                x_switch_pos = x_base + 50
                # Alinear la UPS con el switch más bajo
                y_ups = y_sotano - cfg['SWITCH_VERTICAL_SPACING'] * (len(torre['switches']))

                x_ups = x_switch_pos + cfg['SWITCH_ANCHO'] + cfg['UPS_SEPARACION']

                dibujar_ups(f, cfg, x_ups, y_ups)
                # Guardar la coordenada para el cableado (punto central de la UPS)
                coords['ups_coord'] = (x_ups + cfg['UPS_ANCHO'] / 2, y_ups + cfg['UPS_ALTO'] / 2)
                lisp_escribir(f, '(princ "DONE.")')


            # Dibujar Dispositivos por nivel
            lisp_escribir(f, f'(princ "\\n   - Dibujando dispositivos por nivel...")')
            for nivel_id, nivel_data in torre['niveles'].items():
                y_nivel = alturas_niveles[nivel_id]
                x_dispositivo = x_base + 50
                for tipo_qty, conf in cfg['DISPOSITIVOS'].items():
                    cantidad = nivel_data.get(tipo_qty, 0)
                    if cantidad > 0:
                        y_dispositivo = y_nivel + cfg['DISPOSITIVO_Y_OFFSET']
                        globals()[f"dibujar_icono_{conf['icono']}"](f, cfg, x_dispositivo, y_dispositivo)
                        lisp_dibujar_texto(f, (x_dispositivo, y_dispositivo - 20), 10, f"{cantidad}x{conf['label']}")
                        # Guardar coordenadas para cableado UTP
                        coords[torre_id][f'disp_{nivel_id}_{conf["label"]}'] = (x_dispositivo, y_dispositivo)
                        x_dispositivo += cfg['DISPOSITIVO_ESPACIADO_X']
            lisp_escribir(f, '(princ "DONE.")')

        # --- DEFINIR BANDEJAS DE CABLEADO HORIZONTAL ---
        y_sotano = alturas_niveles[0]
        y_bandeja_start = y_sotano - cfg['SWITCH_VERTICAL_SPACING'] * (len(torres[0]['switches'])) - 150

        tipos_de_cable = list(cfg['SWITCH_CONFIG'].keys()) + ['UPS']
        coords_bandejas = {}
        for i, tipo in enumerate(tipos_de_cable):
            coords_bandejas[tipo] = y_bandeja_start - (i * 40)

        # --- DIBUJAR CABLES ---
        lisp_escribir(f, "\n; === DIBUJAR CABLES ===")

        # 1. Cableado UTP Vertical dentro de cada torre
        lisp_escribir(f, '(princ "\\nDibujando cables UTP...")')
        for torre in torres:
            torre_id = torre['id']
            for nivel_id, nivel_data in torre['niveles'].items():
                for tipo_qty, conf_disp in cfg['DISPOSITIVOS'].items():
                    if nivel_data.get(tipo_qty, 0) > 0:
                        sw_tipo_mapeado = cfg['MAPEO_SWITCH'].get(tipo_qty)
                        if sw_tipo_mapeado and sw_tipo_mapeado in coords[torre_id]['switches']:
                            p_origen = coords[torre_id][f'disp_{nivel_id}_{conf_disp["label"]}']
                            p_destino_sw = coords[torre_id]['switches'][sw_tipo_mapeado]

                            p_intermedio = (p_origen[0], p_destino_sw[1] + 20)
                            lisp_seleccionar_capa_y_color(f, "Cables_UTP", cfg['CAPAS']['Cables_UTP'])
                            lisp_dibujar_polilinea(f, [p_origen, p_intermedio, (p_destino_sw[0], p_intermedio[1]), p_destino_sw])

                            # Etiqueta para el cable UTP
                            cantidad = nivel_data.get(tipo_qty, 0)
                            if cantidad > 0:
                                label_text = f"{cantidad}xUTP"
                                label_pos = (p_origen[0] + 15, (p_origen[1] + p_intermedio[1]) / 2)
                                lisp_dibujar_texto(f, label_pos, 10, label_text, "Cables_UTP", cfg['CAPAS']['Cables_UTP'])
        lisp_escribir(f, '(princ "DONE.")')

        # 2. Cableado de Fibra Óptica de MDF a IDFs
        lisp_escribir(f, '(princ "\\nDibujando cables de Fibra Optica...")')

        fibra_counts = defaultdict(int)
        for torre in torres:
            if torre['id'] == 0: continue
            for sw_tipo in torre['switches']:
                if sw_tipo in cfg['SWITCH_CONFIG']:
                    fibra_counts[sw_tipo] += 1

        mdf_coords = coords[0]

        for sw_tipo in sorted(cfg['SWITCH_CONFIG'].keys()):
            if sw_tipo not in mdf_coords['switches']: continue

            mdf_sw_coord_top = mdf_coords['switches'][sw_tipo]
            y_bandeja = coords_bandejas.get(sw_tipo, y_bandeja_start)
            sw_conf = cfg['SWITCH_CONFIG'].get(sw_tipo, {})
            lisp_seleccionar_capa_y_color(f, sw_conf.get('capa', 'Fibra'), sw_conf.get('color', 2))

            p_start_mdf = (mdf_sw_coord_top[0], mdf_sw_coord_top[1] - cfg['SWITCH_ALTO'] / 2)
            p_drop_mdf = (p_start_mdf[0] + 150, y_bandeja)

            lisp_dibujar_linea(f, p_start_mdf, p_drop_mdf)

            max_x_idf = 0
            last_idf_torre = None
            for torre in reversed(torres):
                if torre['id'] != 0 and sw_tipo in torre['switches']:
                    max_x_idf = coords[torre['id']]['switches'][sw_tipo][0]
                    last_idf_torre = torre
                    break

            if max_x_idf > 0:
                p_end_tray_bus = (max_x_idf - 150, y_bandeja)
                lisp_dibujar_linea(f, p_drop_mdf, p_end_tray_bus)

                label_text = f"{fibra_counts[sw_tipo]}xFO {sw_tipo.replace('SW-', '')}"
                label_pos = ((p_drop_mdf[0] + p_end_tray_bus[0]) / 2, y_bandeja + 10)
                lisp_dibujar_texto(f, label_pos, 10, label_text, sw_conf.get('capa', 'Fibra'), sw_conf.get('color', 2))

            for torre in torres:
                if torre['id'] != 0 and sw_tipo in torre['switches']:
                    idf_sw_coord_top = coords[torre['id']]['switches'][sw_tipo]
                    p_end_idf = (idf_sw_coord_top[0], idf_sw_coord_top[1] - cfg['SWITCH_ALTO'] / 2)
                    p_rise_idf = (p_end_idf[0] - 150, y_bandeja)
                    lisp_dibujar_linea(f, p_rise_idf, p_end_idf)

        lisp_escribir(f, '(princ "DONE.")')

        # 3. Cableado Eléctrico desde la UPS a todos los switches
        lisp_escribir(f, '(princ "\\nDibujando alimentacion desde UPS...")')
        if 'ups_coord' in coords:
            p_origen_ups = coords['ups_coord']
            y_bandeja_ups = coords_bandejas.get('UPS', y_bandeja_start)
            lisp_seleccionar_capa_y_color(f, "UPS", cfg['CAPAS']['UPS'])

            p_drop_ups = (p_origen_ups[0] + 150, y_bandeja_ups)
            lisp_dibujar_linea(f, p_origen_ups, p_drop_ups)

            max_x_switch = 0
            for torre in torres:
                for sw in torre['switches']:
                    sw_coord = coords[torre['id']]['switches'][sw]
                    max_x_switch = max(max_x_switch, sw_coord[0])

            if max_x_switch > 0:
                p_end_tray_bus = (max_x_switch - 150, y_bandeja_ups)
                lisp_dibujar_linea(f, p_drop_ups, p_end_tray_bus)

            for torre_cableado in torres:
                torre_id_cableado = torre_cableado['id']
                if torre_id_cableado in coords and 'switches' in coords[torre_id_cableado]:
                    for sw_nombre, p_destino_sw_top in coords[torre_id_cableado]['switches'].items():
                        p_end_sw = (p_destino_sw_top[0], p_destino_sw_top[1] - cfg['SWITCH_ALTO'] / 2)
                        p_rise_sw = (p_end_sw[0] - 150, y_bandeja_ups)
                        if p_rise_sw[0] > p_drop_ups[0]: # Solo dibujar si el switch está a la derecha del punto de bajada
                            lisp_dibujar_linea(f, p_rise_sw, p_end_sw)

        lisp_escribir(f, '(princ "DONE.")')


        # --- FINALIZAR DIBUJO ---
        lisp_escribir(f, "\n; === FINALIZAR DIBUJO ===")
        lisp_escribir(f, '(princ "\\nFinalizando y haciendo zoom...")')
        lisp_escribir(f, '(command "_.ZOOM" "E")')
        lisp_escribir(f, '(command "_.UNDO" "END")')
        lisp_escribir(f, '(princ "\\n--- PROCESO DE DIBUJO COMPLETADO ---")')
        logging.info(f"Archivo LISP '{LISP_OUTPUT_FILE}' generado con éxito.")

def generar_bom(cfg, torres):
    """Genera el archivo de listado de materiales (BOM)."""
    with open(BOM_OUTPUT_FILE, "w", encoding=ENCODING) as f:
        logging.info(f"Iniciando la generación del BOM: '{BOM_OUTPUT_FILE}'...")

        totales = defaultdict(int)
        switches_totales = defaultdict(int)
        modelos_switches = defaultdict(lambda: defaultdict(int))

        for torre in torres:
            for nivel in torre['niveles'].values():
                for tipo_qty in cfg['DISPOSITIVOS']:
                    totales[tipo_qty] += nivel.get(tipo_qty, 0)
            for sw_nombre, sw_modelo in torre['switches'].items():
                switches_totales[sw_nombre] += 1
                if sw_modelo:
                    modelos_switches[sw_nombre][sw_modelo] += 1

        total_puntos_red = sum(totales.values())
        total_cable_utp = total_puntos_red * 15 # Estimación de 15m por punto

        f.write("============================================================\n")
        f.write("      LISTADO DE MATERIALES (BOM) - PROYECTO HARBORBAY\n")
        f.write("============================================================\n")
        f.write(f"Fecha de Generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        f.write("--- RESUMEN DEL PROYECTO ---\n")
        f.write(f"Total de Torres (MDF+IDF): {len(torres)}\n")
        f.write(f"Total de Puntos de Red:    {total_puntos_red}\n\n")

        f.write("--- TOTAL DE DISPOSITIVOS ---\n")
        for tipo_qty, conf in cfg['DISPOSITIVOS'].items():
            f.write(f"- {conf['label']:<10}: {totales[tipo_qty]} unidades\n")
        f.write("\n")

        f.write("--- TOTAL DE SWITCHES POR TIPO Y MODELO ---\n")
        for sw_nombre, cantidad in sorted(switches_totales.items()):
            f.write(f"- {sw_nombre:<15}: {cantidad} unidades\n")
            if sw_nombre in modelos_switches:
                for modelo, q in modelos_switches[sw_nombre].items():
                    f.write(f"    - Modelo: {modelo} ({q} uds)\n")
        f.write("\n")

        f.write("--- ESTIMACIÓN DE CABLEADO ---\n")
        f.write(f"- Cable UTP CAT6A: ~{total_cable_utp:,} metros\n")
        # Cálculo simple de fibra
        total_fibra = (len(torres) -1) * len(cfg['SWITCH_CONFIG']) * 50
        f.write(f"- Fibra Óptica:    ~{total_fibra:,} metros (estimación bruta)\n\n")

        f.write("--- OBSERVACIONES ---\n")
        f.write("- Las cantidades se basan en el archivo 'torres.csv'.\n")
        f.write("- La longitud del cableado es una estimación y requiere verificación en sitio.\n")
        f.write("- Se incluye 1 UPS centralizada en el MDF.\n")
        f.write("============================================================\n")
    logging.info(f"Archivo BOM '{BOM_OUTPUT_FILE}' generado con éxito.")

def main():
    """Función principal para ejecutar el script."""
    try:
        logging.info("--- INICIO DEL PROCESO DE GENERACIÓN DE PLANOS ---")
        config = cargar_configuracion()
        datos_torres = cargar_datos_csv()

        if not datos_torres:
            logging.warning("No se encontraron datos de torres para procesar. Terminando ejecución.")
            return

        generar_lisp(config, datos_torres)
        generar_bom(config, datos_torres)

        logging.info("--- PROCESO COMPLETADO EXITOSAMENTE ---")

    except Exception as e:
        logging.critical(f"Ha ocurrido un error inesperado en la ejecución: {e}", exc_info=True)

if __name__ == "__main__":
    main()
