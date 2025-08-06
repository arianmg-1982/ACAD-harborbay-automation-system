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
    """Carga y procesa los datos desde el archivo CSV, fusionando filas con el mismo ID de torre y nivel."""
    if not os.path.exists(CSV_INPUT_FILE):
        logging.error(f"Error crítico: El archivo de datos '{CSV_INPUT_FILE}' no fue encontrado.")
        raise FileNotFoundError(f"No se encontró {CSV_INPUT_FILE}")

    torres = defaultdict(lambda: {
        "nombre": "",
        "niveles": defaultdict(lambda: defaultdict(int)),
        "switches": {},
        "modelos_sw": {}
    })

    with open(CSV_INPUT_FILE, mode='r', encoding=ENCODING) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            try:
                torre_id = int(row['TORRE'])
                nivel_id = int(row['NIVEL'])

                if row['torre_nombre']:
                    torres[torre_id]['nombre'] = row['torre_nombre']

                # Fusionar datos del nivel
                for key, value in row.items():
                    if not value: continue

                    # Fusionar cantidades
                    if 'Qty' in key or 'switch_' in key and '_modelo' not in key:
                        try:
                            torres[torre_id]['niveles'][nivel_id][key] += int(value)
                        except (ValueError, TypeError):
                            pass # Ignorar si no es un número
                    # Sobrescribir modelos/nombres
                    elif '_modelo' in key or '_nombre' in key:
                         torres[torre_id]['niveles'][nivel_id][key] = value

                # Consolidar switches y modelos por torre
                for sw_key in row:
                    if sw_key.startswith('switch_') and '_modelo' not in sw_key and (row[sw_key] or '0') != '0':
                        sw_name = "SW-" + sw_key.split('_')[1].upper()
                        modelo_key = sw_key + "_modelo"
                        if row.get(modelo_key):
                            torres[torre_id]['switches'][sw_name] = row.get(modelo_key, '')

            except (ValueError, KeyError) as e:
                logging.error(f"Error procesando fila del CSV: {row}. Error: {e}")
                continue

    # Convertir a lista ordenada por ID de torre
    torres_list = []
    for key, value in sorted(torres.items()):
        # Convertir defaultdict a dict normal para evitar problemas posteriores
        value['niveles'] = {k: dict(v) for k, v in value['niveles'].items()}
        torres_list.append(value)

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

def lisp_dibujar_texto(f, punto, altura, texto, capa="Textos", color=7, justificacion="C"):
    """Dibuja texto en LISP de forma no interactiva con justificación opcional."""
    texto_escaped = texto.replace('"', '\\"')
    lisp_seleccionar_capa_y_color(f, capa, color)
    lisp_escribir(f, f'(command "-TEXT" "S" "Standard" "J" "{justificacion}" (list {punto[0]} {punto[1]}) {altura} 0 "{texto_escaped}")')

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
    capa = "UPS" if nombre == "SW-UPS" else "Switches"
    color = cfg['CAPAS'][capa]
    lisp_seleccionar_capa_y_color(f, capa, color)

    p1 = (x, y)
    p2 = (x + ancho, y + alto)
    lisp_dibujar_rectangulo(f, p1, p2)

    texto_completo = f"{nombre} ({modelo})" if modelo else nombre
    lisp_dibujar_texto(f, (x + ancho / 2, y + alto / 2), cfg['SWITCH_TEXTO_ALTURA'], texto_completo)

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
        coords = defaultdict(lambda: {'switches': {}, 'dispositivos': {}})

        x_torre_actual = cfg['X_INICIAL']
        for i, torre in enumerate(torres):
            torre['id'] = i
            torre['x'] = x_torre_actual
            x_torre_actual += cfg['LONGITUD_PISO'] + cfg['SEPARACION_ENTRE_TORRES']

        y_nivel_actual = cfg['Y_INICIAL']
        alturas_niveles = {}
        niveles_ordenados = sorted(list(set(n_id for t in torres for n_id in t['niveles'])))
        for nivel_id in niveles_ordenados:
            alturas_niveles[nivel_id] = y_nivel_actual
            y_nivel_actual += cfg['ESPACIO_ENTRE_NIVELES']

        # --- DIBUJAR NIVELES, ETIQUETAS, EQUIPOS Y CABLES (LÓGICA UNIFICADA) ---
        lisp_escribir(f, "\n; === INICIO DE DIBUJO DE TORRES Y CABLEADO ===")

        # 1. DIBUJAR NIVELES Y ETIQUETAS
        lisp_escribir(f, '(princ "\\nDibujando lineas de Nivel...")')
        for nivel_id, y_nivel in alturas_niveles.items():
            lisp_seleccionar_capa_y_color(f, "Niveles", cfg['CAPAS']['Niveles'])
            p1 = (cfg['X_INICIAL'] - 100, y_nivel)
            p2 = (torres[-1]['x'] + cfg['LONGITUD_PISO'] + 100, y_nivel)
            lisp_dibujar_linea(f, p1, p2)
            nivel_nombre = next((t['niveles'][nivel_id]['nivel_nombre'] for t in torres if nivel_id in t['niveles']), f"NIVEL {nivel_id}")
            lisp_dibujar_texto(f, (p1[0] - 10, y_nivel), 15, nivel_nombre, justificacion="MR")
        lisp_escribir(f, '(princ "DONE.")')

        # Diccionario para almacenar las coordenadas de los switches para el cableado de fibra
        switch_coords_para_fibra = defaultdict(dict)

        # 2. DIBUJAR TORRES, SWITCHES, DISPOSITIVOS Y CABLEADO UTP
        for torre in torres:
            torre_id = torre['id']
            x_base = torre['x']
            lisp_escribir(f, f'\n; === DIBUJAR TORRE: {torre["nombre"]} ===')
            lisp_escribir(f, f'(princ "\\n>> Dibujando Torre: {torre["nombre"]}...")')

            y_etiqueta_torre = alturas_niveles[min(alturas_niveles.keys())] - cfg['TORRE_LABEL_OFFSET_Y']
            lisp_dibujar_texto(f, (x_base + cfg['LONGITUD_PISO']/2, y_etiqueta_torre), cfg['TORRE_LABEL_ALTURA'], torre['nombre'])

            lisp_escribir(f, '(princ "\\n   - Dibujando switches...")')
            y_sotano = alturas_niveles.get(0, cfg['Y_INICIAL'])
            switches_en_torre = sorted([sw for sw in torre['switches'].keys() if sw.startswith('SW-')])
            y_switch_inicial = y_sotano - cfg['SWITCH_VERTICAL_SPACING'] * 2

            switch_coords_torre_actual = {}
            for i, sw_nombre in enumerate(switches_en_torre):
                x_switch = x_base + 50
                y_switch = y_switch_inicial - i * cfg['SWITCH_VERTICAL_SPACING']
                sw_modelo = torre['switches'][sw_nombre]
                dibujar_switch(f, cfg, x_switch, y_switch, sw_nombre, sw_modelo)
                switch_coords_torre_actual[sw_nombre] = (x_switch, y_switch + cfg['SWITCH_ALTO']/2)
                switch_coords_para_fibra[torre_id][sw_nombre] = (x_switch + cfg['SWITCH_ANCHO']/2, y_switch)
            lisp_escribir(f, '(princ "DONE.")')

            lisp_escribir(f, f'(princ "\\n   - Dibujando dispositivos y cableado UTP...")')
            x_canaleta_utp = torre['x'] + cfg['LONGITUD_PISO'] + cfg['CABLE_TRUNK_OFFSET_X']
            cables_por_troncal = defaultdict(int)
            puntos_altos_troncal = {}

            for nivel_id, nivel_data in sorted(torre['niveles'].items(), reverse=True):
                y_nivel = alturas_niveles[nivel_id]
                x_dispositivo = x_base + 50
                for tipo_qty, conf_disp in cfg['DISPOSITIVOS'].items():
                    cantidad = nivel_data.get(tipo_qty, 0)
                    if cantidad > 0:
                        y_dispositivo = y_nivel + cfg['DISPOSITIVO_Y_OFFSET']
                        globals()[f"dibujar_icono_{conf_disp['icono']}"](f, cfg, x_dispositivo, y_dispositivo)
                        lisp_dibujar_texto(f, (x_dispositivo, y_dispositivo - 20), 10, f"{cantidad}x{conf_disp['label']}")

                        sw_tipo = cfg['MAPEO_SWITCH'].get(tipo_qty)
                        if sw_tipo and sw_tipo in switch_coords_torre_actual:
                            p_origen = (x_dispositivo, y_dispositivo)
                            p_canaleta = (x_canaleta_utp, y_dispositivo)
                            lisp_seleccionar_capa_y_color(f, "Cables_UTP", cfg['CAPAS']['Cables_UTP'])
                            lisp_dibujar_linea(f, p_origen, p_canaleta) # Latiguillo
                            cables_por_troncal[sw_tipo] += cantidad
                            if sw_tipo not in puntos_altos_troncal:
                                puntos_altos_troncal[sw_tipo] = p_canaleta
                        x_dispositivo += cfg['DISPOSITIVO_ESPACIADO_X']

            for sw_tipo, p_troncal_sup in puntos_altos_troncal.items():
                p_destino_sw = switch_coords_torre_actual[sw_tipo]
                p_troncal_inf = (x_canaleta_utp, p_destino_sw[1])
                lisp_dibujar_linea(f, p_troncal_sup, p_troncal_inf) # Troncal
                lisp_dibujar_linea(f, p_troncal_inf, (p_destino_sw[0] + cfg['SWITCH_ANCHO'], p_destino_sw[1])) # Latiguillo
                lisp_dibujar_texto(f, (x_canaleta_utp + 5, (p_troncal_sup[1] + p_troncal_inf[1])/2), 8, f"{cables_por_troncal[sw_tipo]}xCAT6A", justificacion="ML")

            lisp_escribir(f, '(princ "DONE.")')

        # 3. Cableado de Fibra y Alimentación
        lisp_escribir(f, '(princ "\\nDibujando cables de Fibra y Alimentacion...")')
        mdf_switch_coords = switch_coords_para_fibra[0]
        y_canaleta_fibra = cfg['Y_INICIAL'] - cfg['TORRE_LABEL_OFFSET_Y'] - cfg['CABLE_TRUNK_OFFSET_Y']

        for sw_tipo, mdf_sw_coord in sorted(mdf_switch_coords.items(), key=lambda item: item[1][1]):
            sw_conf = cfg['SWITCH_CONFIG'].get(sw_tipo, {})
            lisp_seleccionar_capa_y_color(f, sw_conf.get('capa', 'Fibra_Data'), sw_conf.get('color', 1))

            p_mdf_salida = mdf_sw_coord
            p_mdf_canaleta = (p_mdf_salida[0], y_canaleta_fibra)
            lisp_dibujar_linea(f, p_mdf_salida, p_mdf_canaleta)

            torres_conectadas = [t for t in torres if t['id'] != 0 and sw_tipo in switch_coords_para_fibra[t['id']]]
            if not torres_conectadas:
                y_canaleta_fibra -= 30
                continue

            ultima_idf_coord = switch_coords_para_fibra[torres_conectadas[-1]['id']][sw_tipo]
            p_canaleta_fin = (ultima_idf_coord[0], y_canaleta_fibra)
            lisp_dibujar_linea(f, p_mdf_canaleta, p_canaleta_fin)
            lisp_dibujar_texto(f, ( (p_mdf_canaleta[0]+p_canaleta_fin[0])/2, y_canaleta_fibra+5 ), 8, f"{len(torres_conectadas)}x {sw_conf.get('label', 'FO')}")

            for torre in torres_conectadas:
                idf_sw_coord = switch_coords_para_fibra[torre['id']][sw_tipo]
                p_idf_canaleta = (idf_sw_coord[0], y_canaleta_fibra)
                lisp_dibujar_linea(f, p_idf_canaleta, idf_sw_coord)

            y_canaleta_fibra -= 30
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
