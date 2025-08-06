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
    """Genera el comando LISP para crear una capa."""
    nombre_escaped = nombre.replace('"', '\\"')
    lisp_escribir(f, f'(command "-LAYER" "MAKE" "{nombre_escaped}" "COLOR" "{color}" "" "")')

def lisp_seleccionar_capa_y_color(f, capa, color):
    """Genera comandos para seleccionar capa y color."""
    lisp_escribir(f, f'(setvar "CLAYER" "{capa}")')
    lisp_escribir(f, f'(command "_.COLOR" "{color}")')

def lisp_dibujar_linea(f, p1, p2):
    """Dibuja una línea en LISP."""
    lisp_escribir(f, f'(command "_.LINE" (list {p1[0]} {p1[1]}) (list {p2[0]} {p2[1]}) "")')

def lisp_dibujar_texto(f, punto, altura, texto, capa="Textos", color=7):
    """Dibuja texto en LISP."""
    texto_escaped = texto.replace('"', '\\"')
    lisp_seleccionar_capa_y_color(f, capa, color)
    lisp_escribir(f, f'(command "_.TEXT" "Style" "Standard" (list {punto[0]} {punto[1]}) {altura} 0 "{texto_escaped}")')

def lisp_dibujar_rectangulo(f, p1, p2):
    """Dibuja un rectángulo en LISP."""
    lisp_escribir(f, f'(command "_.RECTANG" (list {p1[0]} {p1[1]}) (list {p2[0]} {p2[1]}) "")')

def lisp_dibujar_circulo(f, centro, radio):
    """Dibuja un círculo en LISP."""
    lisp_escribir(f, f'(command "_.CIRCLE" (list {centro[0]} {centro[1]}) {radio})')

def lisp_dibujar_polilinea(f, puntos):
    """Dibuja una polilínea."""
    comando = '(command "_.PLINE"'
    for p in puntos:
        comando += f' (list {p[0]} {p[1]})'
    comando += ' "")'
    lisp_escribir(f, comando)

def lisp_dibujar_hatch(f, punto_interno):
    """Rellena un área cerrada de forma robusta para scripts."""
    lisp_escribir(f, f'(command "-HATCH" "P" "SOLID" "1" "0" "" (list {punto_interno[0]} {punto_interno[1]}) "")')

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
    lisp_dibujar_polilinea(f, [p1, p2, p3, p1])
    lisp_dibujar_hatch(f, (x, y + 5))

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

        # --- CÁLCULO DE POSICIONES ---
        x_torre_actual = cfg['X_INICIAL']
        max_dispositivos_por_nivel = defaultdict(int)

        for torre in torres:
            torre['x'] = x_torre_actual
            x_torre_actual += cfg['LONGITUD_PISO'] + cfg['SEPARACION_ENTRE_TORRES']
            for nivel_id, nivel_data in torre['niveles'].items():
                num_dispositivos = sum(nivel_data.get(k, 0) for k in cfg['DISPOSITIVOS'])
                if num_dispositivos > max_dispositivos_por_nivel[nivel_id]:
                    max_dispositivos_por_nivel[nivel_id] = num_dispositivos

        y_nivel_actual = cfg['Y_INICIAL']
        alturas_niveles = {}
        niveles_ordenados = sorted(max_dispositivos_por_nivel.keys())
        for nivel_id in niveles_ordenados:
            alturas_niveles[nivel_id] = y_nivel_actual
            altura_dinamica = 150 + max(0, max_dispositivos_por_nivel[nivel_id] - 5) * 20
            y_nivel_actual += altura_dinamica

        # --- DIBUJAR NIVELES Y TORRES ---
        lisp_escribir(f, "\n; === DIBUJAR NIVELES Y TORRES ===")
        lisp_escribir(f, '(princ "\\nDibujando lineas de Nivel...")')
        for nivel_id, y_nivel in alturas_niveles.items():
            lisp_seleccionar_capa_y_color(f, "Niveles", cfg['CAPAS']['Niveles'])
            p1 = (cfg['X_INICIAL'] - 50, y_nivel)
            p2 = (torres[-1]['x'] + cfg['LONGITUD_PISO'] + 50, y_nivel)
            lisp_dibujar_linea(f, p1, p2)
            nivel_nombre = torres[0]['niveles'].get(nivel_id, {}).get('nivel_nombre', f"NIVEL {nivel_id}")
            lisp_dibujar_texto(f, (p1[0] + cfg['OFFSET_NOMBRE_NIVEL_X'], y_nivel + cfg['OFFSET_NOMBRE_NIVEL_Y']), 15, nivel_nombre)
        lisp_escribir(f, '(princ "DONE.")')

        for torre in torres:
            lisp_escribir(f, f'\n; === DIBUJAR TORRE: {torre["nombre"]} ===')
            lisp_escribir(f, f'(princ "\\n>> Dibujando Torre: {torre["nombre"]}...")')
            x_base = torre['x']
            lisp_dibujar_texto(f, (x_base, alturas_niveles[min(alturas_niveles.keys())] + cfg['TORRE_LABEL_OFFSET_Y']), cfg['TORRE_LABEL_ALTURA'], torre['nombre'])

            lisp_escribir(f, f'(princ "\\n   - Dibujando dispositivos por nivel...")')
            for nivel_id, nivel_data in torre['niveles'].items():
                y_nivel = alturas_niveles[nivel_id]
                x_dispositivo = x_base + 50
                for tipo_qty, conf in cfg['DISPOSITIVOS'].items():
                    cantidad = nivel_data.get(tipo_qty, 0)
                    if cantidad > 0:
                        y_dispositivo = y_nivel + cfg['DISPOSITIVO_Y_OFFSET']
                        globals()[f"dibujar_icono_{conf['icono']}"](f, cfg, x_dispositivo, y_dispositivo)
                        lisp_dibujar_texto(f, (x_dispositivo + 25, y_dispositivo + 10), 10, f"{cantidad}x{conf['label']}")
                        x_dispositivo += cfg['DISPOSITIVO_ESPACIADO_X']
            lisp_escribir(f, '(princ "DONE.")')

            lisp_escribir(f, f'(princ "\\n   - Dibujando switches...")')
            y_switch = alturas_niveles[0] - 80
            for sw_nombre, sw_modelo in sorted(torre['switches'].items()):
                dibujar_switch(f, cfg, x_base + 50, y_switch, sw_nombre, sw_modelo)
                y_switch -= cfg['SWITCH_VERTICAL_SPACING']
            lisp_escribir(f, '(princ "DONE.")')

            if torre['nombre'] == 'MDF':
                lisp_escribir(f, f'(princ "\\n   - Dibujando UPS en MDF...")')
                dibujar_ups(f, cfg, x_base - 150, y_switch)
                lisp_escribir(f, '(princ "DONE.")')

        lisp_escribir(f, "\n; === DIBUJAR CABLES ===")
        lisp_escribir(f, '(princ "\\nDibujando cables de Fibra Optica...")')
        mdf_x = torres[0]['x'] + 50 + cfg['SWITCH_ANCHO'] / 2
        mdf_y_base = alturas_niveles[0] - 80 + cfg['SWITCH_ALTO'] / 2

        for i, torre in enumerate(torres):
            if i == 0: continue
            idf_x = torre['x'] + 50 + cfg['SWITCH_ANCHO'] / 2
            y_fibra_offset = 0
            for sw_nombre, sw_conf in cfg['SWITCH_CONFIG'].items():
                if sw_nombre in torre['switches'] and sw_nombre not in ["SW-FIREWALL", "SW-CORE"]:
                    lisp_seleccionar_capa_y_color(f, sw_conf['capa'], sw_conf['color'])
                    p1 = (mdf_x, mdf_y_base - y_fibra_offset)
                    p2 = (idf_x, mdf_y_base - y_fibra_offset)
                    lisp_dibujar_linea(f, p1, p2)
                    y_fibra_offset += 5
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
