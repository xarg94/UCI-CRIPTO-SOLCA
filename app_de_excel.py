# -*- coding: utf-8 -*-
#
# Aplicacion Flask para replicar logica de una hoja de calculo de Excel.
# Esta version utiliza Flask, resultados condicionales, boton de calculo e impresion.
#
# REQUISITOS: 'Flask', 'pandas', 'openpyxl', 'gunicorn' (para despliegue global)
# INSTRUCCION: Coloca tu archivo de Excel nombrado 'datos.xlsx' en la misma carpeta.

from flask import Flask, request, render_template_string
import pandas as pd
import json
import math
import datetime
import re 

# 1. Configuracion de la aplicacion Flask
app = Flask(__name__)
EXCEL_FILE_PATH = 'datos.xlsx'
datos_hojas = {}
error_lectura = None

# --- CONSTANTES DE CONFIGURACION ---
HOJAS_PANEL = [
    'Panel',
    'Microdinamia',
    'Macrodinamia',
    'Ventilatorio',
    'Neurocritico' 
]

BACKGROUND_IMAGES = {}

# --- Funcion para cargar el Excel ---
def cargar_datos_excel():
    """Carga todas las hojas del archivo de Excel usando pandas."""
    global datos_hojas, error_lectura
    try:
        datos_hojas = pd.read_excel(EXCEL_FILE_PATH, sheet_name=HOJAS_PANEL)
        error_lectura = None
        return True
    except FileNotFoundError:
        error_lectura = f"ERROR: El archivo '{EXCEL_FILE_PATH}' no se encontro en la carpeta."
        return False
    except ValueError as ve:
        error_lectura = f"ERROR al leer el archivo Excel: {ve}"
        return False
    except Exception as e:
        error_lectura = f"ERROR al leer el archivo Excel: {e}"
        return False

# Carga inicial de datos
cargar_datos_excel()

# Definicion del template HTML
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ICU-CRIPTOS | Monitoreo UCI</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @media print {
            body * { visibility: hidden !important; }
            #print-area, #print-area * { visibility: visible !important; }
            #print-area { 
                position: absolute; left: 0; top: 0; 
                padding: 10px; 
                margin: 0;
                font-size: 8pt; 
            }
            .print-hidden { display: none !important; }
            .bg-panel::before { content: none !important; } 
            .bg-panel { background-image: none !important; background-color: #ffffff !important; box-shadow: none !important; border: none !important;}
        }
        
        /* --- ESTILOS DE FONDO Y TRANSPARENCIA (SOLICITADOS POR EL USUARIO) --- */
        body { 
            font-family: 'Inter', sans-serif; 
            background-color: #f8fafc;
            /* IMAGEN DE FONDO PROPORCIONADA POR EL USUARIO (Imgur) */
            background-image: url('https://i.imgur.com/7zQ8s18.jpeg');
            background-size: cover; 
            background-position: center; /* Centrado */
            background-attachment: fixed;
        }
        /* Tarjeta principal con 50% de transparencia */
        .main-app-card {
             background-color: rgba(255, 255, 255, 0.5); /* Alpha 0.5 para 50% de transparencia */
             backdrop-filter: blur(5px);
        }
        
        .subtitle-italic { font-style: italic; }
        .input-base { transition: background-color 0.2s; border-radius: 0.5rem; font-size: 0.875rem; }
        .text-base { font-size: 0.875rem; }
        .text-xl { font-size: 1.125rem; }
        .text-2xl { font-size: 1.5rem; }
        .text-4xl { font-size: 2.25rem; }

        /* Estilos para Fondos y Superposiciones */
        .bg-panel {
            background-size: cover; background-position: center; position: relative; overflow: hidden;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        .bg-panel::before {
            content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
            background-color: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(1px); z-index: 1;
        }
        .bg-panel > * { position: relative; z-index: 10; }
        
        /* ESTILO DEL SEPARADOR / TÍTULO DE SECCIÓN */
        .result-separator { 
            color: #4338ca; /* Indigo 700 */
            font-size: 1.1rem; /* Un poco más grande */
            font-weight: 800; 
            margin-top: 15px; margin-bottom: 8px;
            padding-bottom: 4px; border-bottom: 2px solid #a5b4fc; 
            text-transform: uppercase;
        }
    </style>
</head>
<body class="p-4 md:p-8">
    <div class="max-w-4xl mx-auto rounded-2xl shadow-xl p-6 md:p-10 main-app-card">
        
        <h1 class="text-4xl font-extrabold text-center text-indigo-700 mb-1 print-hidden">
            Monitoreo UCI
        </h1>
        <p class="text-center text-gray-700 mb-2 text-xl subtitle-italic leading-tight print-hidden">
            ICU–CRIPTOS| Hemodynamic, Respiratory & Neurocritical Intelligence<br>
            <span class="text-gray-500 text-base">Monitoreo del paciente critico, en la palma de mi mano</span>
        </p>

        <!-- FRASE DE AUTOR MOVIDA AQUÍ -->
        <p class="text-center text-sm text-gray-600 mb-6 font-semibold print-hidden">
            Elaborada por X. Real, P. Olivos y O. Bolaños | Agradecimiento especial a G. Maldonado por su colaboración
        </p>
        
        {% if error_lectura %}
            <div class="p-4 mb-6 bg-red-100 border border-red-400 text-red-700 rounded-lg text-base print-hidden">
                <p class="font-bold">Error de Carga:</p>
                <p>{{ error_lectura }}</p>
            </div>
        {% else %}
            <form method="POST" action="/" class="p-6 border border-gray-200 rounded-xl shadow-inner bg-gray-50 mb-8 space-y-6 print-hidden" id="data-form">
                
                <!-- 1. Datos Antropometricos -->
                <div class="grid md:grid-cols-4 gap-4 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Datos Antropometricos</h2>
                    </div>
                    {% for label, name, type, placeholder, options in [
                        ("Sexo:", "sexo", "select", "H", [("H", "Hombre"), ("M", "Mujer")]),
                        ("Edad (anos):", "edad_anos", "number", "anos", None),
                        ("Peso (Kg):", "peso_kg", "number", "Kg", None),
                        ("Talla (m):", "talla_m", "number", "m", None)
                    ] %}
                    <div>
                        <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                        {% if type == 'select' %}
                        <select name="{{ name }}" id="{{ name }}" onchange="updateBackground(this)" class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500 bg-white">
                            <option value="">Selecciona</option>
                            {% for value, text in options %}
                            <option value="{{ value }}" {% if inputs.get(name) == value %}selected{% endif %}>{{ text }}</option>
                            {% endfor %}
                        </select>
                        {% else %}
                        <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                               placeholder="{{ placeholder }}" 
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500 bg-white" 
                               oninput="updateBackground(this)">
                        {% endif %}
                    </div>
                    {% endfor %}
                </div>
                
                <!-- 2. Signos Vitales -->
                <div class="grid md:grid-cols-4 gap-4 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Signos Vitales</h2>
                    </div>
                    {% for label, name, placeholder, type, options in [
                        ("TAS:", "tas", "mmHg", "number", None), 
                        ("TAD:", "tad", "mmHg", "number", None), 
                        ("FC:", "fc", "lpm", "number", None), 
                        ("SatO₂ Pulsioximetria:", "sato2_sv", "%", "number", None)
                    ] %}
                    <div>
                        <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                        <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                               placeholder="{{ placeholder }}" 
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm" 
                               oninput="updateBackground(this)">
                    </div>
                    {% endfor %}
                </div>
                
                <!-- 3. Gasometria Arterial y Venosa -->
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <!-- Gasometria Arterial -->
                    <div class="bg-gray-100 p-4 rounded-lg">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Gasometria Arterial 🩸</h2>
                        <div class="grid grid-cols-2 gap-4 text-sm">
                            {% for label, name, placeholder, type, options in [
                                ("pH:", "ph_a", "pH", "number", None), 
                                ("PaCO₂:", "paco2", "mmHg", "number", None), 
                                ("PaO₂:", "pao2", "mmHg", "number", None),
                                ("SatO₂ (a):", "sato2_a", "%", "number", None), 
                                ("Lactato:", "lactato", "mmol/L", "number", None), 
                                ("Hb (g/dL):", "hb", "g/dL", "number", None)
                            ] %}
                            <div>
                                <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                                <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                                       placeholder="{{ placeholder }}" 
                                       class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm" 
                                       oninput="updateBackground(this)">
                            </div>
                            {% endfor %}
                        </div>
                    </div>
                    
                    <!-- Gasometria Venosa -->
                    <div class="bg-gray-100 p-4 rounded-lg">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Gasometria Venosa 🔵</h2>
                        <div class="grid grid-cols-2 gap-4 text-sm">
                            {% for label, name, placeholder, type, options in [
                                ("pHv:", "ph_v", "pH", "number", None), 
                                ("PvCO₂:", "pvco2", "mmHg", "number", None), 
                                ("PvO₂:", "pvo2", "mmHg", "number", None),
                                ("SatvO₂:", "satvo2", "%", "number", None)
                            ] %}
                            <div>
                                <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                                <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                                       placeholder="{{ placeholder }}" 
                                       class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm" 
                                       oninput="updateBackground(this)">
                            </div>
                            {% endfor %}
                        </div>
                    </div>
                </div>

                <!-- 4. Macrodinamia (POCUS) -->
                <div class="grid grid-cols-5 gap-4 text-sm bg-gray-50 p-4 rounded-lg shadow-inner">
                    <div class="col-span-5">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">POCUS (Macrodinamia) 🩺</h2>
                    </div>
                    {% for label, name, placeholder, type, options in [
                        ("VTI:", "vti", "cm", "number", None),
                        ("TSVI:", "tsvi", "cm", "number", None),
                        ("VCI:", "vci", "cm", "number", None),
                        ("PVC Medido:", "pvc_medido", "mmHg", "number", None),
                        ("VCI Colaps.:", "vci_colaps", "Selecciona Colapso", "select", [("total", "Total"), (">50%", ">50%"), ("<50%", "<50%"), ("No cambios", "No cambios")])
                    ] %}
                    <div>
                        <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                        {% if type == 'select' %}
                        <select name="{{ name }}" id="{{ name }}" onchange="updateBackground(this)" class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white">
                            <option value="">{{ placeholder }}</option>
                            {% for value, text in options %}
                            <option value="{{ value }}" {% if inputs.get(name) == value %}selected{% endif %}>{{ text }}</option>
                            {% endfor %}
                        </select>
                        {% else %}
                        <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                               placeholder="{{ placeholder }}" 
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white" 
                               oninput="updateBackground(this)">
                        {% endif %}
                    </div>
                    {% endfor %}
                </div>

                <!-- 5. Hemodinamia (VI/VD) -->
                <div class="bg-gray-50 p-4 rounded-lg shadow-inner">
                    <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Hemodinamia (ECHO Av.)</h2>
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                        <!-- VI -->
                        {% for label, name, placeholder, type, options in [
                            ("MAPSE L:", "mapse_l", "cm", "number", None), 
                            ("MAPSE S:", "mapse_s", "cm", "number", None), 
                            ("E (onda):", "e_onda", "m/s", "number", None), 
                            ("A (onda):", "a_onda", "m/s", "number", None), 
                            ("E' lat:", "eprim_lat", "cm/s", "number", None), 
                            ("E' med:", "eprim_med", "cm/s", "number", None),
                            ("VFS:", "vfs", "ml", "number", None), 
                            ("VFD:", "vfd", "ml", "number", None), 
                            ("Long. VI:", "long_vi", "cm", "number", None)
                        ] %}
                        <div>
                            <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                            <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                                   placeholder="{{ placeholder }}" 
                                   class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white" 
                                   oninput="updateBackground(this)">
                        </div>
                        {% endfor %}
                        <!-- VD -->
                        {% for label, name, placeholder, type, options in [
                            ("VTmax:", "vtmax", "m/s", "number", None), 
                            ("TAPSE:", "tapse", "mm", "number", None), 
                            ("VTI Pulmonar:", "vti_pulmonar", "cm", "number", None)
                        ] %}
                        <div>
                            <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                            <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                                   placeholder="{{ placeholder }}" 
                                   class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white" 
                                   oninput="updateBackground(this)">
                        </div>
                        {% endfor %}
                    </div>
                </div>

                <!-- 6. Datos Ventilatorios -->
                <div class="bg-gray-50 p-4 rounded-lg shadow-inner">
                    <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Datos Ventilatorios 🌬️</h2>
                    <div class="grid grid-cols-3 gap-4 text-sm">
                        {% for label, name, type, placeholder, options in [
                            ("MODO:", "modo", "select", "Selecciona Modo", [("PCV", "PCV"), ("VCV", "VCV")]),
                            ("VT protec. (ml/kg):", "vt_protec", "number", "ml/kg", None),
                            ("VT Ventilador (ml):", "vt_ventilador", "number", "ml", None),
                            ("FR (lpm):", "fr", "number", "lpm", None),
                            ("PeCO₂ (mmHg):", "peco2", "number", "mmHg", None),
                            ("PEEP (cmH₂O):", "peep", "number", "cmH₂O", None),
                            ("FIO₂ (0.x):", "fio2", "number", "0.x", None),
                            ("Plateau (cmH₂O):", "plateau", "number", "cmH₂O", None),
                            ("Ppico (cmH₂O):", "ppico", "number", "cmH₂O", None),
                            ("Cstat (medida):", "cstat_input", "number", "ml/cmH₂O", None),
                            ("Cdin (medida):", "cdin_input", "number", "ml/cmH₂O", None),
                            ("V/min (L/min):", "v_min", "number", "L/min", None),
                            ("POCC (cmH₂O):", "pocc", "number", "cmH₂O", None)
                        ] %}
                        <div>
                            <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                            {% if type == 'select' %}
                            <select name="{{ name }}" id="{{ name }}" onchange="updateBackground(this)" class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white">
                                <option value="">Selecciona</option>
                                {% for value, text in options %}
                                <option value="{{ value }}" {% if inputs.get(name) == value %}selected{% endif %}>{{ text }}</option>
                                {% endfor %}
                            </select>
                            {% else %}
                            <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                                   placeholder="{{ placeholder }}" 
                                   class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white" 
                                   oninput="updateBackground(this)">
                            {% endif %}
                        </div>
                        {% endfor %}
                    </div>
                </div>

                <!-- 7. Monitorizacion Neurocritica -->
                <div class="bg-gray-50 p-4 rounded-lg shadow-inner">
                    <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Monitorizacion Neurocritica 🧠</h2>
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                        {% for label, name, placeholder, type, options in [
                            ("VS (ACM):", "vs_acm", "cm/s", "number", None), 
                            ("VD (ACM):", "vd_acm", "cm/s", "number", None),
                            ("VS (AB):", "vs_ab", "cm/s", "number", None), 
                            ("VD (AB):", "vd_ab", "cm/s", "number", None),
                            ("Arteria Medida:", "vaso_dtc", "Selecciona Arteria", "select", [("ACM", "ACM"), ("ACA", "ACA"), ("ACP", "ACP"), ("AB", "AB")]),
                            ("VS (Gen.):", "vs_dtc", "cm/s", "number", None), 
                            ("VD (Gen.):", "vd_dtc", "cm/s", "number", None),
                            ("VM Art. Carotida Int.:", "vm_aci", "cm/s", "number", None), 
                            ("VM Art. Vertebral:", "vm_ave", "cm/s", "number", None),
                            ("VNO Der. (mm):", "vno_der", "mm", "number", None), 
                            ("VNO Izq. (mm):", "vno_izq", "mm", "number", None), 
                            ("DGO (mm):", "vno_dgo", "mm", "number", None),
                            ("pH jO₂:", "ph_jo2", "pH", "number", None), 
                            ("PjCO₂:", "paco2_jo2", "mmHg", "number", None), 
                            ("PjO₂:", "pao2_jo2", "mmHg", "number", None),
                            ("SjO₂:", "sato2_jo2", "%", "number", None), 
                            ("Lactato jO₂:", "lactato_jo2", "mmol/L", "number", None)
                        ] %}
                        <div>
                            <label for="{{ name }}" class="block text-sm font-medium text-gray-700">{{ label }}</label>
                            {% if type == 'select' %}
                            <select name="{{ name }}" id="{{ name }}" onchange="updateBackground(this)" class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white">
                                <option value="">{{ placeholder }}</option>
                                {% for value, text in options %}
                                <option value="{{ value }}" {% if inputs.get(name) == value %}selected{% endif %}>{{ text }}</option>
                                {% endfor %}
                            </select>
                            {% else %}
                            <input type="number" step="any" id="{{ name }}" name="{{ name }}" value="{{ inputs.get(name) or '' }}" 
                                   placeholder="{{ placeholder }}" 
                                   class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm bg-white" 
                                   oninput="updateBackground(this)">
                            {% endif %}
                        </div>
                        {% endfor %}
                    </div>
                </div>
                
                <!-- BOTONES DE ACCION -->
                <div class="flex justify-center mt-6 space-x-4">
                    <button type="submit" name="action" value="calculate"
                            class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-6 rounded-lg shadow-lg transition duration-150 ease-in-out">
                        Mostrar Resultados
                    </button>
                    <button type="button" onclick="clearForm()"
                            class="bg-yellow-300 hover:bg-yellow-400 text-gray-800 font-bold py-2 px-6 rounded-lg shadow-lg transition duration-150 ease-in-out">
                        Limpiar
                    </button>
                </div>
            </form>

            <!-- 3. SECCION DE RESULTADOS -->
            {% if results_json and not error_calculo and show_results %}
                <div class="mt-8" id="print-area">
                    <h2 class="text-2xl font-bold text-gray-800 mb-4 print-hidden">Resultados por Seccion</h2>
                    <div class="flex justify-between text-sm text-gray-500 mb-4">
                        <span>Fecha y Hora: {{ now }}</span>
                    </div>

                    <div class="grid md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                        
                        {% set results = json.loads(results_json) %}
                        {% for panel_nombre, panel_data in results.items() %}
                            {% set bg_img = BACKGROUND_IMAGES.get(panel_nombre, '') %}
                            {% if panel_data and not panel_data.get('error') %}
                                <div class="bg-panel rounded-xl shadow-md p-5 transition duration-200 hover:shadow-lg"
                                     style="{% if bg_img %}background-image: url('{{ bg_img }}');{% endif %}">
                                    <h3 class="text-xl font-bold mb-3 text-indigo-700">{{ panel_nombre }}</h3>
                                    
                                    <div class="text-sm space-y-1">
                                        {% for key, valor in panel_data.items() %}
                                            {# CORRECCIÓN CRÍTICA: Solo el separador principal empieza con -- #}
                                            {% if key.startswith('--') %}
                                                <div class="result-separator">
                                                    {{ key | replace('--', '') | trim }}
                                                </div>
                                            {% else %}
                                                <div class="flex justify-between items-start py-1 border-b border-gray-200 last:border-b-0">
                                                    <span class="text-gray-600 font-medium w-1/2 pr-2">{{ key }}:</span>
                                                    <span class="text-gray-900 font-bold w-1/2 text-right">{{ valor | safe }}</span>
                                                </div>
                                            {% endif %}
                                        {% endfor %}
                                    </div>
                                </div>
                            {% endif %}
                        {% endfor %}
                    </div>
                    
                    <!-- Cuadro de Abreviaturas Ventilatorias -->
                    {% if 'Ventilatorio' in results and not results.Ventilatorio.get('error') %}
                        <div class="mt-8 p-5 bg-blue-50 border-l-4 border-blue-400 rounded-lg shadow-inner text-base">
                            <h4 class="text-lg font-semibold text-blue-800 mb-3">Abreviaturas Ventilatorias</h4>
                            <div class="grid grid-cols-2 sm:grid-cols-3 gap-2 text-sm text-gray-700">
                                <div><span class="font-bold">EM:</span> Espacio Muerto</div>
                                <div><span class="font-bold">EV:</span> Eficiencia Ventilatoria</div>
                                <div><span class="font-bold">Shunt:</span> Cortocircuito Intrapulmonar</div>
                                <div><span class="font-bold">PpMt:</span> Presion transpulmonar muscular</div>
                                <div><span class="font-bold">PM:</span> Poder Mecanico</div>
                                <div><span class="font-bold">Raw:</span> Resistencia de Via Aerea</div>
                            </div>
                        </div>
                    {% endif %}
                </div>
                
                <div class="flex justify-center mt-8 print-hidden">
                    <button onclick="window.print()" type="button"
                            class="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-6 rounded-lg shadow-lg transition duration-150 ease-in-out">
                        Imprimir Resultados
                    </button>
                </div>
                
            {% elif error_calculo and show_results %}
                <div class="mt-8 p-4 bg-red-100 border border-red-400 text-red-700 rounded-lg text-base print-hidden" role="alert">
                    <p class="font-bold">Error en el Calculo de Formulas</p>
                    <p class="sm">
                        {{ error_calculo }}
                        <br>
                        <span class="font-semibold text-xs text-red-600 block mt-1">
                            Asegurese de que todos los campos relevantes para el calculo tengan un valor numerico valido.
                        </span>
                    </p>
                </div>
            {% endif %}
        {% endif %}
    </div>
    
    <!-- FOOTER SOLICITADO (OCULTADO para mover el contenido) -->
    <footer class="text-center text-sm text-gray-500 mt-8 pb-4 print-hidden hidden">
        Elaborada por X. Real, P. Olivos y O. Bolaños
    </footer>

    <script>
        function updateBackground(element) {
            const isSelect = element.tagName === 'SELECT';
            let value = element.value.trim();
            element.classList.remove('bg-white', 'bg-green-100', 'bg-red-100');
            if (isSelect) {
                if (value !== '' && value !== 'Selecciona' && value !== 'Selecciona Colapso' && value !== 'Selecciona Modo' && value !== 'Selecciona Vaso' && value !== 'Selecciona Arteria') {
                    element.classList.add('bg-green-100');
                } else {
                    element.classList.add('bg-white');
                }
            } else {
                value = value.replace(',', '.');
                const isNumeric = !isNaN(parseFloat(value)) && isFinite(value);
                if (value === '') {
                    element.classList.add('bg-white');
                } else if (isNumeric) {
                    element.classList.add('bg-green-100');
                } else {
                    element.classList.add('bg-red-100');
                }
            }
        }
        document.addEventListener('DOMContentLoaded', () => {
            document.querySelectorAll('.input-base').forEach(updateBackground);
        });
        function clearForm() {
            document.querySelectorAll('#data-form input, #data-form select').forEach(element => {
                if (element.tagName === 'SELECT') {
                    element.selectedIndex = 0; 
                } else {
                    element.value = '';
                }
                updateBackground(element);
            });
            document.getElementById('data-form').submit();
        }
    </script>
</body>
</html>
"""

# 4. --- Logica de Replicacion de Formulas ---
def replicar_formulas(user_inputs):
    """
    Funcion que replica la logica de las formulas de Excel.
    Recibe la entrada dinamica del usuario.
    """
    resultados = {}

    if error_lectura or not datos_hojas:
        return {"error": "No se cargaron los datos de Excel."}, None

    # --- ENTRADAS DEL USUARIO ---
    
    try:
        def get_float(name):
            val = user_inputs.get(name)
            if val is None or val == '':
                return None
            try:
                return float(str(val).replace(',', '.'))
            except ValueError:
                return None
        
        # Antropometricos
        sexo = user_inputs.get('sexo')
        edad = get_float('edad_anos')
        peso_kg = get_float('peso_kg')
        talla_m = get_float('talla_m')
        
        # Signos Vitales
        tas = get_float('tas')
        tad = get_float('tad')
        fc = get_float('fc')
        sato2_sv = get_float('sato2_sv') # SatO₂ Pulsioximetria
        
        # Gasometria Arterial
        ph_a = get_float('ph_a')
        paco2 = get_float('paco2') # PaCO₂
        pao2 = get_float('pao2')
        sato2_a = get_float('sato2_a') # SatO₂ (a)
        lactato = get_float('lactato')
        hb = get_float('hb')
        
        # Gasometria Venosa
        ph_v = get_float('ph_v')
        pvco2 = get_float('pvco2') # PvCO₂
        pvo2 = get_float('pvo2')
        satvo2 = get_float('satvo2') # SatvO₂

        # Macrodinamia / POCUS
        vti = get_float('vti')
        tsvi = get_float('tsvi')
        vci = get_float('vci')
        vci_colaps = user_inputs.get('vci_colaps')
        pvc_medido = get_float('pvc_medido')
        
        # Hemodinamia (VI/VD)
        mapse_l = get_float('mapse_l')
        mapse_s = get_float('mapse_s')
        e_onda = get_float('e_onda')
        a_onda = get_float('a_onda')
        eprim_lat = get_float('eprim_lat')
        eprim_med = get_float('eprim_med')
        vfs = get_float('vfs') 
        vfd = get_float('vfd')
        long_vi = get_float('long_vi')
        vtmax = get_float('vtmax')
        tapse = get_float('tapse') 
        vti_pulmonar = get_float('vti_pulmonar') 

        # Ventilatorio Inputs
        modo = user_inputs.get('modo') 
        vt_protec_ml_kg = get_float('vt_protec') 
        vt_ventilador = get_float('vt_ventilador') 
        fr = get_float('fr') 
        peco2 = get_float('peco2') # PeCO₂
        peep = get_float('peep') 
        fio2 = get_float('fio2') # FiO₂
        plateau = get_float('plateau') 
        ppico = get_float('ppico') 
        cstat_input = get_float('cstat_input') 
        cdin_input = get_float('cdin_input') 
        v_min = get_float('v_min') # V/min
        pocc = get_float('pocc') 

        # Neurocritico Inputs
        vs_acm = get_float('vs_acm')
        vd_acm = get_float('vd_acm')
        vs_ab = get_float('vs_ab')
        vd_ab = get_float('vd_ab')
        vaso_dtc = user_inputs.get('vaso_dtc')
        vs_dtc = get_float('vs_dtc')
        vd_dtc = get_float('vd_dtc')
        vm_aci = get_float('vm_aci')
        vm_ave = get_float('vm_ave')
        vno_der = get_float('vno_der')
        vno_izq = get_float('vno_izq')
        vno_dgo = get_float('vno_dgo')
        ph_jo2 = get_float('ph_jo2') 
        paco2_jo2 = get_float('paco2_jo2') 
        pao2_jo2 = get_float('pao2_jo2') 
        sato2_jo2 = get_float('sato2_jo2')
        lactato_jo2 = get_float('lactato_jo2') 
        
        # Verificar si hay suficientes datos base para iniciar
        if not (peso_kg and talla_m):
             return json.dumps(resultados), None


        # --- CALCULOS INTERMEDIOS BASE (PANEL) ---
        
        # D20: TAM
        tam = (tas + (2 * tad)) / 3 if tas and tad else None
        
        # SCT (D10)
        talla_cm = talla_m * 100 if talla_m else None
        sct = (0.020247 * (peso_kg ** 0.425) * (talla_m ** 0.725)) * 10 if peso_kg and talla_m else None
        
        # D11: PI (Peso Ideal) -- FÓRMULA DE MILLER/BROCA CORREGIDA
        pi = None
        if talla_m and sexo in ["H", "M"]:
            # Fórmula exacta de Excel: SI(D5="H",56.2+1.41*((D10*100)/2.54-60),SI(D5="M",53.1+1.36*((D10*100)/2.54-60),"Error"))
            talla_pulgadas_menos_60 = (talla_m * 100 / 2.54) - 60
            if sexo == "H":
                pi = 56.2 + 1.41 * talla_pulgadas_menos_60
            elif sexo == "M":
                pi = 53.1 + 1.36 * talla_pulgadas_menos_60

        # --- CALCULOS DE MACRODINAMIA CENTRAL (POCUS) ---

        # D10: VS (Volumen Sistolico - Macrodinamia)
        vs_macro = ((tsvi ** 2) * 0.785) * vti if tsvi and vti else None
        
        # D11: GC (Gasto Cardiaco - Macrodinamia)
        gc = (vs_macro / 1000) * fc if vs_macro and fc else None
        
        # D12: IC (Indice Cardíaco)
        ic = gc / sct if sct and gc else None
        
        # D15: PVC ECO
        pvc_eco = None
        if vci and vci_colaps and vci > 0:
            if vci < 1.5: pvc_eco = 5
            elif vci >= 1.5 and vci <= 2.5:
                if vci_colaps in ["total", ">50%"]: pvc_eco = 8
                elif vci_colaps == "<50%": pvc_eco = 13
            elif vci > 2.5:
                if vci_colaps == "<50%": pvc_eco = 18
                elif vci_colaps == "No cambios": pvc_eco = 20
        elif pvc_medido is not None:
             pvc_eco = pvc_medido

        # --- CALCULOS DE MICRODINAMIA (BASE) ---
        cao2, cvo2, davo2, exto2, cco2 = None, None, None, None, None
        
        if hb and sato2_a and pao2:
            sato2_frac = sato2_a / 100.0
            cao2 = (1.36 * hb * sato2_frac) + (0.0031 * pao2)
        
        if hb and satvo2 and pvo2:
            satvo2_frac = satvo2 / 100.0
            cvo2 = (1.36 * hb * satvo2_frac) + (0.0031 * pvo2)
        
        if hb and pao2:
            cco2 = (1.36 * hb * 1.0) + (0.0031 * pao2) 
        
        if cao2 and cvo2:
            davo2 = cao2 - cvo2
        
        if davo2 and cao2 and cao2 != 0.0:
            exto2 = (davo2 / cao2) * 100 
        
        shunt_denominador = (cco2 - cvo2)
        shunt_numerador = (cco2 - cao2)
        shunt = None
        if cco2 and cao2 and cvo2 and shunt_denominador != 0.0:
            shunt_val = (shunt_numerador / shunt_denominador) * 100
            shunt = shunt_val 
        
        # *****************************************************************
        # --- INICIO DE LA LOGICA DE TUS 6 HOJAS ---
        # *****************************************************************
        
        # --- 1. PANEL ---
        panel_resultados = {}
        
        # Antropometricos
        imc = peso_kg / (talla_m ** 2) if talla_m and talla_m != 0.0 else None
        act = None
        if sexo and edad and peso_kg and talla_cm:
            if sexo == "H":
                act = 2.447 - (0.09156 * edad) + (0.3362 * peso_kg) + (0.1074 * talla_cm)
            elif sexo == "M":
                act = -2.097 + (0.1069 * talla_cm) + (0.2466 * peso_kg)

        panel_resultados['-- Datos Antropometricos --'] = " "
        if user_inputs.get('sexo'): panel_resultados['Sexo'] = sexo
        if edad is not None: panel_resultados['Edad'] = f"{edad:.0f} anos"
        if peso_kg is not None: panel_resultados['Peso'] = f"{peso_kg:.0f} Kg"
        if talla_m is not None: panel_resultados['Talla'] = f"{talla_m:.2f} m"
        if imc is not None: panel_resultados['IMC'] = f"{imc:.2f}"
        if sct is not None: panel_resultados['SCT'] = f"{sct:.2f} m²"
        if pi is not None: panel_resultados['PI'] = f"{pi:.2f} Kg"
        if act is not None: panel_resultados['ACT'] = f"{act:.2f} L"
        
        panel_resultados['-- Signos Vitales --'] = " "
        if tas is not None: panel_resultados['TAS'] = f"{tas:.0f} mmHg"
        if tad is not None: panel_resultados['TAD'] = f"{tad:.0f} mmHg"
        if tam is not None: panel_resultados['TAM'] = f"{tam:.0f} mmHg"
        if fc is not None: panel_resultados['FC'] = f"{fc:.0f} lpm"
        if sato2_sv is not None: panel_resultados['SatO₂ Pulsioximetria'] = f"{sato2_sv:.0f} %"
        
        panel_resultados['-- Gasometria Arterial 🩸 --'] = " "
        if ph_a is not None: panel_resultados['pH (a)'] = f"{ph_a:.2f}"
        if paco2 is not None: panel_resultados['PaCO₂'] = f"{paco2:.1f} mmHg"
        if pao2 is not None: panel_resultados['PaO₂'] = f"{pao2:.1f} mmHg"
        if sato2_a is not None: panel_resultados['SatO₂ (a)'] = f"{sato2_a:.1f} %"
        if lactato is not None: panel_resultados['Lactato'] = f"{lactato:.2f} mmol/L"
        if hb is not None: panel_resultados['Hb'] = f"{hb:.1f} g/dL"
        
        panel_resultados['-- Gasometria Venosa 🔵 --'] = " "
        if ph_v is not None: panel_resultados['pHv'] = f"{ph_v:.2f}"
        if pvco2 is not None: panel_resultados['PvCO₂'] = f"{pvco2:.1f} mmHg"
        if pvo2 is not None: panel_resultados['PvO₂'] = f"{pvo2:.1f} mmHg"
        if satvo2 is not None: panel_resultados['SatvO₂'] = f"{satvo2:.1f} %"
        
        resultados['Panel'] = {k: v for k, v in panel_resultados.items() if (isinstance(v, str) and v.startswith('--')) or not (isinstance(v, str) and v.startswith('Error'))}


        # --- 2. MACRODINAMIA (POCUS Central) ---
        macrodinamia_resultados = {}
        
        tsvi_inf = (0.01 * talla_cm) + 0.25 if talla_cm else None
        
        rvs = ((tam - pvc_eco) * 80) / gc if tam and pvc_eco and gc and gc != 0.0 else None
        rvsi = rvs / sct if sct and rvs and sct != 0.0 else None

        if tsvi is not None: macrodinamia_resultados['TSVI'] = f"{tsvi:.2f} cm"
        if vti is not None: macrodinamia_resultados['VTI'] = f"{vti:.2f} cm"
        if tsvi_inf is not None: macrodinamia_resultados['TSVI Inferido'] = f"{tsvi_inf:.2f} cm"
        if vs_macro is not None: macrodinamia_resultados['VS'] = f"{vs_macro:.0f} ml"
        if gc is not None: macrodinamia_resultados['GC'] = f"{gc:.2f} L/min"
        if ic is not None: macrodinamia_resultados['IC'] = f"{ic:.2f} L/min/m²"
        if vci is not None: macrodinamia_resultados['VCI'] = f"{vci:.2f} cm"
        if vci_colaps and vci_colaps != 'Selecciona Colapso': macrodinamia_resultados['VCI Colaps.'] = f"{vci_colaps}"
        if pvc_eco is not None: macrodinamia_resultados['PVC ECO'] = f"{pvc_eco:.0f} mmHg"
        if pvc_medido is not None: macrodinamia_resultados['PVC Medido'] = f"{pvc_medido:.0f} mmHg"
        if rvs is not None: macrodinamia_resultados['RVS'] = f"{rvs:.0f} dyn.s/cm⁵"
        if rvsi is not None: macrodinamia_resultados['RVSI'] = f"{rvsi:.0f} dyn.s/cm⁵/m²"

        resultados['Macrodinamia'] = {k: v for k, v in macrodinamia_resultados.items() if not (isinstance(v, str) and v.startswith('Error'))}


        # --- 3. MICRODINAMIA ---
        micro_resultados = {}
        
        vo2 = gc * davo2 * 10 if gc and davo2 else None
        vo2i = vo2 / sct if sct and vo2 else None
        do2 = (gc * cao2) * 10 if gc and cao2 else None
        do2i = do2 / sct if sct and do2 else None
        davco2 = pvco2 - paco2 if pvco2 and paco2 else None
        
        if cao2 is not None: micro_resultados['CaO₂'] = f"{cao2:.2f} ml/dL"
        if cvo2 is not None: micro_resultados['CvO₂'] = f"{cvo2:.2f} ml/dL"
        if cco2 is not None: micro_resultados['CcO₂'] = f"{cco2:.2f} ml/dL"
        if davo2 is not None: micro_resultados['DavO₂'] = f"{davo2:.2f} ml/dL"
        if vo2 is not None: micro_resultados['VO₂'] = f"{vo2:.2f} ml/min"
        if vo2i is not None: micro_resultados['VO₂I'] = f"{vo2i:.2f} ml/min/m²"
        if do2 is not None: micro_resultados['DO₂'] = f"{do2:.2f} ml/min"
        if do2i is not None: micro_resultados['DO₂I'] = f"{do2i:.2f} ml/min/m²"
        if exto2 is not None: micro_resultados['ExtO₂'] = f"{exto2:.2f} %"
        if davco2 is not None: micro_resultados['DavCO₂'] = f"{davco2:.1f} mmHg"
        if lactato is not None: micro_resultados['Lactato'] = f"{lactato:.2f} mmol/L"

        # GC Fick: Formula =((D8*100)/D5)/D8. Mapeo: D8 -> vo2, D5 -> davo2.
        gc_fick_calc = None
        if vo2 is not None and davo2 is not None and vo2 != 0.0 and davo2 != 0.0:
            # Implementando la formula literal de Excel: (VO2 * 100 / DavO2) / VO2
            raw_gc_fick = ((vo2 * 100) / davo2) / vo2 
            gc_fick_calc = raw_gc_fick / 10 # MODIFICACIÓN SOLICITADA
        
        # Inserción de resultados de Microdinamia (incluyendo SatvO2 y GC Fick)
        micro_resultados = {}
        if cao2 is not None: micro_resultados['CaO₂'] = f"{cao2:.2f} ml/dL"
        if cvo2 is not None: micro_resultados['CvO₂'] = f"{cvo2:.2f} ml/dL"
        if cco2 is not None: micro_resultados['CcO₂'] = f"{cco2:.2f} ml/dL"
        if davo2 is not None: micro_resultados['DavO₂'] = f"{davo2:.2f} ml/dL"
        if vo2 is not None: micro_resultados['VO₂'] = f"{vo2:.2f} ml/min"
        if vo2i is not None: micro_resultados['VO₂I'] = f"{vo2i:.2f} ml/min/m²"
        if do2 is not None: micro_resultados['DO₂'] = f"{do2:.2f} ml/min"
        if do2i is not None: micro_resultados['DO₂I'] = f"{do2i:.2f} ml/min/m²"
        if exto2 is not None: micro_resultados['ExtO₂'] = f"{exto2:.2f} %"
        if davco2 is not None: micro_resultados['DavCO₂'] = f"{davco2:.1f} mmHg"
        if lactato is not None: micro_resultados['Lactato'] = f"{lactato:.2f} mmol/L"
        if satvo2 is not None: micro_resultados['SatvO₂'] = f"{satvo2:.1f} %" 
        if gc_fick_calc is not None: micro_resultados['GC Fick'] = f"{gc_fick_calc:.2f} L/min" 
        
        resultados['Microdinamia'] = {k: v for k, v in micro_resultados.items() if not (isinstance(v, str) and v.startswith('Error'))}


        # --- 4. HEMODINAMIA (VI/VD) ---
        hemodinamia_resultados = {}
        
        # VI Calculos
        eprim_prom = (eprim_lat + eprim_med) / 2 if eprim_lat and eprim_med else None
        e_eprim = e_onda / eprim_prom if e_onda and eprim_prom and eprim_prom != 0.0 else None
        fevi_simp = ((vfd - vfs) / vfd) * 100 if vfd and vfs and vfd != 0.0 else None
        
        # Strain MAPSE: Se aplica la negación al resultado final para que sea negativo.
        strain_mapse = None
        if mapse_l and mapse_s and long_vi and long_vi != 0.0:
            raw_strain = ((mapse_l + mapse_s) / 2) / long_vi * 100
            strain_mapse = -raw_strain # MODIFICACIÓN SOLICITADA
        
        # Ea = (0.9 * TAS) / VS_Macro
        ea = (0.9 * tas) / vs_macro if tas and vs_macro and vs_macro != 0.0 else None
        
        # Ee = (0.9 * TAS) / VFS
        ee = (0.9 * tas) / vfs if tas and vfs and vfs != 0.0 else None
        
        # AVA = Ea / Ee
        ava = ea / ee if ea and ee and ee != 0.0 else None
        
        power_c = (tam * gc) / 451 if tam and gc else None
        
        hemodinamia_resultados['-- Ventriculo Izquierdo --'] = " "
        if mapse_l is not None: hemodinamia_resultados['MAPSE L'] = f"{mapse_l:.2f} cm" 
        if mapse_s is not None: hemodinamia_resultados['MAPSE S'] = f"{mapse_s:.2f} cm" 
        if e_onda is not None: hemodinamia_resultados['E'] = f"{e_onda:.2f} m/s" 
        if a_onda is not None: hemodinamia_resultados['A'] = f"{a_onda:.2f} m/s" 
        if e_onda and a_onda and a_onda != 0.0: hemodinamia_resultados['E/A'] = f"{e_onda/a_onda:.2f}" 
        if eprim_lat is not None: hemodinamia_resultados['E\' lat'] = f"{eprim_lat:.2f} cm/s" 
        if eprim_med is not None: hemodinamia_resultados['E\' med'] = f"{eprim_med:.2f} cm/s" 
        if eprim_prom is not None: hemodinamia_resultados['E\' Prom'] = f"{eprim_prom:.2f} cm/s" 
        if e_eprim is not None: hemodinamia_resultados['E/E\''] = f"{e_eprim:.2f}" 
        if vfs is not None: hemodinamia_resultados['VFS'] = f"{vfs:.0f} ml" 
        if vfd is not None: hemodinamia_resultados['VFD'] = f"{vfd:.0f} ml" 
        if fevi_simp is not None: hemodinamia_resultados['FEVI SIMP'] = f"{fevi_simp:.1f} %" 
        if long_vi is not None: hemodinamia_resultados['Long. VI'] = f"{long_vi:.1f} cm" 
        if strain_mapse is not None: hemodinamia_resultados['Strain MAPSE'] = f"{strain_mapse:.2f} %" # Ya es negativo
        if ea is not None: hemodinamia_resultados['Ea'] = f"{ea:.2f} mmHg/ml" 
        if ee is not None: hemodinamia_resultados['Ee'] = f"{ee:.2f} mmHg/ml" 
        if ava is not None: hemodinamia_resultados['AVA'] = f"{ava:.2f}" 
        if power_c is not None: hemodinamia_resultados['Power C'] = f"{power_c:.2f} W" 
        
        # VD Calculos
        welch = (e_eprim * 1.24) + 1.9 if e_eprim else None
        gradiente_it = 4 * (vtmax ** 2) if vtmax else None
        psap = gradiente_it + pvc_eco if gradiente_it and pvc_eco else None 
        pmap = ((0.6 * psap) + 2) if psap else None
        rvs_pulm = (((vtmax / vti_pulmonar) * 10) + 0.16) if vtmax and vti_pulmonar and vti_pulmonar != 0.0 else None
        rvs_pulm_in = (((pmap - welch) / ic) * 80) if pmap and welch and ic and ic != 0.0 else None
        avd = tapse / psap if tapse and psap and psap != 0.0 else None

        hemodinamia_resultados['-- Ventriculo Derecho --'] = " "
        if welch is not None: hemodinamia_resultados['Welch'] = f"{welch:.2f}" 
        if vtmax is not None: hemodinamia_resultados['VTmax'] = f"{vtmax:.2f} m/s" 
        if gradiente_it is not None: hemodinamia_resultados['Gradiente IT'] = f"{gradiente_it:.2f} mmHg" 
        if tapse is not None: hemodinamia_resultados['TAPSE'] = f"{tapse:.2f} mm" 
        if vti_pulmonar is not None: hemodinamia_resultados['VTI Pulmonar'] = f"{vti_pulmonar:.2f} cm" 
        if psap is not None: hemodinamia_resultados['PSAP'] = f"{psap:.2f} mmHg" 
        if pmap is not None: hemodinamia_resultados['PMAP'] = f"{pmap:.2f} mmHg" 
        if rvs_pulm is not None: hemodinamia_resultados['RVSPulm.'] = f"{rvs_pulm:.2f} UW" # MODIFICACIÓN SOLICITADA
        if rvs_pulm_in is not None: hemodinamia_resultados['RVSPulm. In.'] = f"{rvs_pulm_in:.2f} Dynas/m²" # MODIFICACIÓN SOLICITADA
        if avd is not None: hemodinamia_resultados['AVD'] = f"{avd:.2f}" 
        
        resultados['Hemodinamia'] = {k: v for k, v in hemodinamia_resultados.items() if (isinstance(v, str) and v.startswith('--')) or not (isinstance(v, str) and v.startswith('Error'))}


        # --- 5. VENTILATORIO ---
        ventilatorio_resultados = {}
        
        # Fórmulas de cálculo condicional (Raw, Cstat Calc, etc.)
        
        # Cálculo de PI con la fórmula exacta de Excel
        pi = None
        if talla_m and sexo in ["H", "M"]:
            talla_pulgadas_menos_60 = (talla_m * 100 / 2.54) - 60
            if sexo == "H":
                pi = 56.2 + 1.41 * talla_pulgadas_menos_60
            elif sexo == "M":
                pi = 53.1 + 1.36 * talla_pulgadas_menos_60
        
        peso_sdra = pi 
        vt_protec_calc = vt_protec_ml_kg * peso_sdra if vt_protec_ml_kg and peso_sdra else None
        
        # Driving Pressure Calculation
        driving_p = plateau - peep if plateau and peep else None 
        
        # Compliance Static Calculation (Requires driving_p != 0)
        cstat_calc = None
        if vt_ventilador is not None and driving_p is not None and driving_p != 0.0:
            cstat_calc = vt_ventilador / driving_p
            
        ppico_menos_peep = ppico - peep if ppico and peep else None
        
        # Compliance Dynamic Calculation (Requires ppico_menos_peep != 0)
        cdin_calc = None
        if vt_ventilador is not None and ppico_menos_peep is not None and ppico_menos_peep != 0.0:
            cdin_calc = vt_ventilador / ppico_menos_peep
            
        raw = ppico - plateau if ppico and plateau else None

        # --- Inserción de Inputs y Cálculos Básicos ---
        
        if modo and modo != 'Selecciona Modo': ventilatorio_resultados['MODO'] = modo
        if peso_sdra is not None: ventilatorio_resultados['Peso SDRA (PI)'] = f"{peso_sdra:.2f} Kg"
        if vt_protec_ml_kg is not None: ventilatorio_resultados['VT protec.'] = f"{vt_protec_ml_kg:.1f} ml/Kg"
        if vt_protec_calc is not None: ventilatorio_resultados['VT protec. C.'] = f"{vt_protec_calc:.0f} ml"
        if vt_ventilador is not None: ventilatorio_resultados['VT Ventilador'] = f"{vt_ventilador:.0f} ml"
        if fr is not None: ventilatorio_resultados['FR'] = f"{fr:.0f} lpm"
        if paco2 is not None: ventilatorio_resultados['PaCO₂'] = f"{paco2:.1f} mmHg"
        if peco2 is not None: ventilatorio_resultados['PeCO₂'] = f"{peco2:.1f} mmHg"
        if peep is not None: ventilatorio_resultados['PEEP'] = f"{peep:.0f} cmH₂O"
        if fio2 is not None: ventilatorio_resultados['FIO₂'] = f"{fio2:.2f}" 
        if plateau is not None: ventilatorio_resultados['Plateau'] = f"{plateau:.0f} cmH₂O"
        if driving_p is not None: ventilatorio_resultados['Driving P.'] = f"{driving_p:.0f} cmH₂O"
        if ppico is not None: ventilatorio_resultados['Ppico'] = f"{ppico:.0f} cmH₂O"
        if cstat_input is not None: ventilatorio_resultados['Cstat (medida)'] = f"{cstat_input:.1f} ml/cmH₂O"
        if cstat_calc is not None: ventilatorio_resultados['Cstat Calc'] = f"{cstat_calc:.1f} ml/cmH₂O"
        if cdin_input is not None: ventilatorio_resultados['Cdin (medida)'] = f"{cdin_input:.1f} ml/cmH₂O"
        if cdin_calc is not None: ventilatorio_resultados['Cdin Calc'] = f"{cdin_calc:.1f} ml/cmH₂O"
        if raw is not None: ventilatorio_resultados['Raw'] = f"{raw:.1f} cmH₂O/L/s"
        if v_min is not None: ventilatorio_resultados['V/min'] = f"{v_min:.1f} L/min"
        if pocc is not None: ventilatorio_resultados['POCC'] = f"{pocc:.1f} cmH₂O"

        # --- FÓRMULAS DE MONITOREO VENTILATORIO ---
        
        # 1. EM (Espacio Muerto Fisiológico)
        em_calc = None
        if paco2 is not None and peco2 is not None and paco2 != 0.0:
            em_numerador = (paco2 - peco2)
            em_calc = (em_numerador / paco2) * 100 
        if em_calc is not None: ventilatorio_resultados['EM'] = f"{em_calc:.2f} %" 
        
        # 2. EV (Eficiencia Ventilatoria)
        ev_calc = None
        if pi is not None and paco2 is not None and v_min is not None:
            ev_denominador = ((pi / 10) * 37.5)
            if ev_denominador != 0.0:
                ev_numerador = (paco2 * v_min)
                ev_calc = ev_numerador / ev_denominador
        if ev_calc is not None: ventilatorio_resultados['EV'] = f"{ev_calc:.2f}" 

        # 3. Shunt
        shunt_calc = None
        shunt_denominador = (cco2 - cvo2)
        if cco2 is not None and cao2 is not None and shunt_denominador is not None and shunt_denominador != 0.0:
            shunt_numerador = (cco2 - cao2)
            shunt_calc = ((shunt_numerador / shunt_denominador) * 100)
        if shunt_calc is not None: ventilatorio_resultados['Shunt'] = f"{shunt_calc:.2f} %"

        # 4. PM (Power Mechanical)
        pm_calc = None
        if vt_ventilador is not None and fr is not None:
            C1 = 0.098 * fr * (vt_ventilador / 1000)
            
            if modo == "VCV" and ppico is not None and driving_p is not None:
                term_p = ppico - (driving_p / 2) 
                pm_calc = C1 * term_p
            
            elif modo == "PCV" and driving_p is not None and peep is not None:
                term_p = driving_p + peep
                pm_calc = C1 * term_p
        if pm_calc is not None: ventilatorio_resultados['PM'] = f"{pm_calc:.2f} J/min"
        
        # 5. PpMt
        ppmt_calc = None
        if ppico is not None and peep is not None and pocc is not None and pocc != 0.0:
            ppmt_numerador = (ppico - peep) - 2
            ppmt_denominador = 3 * pocc 
            if ppmt_denominador != 0.0:
                ppmt_calc = ppmt_numerador / ppmt_denominador
        if ppmt_calc is not None: ventilatorio_resultados['PpMt'] = f"{ppmt_calc:.2f}"


        resultados['Ventilatorio'] = {k: v for k, v in ventilatorio_resultados.items() if not (isinstance(v, str) and v.startswith('Error'))}


        # --- 3. MICRODINAMIA ---
        micro_resultados = {}
        
        vo2 = gc * davo2 * 10 if gc and davo2 else None
        vo2i = vo2 / sct if sct and vo2 else None
        do2 = (gc * cao2) * 10 if gc and cao2 else None
        do2i = do2 / sct if sct and do2 else None
        davco2 = pvco2 - paco2 if pvco2 and paco2 else None
        
        if cao2 is not None: micro_resultados['CaO₂'] = f"{cao2:.2f} ml/dL"
        if cvo2 is not None: micro_resultados['CvO₂'] = f"{cvo2:.2f} ml/dL"
        if cco2 is not None: micro_resultados['CcO₂'] = f"{cco2:.2f} ml/dL"
        if davo2 is not None: micro_resultados['DavO₂'] = f"{davo2:.2f} ml/dL"
        if vo2 is not None: micro_resultados['VO₂'] = f"{vo2:.2f} ml/min"
        if vo2i is not None: micro_resultados['VO₂I'] = f"{vo2i:.2f} ml/min/m²"
        if do2 is not None: micro_resultados['DO₂'] = f"{do2:.2f} ml/min"
        if do2i is not None: micro_resultados['DO₂I'] = f"{do2i:.2f} ml/min/m²"
        if exto2 is not None: micro_resultados['ExtO₂'] = f"{exto2:.2f} %"
        if davco2 is not None: micro_resultados['DavCO₂'] = f"{davco2:.1f} mmHg"
        if lactato is not None: micro_resultados['Lactato'] = f"{lactato:.2f} mmol/L"

        # GC Fick: Formula =((D8*100)/D5)/D8. Mapeo: D8 -> vo2, D5 -> davo2.
        gc_fick_calc = None
        if vo2 is not None and davo2 is not None and vo2 != 0.0 and davo2 != 0.0:
            # Implementando la formula literal de Excel: (VO2 * 100 / DavO2) / VO2
            raw_gc_fick = ((vo2 * 100) / davo2) / vo2 
            gc_fick_calc = raw_gc_fick / 10 # MODIFICACIÓN SOLICITADA
        
        # Inserción de resultados de Microdinamia (incluyendo SatvO2 y GC Fick)
        micro_resultados = {}
        if cao2 is not None: micro_resultados['CaO₂'] = f"{cao2:.2f} ml/dL"
        if cvo2 is not None: micro_resultados['CvO₂'] = f"{cvo2:.2f} ml/dL"
        if cco2 is not None: micro_resultados['CcO₂'] = f"{cco2:.2f} ml/dL"
        if davo2 is not None: micro_resultados['DavO₂'] = f"{davo2:.2f} ml/dL"
        if vo2 is not None: micro_resultados['VO₂'] = f"{vo2:.2f} ml/min"
        if vo2i is not None: micro_resultados['VO₂I'] = f"{vo2i:.2f} ml/min/m²"
        if do2 is not None: micro_resultados['DO₂'] = f"{do2:.2f} ml/min"
        if do2i is not None: micro_resultados['DO₂I'] = f"{do2i:.2f} ml/min/m²"
        if exto2 is not None: micro_resultados['ExtO₂'] = f"{exto2:.2f} %"
        if davco2 is not None: micro_resultados['DavCO₂'] = f"{davco2:.1f} mmHg"
        if lactato is not None: micro_resultados['Lactato'] = f"{lactato:.2f} mmol/L"
        if satvo2 is not None: micro_resultados['SatvO₂'] = f"{satvo2:.1f} %" 
        if gc_fick_calc is not None: micro_resultados['GC Fick'] = f"{gc_fick_calc:.2f} L/min" 
        
        resultados['Microdinamia'] = {k: v for k, v in micro_resultados.items() if not (isinstance(v, str) and v.startswith('Error'))}


        # --- 4. HEMODINAMIA (VI/VD) ---
        hemodinamia_resultados = {}
        
        # VI Calculos
        eprim_prom = (eprim_lat + eprim_med) / 2 if eprim_lat and eprim_med else None
        e_eprim = e_onda / eprim_prom if e_onda and eprim_prom and eprim_prom != 0.0 else None
        fevi_simp = ((vfd - vfs) / vfd) * 100 if vfd and vfs and vfd != 0.0 else None
        
        # Strain MAPSE: Se aplica la negación al resultado final para que sea negativo.
        strain_mapse = None
        if mapse_l and mapse_s and long_vi and long_vi != 0.0:
            raw_strain = ((mapse_l + mapse_s) / 2) / long_vi * 100
            strain_mapse = -raw_strain # MODIFICACIÓN SOLICITADA
        
        # Ea = (0.9 * TAS) / VS_Macro
        ea = (0.9 * tas) / vs_macro if tas and vs_macro and vs_macro != 0.0 else None
        
        # Ee = (0.9 * TAS) / VFS
        ee = (0.9 * tas) / vfs if tas and vfs and vfs != 0.0 else None
        
        # AVA = Ea / Ee
        ava = ea / ee if ea and ee and ee != 0.0 else None
        
        power_c = (tam * gc) / 451 if tam and gc else None
        
        hemodinamia_resultados['-- Ventriculo Izquierdo --'] = " "
        if mapse_l is not None: hemodinamia_resultados['MAPSE L'] = f"{mapse_l:.2f} cm" 
        if mapse_s is not None: hemodinamia_resultados['MAPSE S'] = f"{mapse_s:.2f} cm" 
        if e_onda is not None: hemodinamia_resultados['E'] = f"{e_onda:.2f} m/s" 
        if a_onda is not None: hemodinamia_resultados['A'] = f"{a_onda:.2f} m/s" 
        if e_onda and a_onda and a_onda != 0.0: hemodinamia_resultados['E/A'] = f"{e_onda/a_onda:.2f}" 
        if eprim_lat is not None: hemodinamia_resultados['E\' lat'] = f"{eprim_lat:.2f} cm/s" 
        if eprim_med is not None: hemodinamia_resultados['E\' med'] = f"{eprim_med:.2f} cm/s" 
        if eprim_prom is not None: hemodinamia_resultados['E\' Prom'] = f"{eprim_prom:.2f} cm/s" 
        if e_eprim is not None: hemodinamia_resultados['E/E\''] = f"{e_eprim:.2f}" 
        if vfs is not None: hemodinamia_resultados['VFS'] = f"{vfs:.0f} ml" 
        if vfd is not None: hemodinamia_resultados['VFD'] = f"{vfd:.0f} ml" 
        if fevi_simp is not None: hemodinamia_resultados['FEVI SIMP'] = f"{fevi_simp:.1f} %" 
        if long_vi is not None: hemodinamia_resultados['Long. VI'] = f"{long_vi:.1f} cm" 
        if strain_mapse is not None: hemodinamia_resultados['Strain MAPSE'] = f"{strain_mapse:.2f} %" # Ya es negativo
        if ea is not None: hemodinamia_resultados['Ea'] = f"{ea:.2f} mmHg/ml" 
        if ee is not None: hemodinamia_resultados['Ee'] = f"{ee:.2f} mmHg/ml" 
        if ava is not None: hemodinamia_resultados['AVA'] = f"{ava:.2f}" 
        if power_c is not None: hemodinamia_resultados['Power C'] = f"{power_c:.2f} W" 
        
        # VD Calculos
        welch = (e_eprim * 1.24) + 1.9 if e_eprim else None
        gradiente_it = 4 * (vtmax ** 2) if vtmax else None
        psap = gradiente_it + pvc_eco if gradiente_it and pvc_eco else None 
        pmap = ((0.6 * psap) + 2) if psap else None
        rvs_pulm = (((vtmax / vti_pulmonar) * 10) + 0.16) if vtmax and vti_pulmonar and vti_pulmonar != 0.0 else None
        rvs_pulm_in = (((pmap - welch) / ic) * 80) if pmap and welch and ic and ic != 0.0 else None
        avd = tapse / psap if tapse and psap and psap != 0.0 else None

        hemodinamia_resultados['-- Ventriculo Derecho --'] = " "
        if welch is not None: hemodinamia_resultados['Welch'] = f"{welch:.2f}" 
        if vtmax is not None: hemodinamia_resultados['VTmax'] = f"{vtmax:.2f} m/s" 
        if gradiente_it is not None: hemodinamia_resultados['Gradiente IT'] = f"{gradiente_it:.2f} mmHg" 
        if tapse is not None: hemodinamia_resultados['TAPSE'] = f"{tapse:.2f} mm" 
        if vti_pulmonar is not None: hemodinamia_resultados['VTI Pulmonar'] = f"{vti_pulmonar:.2f} cm" 
        if psap is not None: hemodinamia_resultados['PSAP'] = f"{psap:.2f} mmHg" 
        if pmap is not None: hemodinamia_resultados['PMAP'] = f"{pmap:.2f} mmHg" 
        if rvs_pulm is not None: hemodinamia_resultados['RVSPulm.'] = f"{rvs_pulm:.2f} UW" # MODIFICACIÓN SOLICITADA
        if rvs_pulm_in is not None: hemodinamia_resultados['RVSPulm. In.'] = f"{rvs_pulm_in:.2f} Dynas/m²" # MODIFICACIÓN SOLICITADA
        if avd is not None: hemodinamia_resultados['AVD'] = f"{avd:.2f}" 
        
        resultados['Hemodinamia'] = {k: v for k, v in hemodinamia_resultados.items() if (isinstance(v, str) and v.startswith('--')) or not (isinstance(v, str) and v.startswith('Error'))}


        # --- 6. NEUROCRÍTICO (ORDEN ESTRICTO APLICADO) ---
        neuro_resultados = {}
        
        # Calculos Previos
        vm_acm = (vs_acm + (2 * vd_acm)) / 3 if vs_acm and vd_acm else None
        ip_acm = (vs_acm - vd_acm) / vm_acm if vs_acm and vd_acm and vm_acm else None
        ir_acm = (vs_acm - vd_acm) / vs_acm if vs_acm and vd_acm and vs_acm != 0.0 else None
        pic = (10.93 * ip_acm) - 1.28 if ip_acm else None
        ppc = tam - pic if tam and pic else None
        
        vm_ab = (vs_ab + (2 * vd_ab)) / 3 if vs_ab and vd_ab else None
        ip_ab = (vs_ab - vd_ab) / vm_ab if vs_ab and vd_ab and vm_ab else None
        ir_ab = (vs_ab - vd_ab) / vs_ab if vs_ab and vd_ab and vs_ab != 0.0 else None

        vm_dtc = (vs_dtc + (2 * vd_dtc)) / 3 if vs_dtc and vd_dtc else None
        ip_dtc = (vs_dtc - vd_dtc) / vm_dtc if vs_dtc and vd_dtc and vm_dtc else None
        ir_dtc = (vs_dtc - vd_dtc) / vs_dtc if vs_dtc and vd_dtc and vs_dtc != 0.0 else None

        il = vm_acm / vm_aci if vm_acm and vm_aci and vm_aci != 0.0 else None
        isou = vm_ab / vm_ave if vm_ab and vm_ave and vm_ave != 0.0 else None
        vno_dgo_calc = (vno_der + vno_izq) / (2 * vno_dgo) if vno_der and vno_izq and vno_dgo and vno_dgo != 0.0 else None

        sjo2_calc = sato2_jo2
        
        # CÁLCULOS NEUROCRÍTICOS
        cvjo2 = (1.36 * hb * (sjo2_calc / 100.0)) + (0.0031 * pao2_jo2) if hb and sjo2_calc and pao2_jo2 else None
        avdo2_calc = cao2 - cvjo2 if cao2 and cvjo2 else None
        ceo2_calc = (avdo2_calc / cao2) * 100 if avdo2_calc and cao2 and cao2 != 0.0 else None


        # --- SECCION 1: DTC (ACM) - ORDEN FORZADO ---
        neuro_resultados['-- DTC (ACM) --'] = " "
        
        if vs_acm is not None: neuro_resultados['VS (ACM)'] = f"{vs_acm:.1f} cm/s"
        if vd_acm is not None: neuro_resultados['VD (ACM)'] = f"{vd_acm:.1f} cm/s"
        if vm_acm is not None: neuro_resultados['VM (ACM)'] = f"{vm_acm:.1f} cm/s"
        if ip_acm is not None: neuro_resultados['IP (ACM)'] = f"{ip_acm:.2f}"
        if ir_acm is not None: neuro_resultados['IR (ACM)'] = f"{ir_acm:.2f}"
        if pic is not None: neuro_resultados['PIC (Calc.)'] = f"{pic:.1f} mmHg"
        if ppc is not None: neuro_resultados['PPC (Calc.)'] = f"{ppc:.1f} mmHg"

        # --- SECCION 2: DTC (AB) ---
        neuro_resultados['-- DTC (AB) --'] = " "
        if vs_ab is not None: neuro_resultados['VS (AB)'] = f"{vs_ab:.1f} cm/s"
        if vd_ab is not None: neuro_resultados['VD (AB)'] = f"{vd_ab:.1f} cm/s" 
        if vm_ab is not None: neuro_resultados['VM (AB)'] = f"{vm_ab:.1f} cm/s"
        if ip_ab is not None: neuro_resultados['IP (AB)'] = f"{abs(ip_ab):.2f}" 
        if ir_ab is not None: neuro_resultados['IR (AB)'] = f"{abs(ir_ab):.2f}" 

        # --- SECCION 3: DTC (GENERICO) ---
        neuro_resultados['-- DTC (Genérico) --'] = " "
        
        if user_inputs.get('vaso_dtc') and user_inputs.get('vaso_dtc') != 'Selecciona Arteria': 
            neuro_resultados['Arteria Medida'] = vaso_dtc
            
        if vs_dtc is not None: neuro_resultados['VS'] = f"{vs_dtc:.1f} cm/s"
        if vd_dtc is not None: neuro_resultados['VD'] = f"{vd_dtc:.1f} cm/s"
        if vm_dtc is not None: neuro_resultados['VM'] = f"{vm_dtc:.1f} cm/s"
        if ip_dtc is not None: neuro_resultados['IP'] = f"{ip_dtc:.2f}"
        if ir_dtc is not None: neuro_resultados['IR (DTc)'] = f"{ir_dtc:.2f}"

        # --- SECCION 4: FLUJO VASCULAR ---
        neuro_resultados['-- Flujo Vascular Extracraneal --'] = " "
        if vm_aci is not None: neuro_resultados['VM Art. Carótida Int.'] = f"{vm_aci:.1f} cm/s"
        if vm_ave is not None: neuro_resultados['VM Art. Vertebral'] = f"{vm_ave:.1f} cm/s"

        neuro_resultados['-- Indices Combinados --'] = " "
        if il is not None: neuro_resultados['Indice Lindergard'] = f"{il:.2f}"
        if isou is not None: neuro_resultados['Indice de Soustiel'] = f"{isou:.2f}"
        
        # --- SECCION 5: VNO ---
        neuro_resultados['-- VNO (Vaina Nervio Optico) --'] = " "
        if vno_der is not None: neuro_resultados['Der.'] = f"{vno_der:.1f} mm"
        if vno_izq is not None: neuro_resultados['Izq.'] = f"{vno_izq:.1f} mm"
        if vno_dgo is not None: neuro_resultados['DGO'] = f"{vno_dgo:.1f} mm"
        if vno_dgo_calc is not None: neuro_resultados['VNO/DGO'] = f"{vno_dgo_calc:.2f}"

        # --- SECCION 6: GASOMETRÍA YUGULAR (Título y Contenido Corregidos) ---
        neuro_resultados['-- Gasometria yugular (jO₂) --'] = " " 
        if ph_jo2 is not None: neuro_resultados['pH'] = f"{ph_jo2:.2f}"
        if paco2_jo2 is not None: neuro_resultados['PjCO₂'] = f"{paco2_jo2:.1f} mmHg"
        if pao2_jo2 is not None: neuro_resultados['PjO₂'] = f"{pao2_jo2:.1f} mmHg"
        if sato2_jo2 is not None: neuro_resultados['SjO₂'] = f"{sato2_jo2:.1f} %"
        if lactato_jo2 is not None: neuro_resultados['Lactato'] = f"{lactato_jo2:.2f} mmol/L"

        # TÍTULO CORREGIDO Y RESULTADOS DE NEUROMONITOREO
        neuro_resultados['-- Neuro / Golfo Yugular --'] = " "
        
        if sjo2_calc is not None: neuro_resultados['SjO₂ (Monit.)'] = f"{sjo2_calc:.1f} %"
        
        # AVDO₂ CORREGIDO
        if avdo2_calc is not None: neuro_resultados['AVDO₂'] = f"{avdo2_calc:.2f}"
        
        # CEO₂ CORREGIDO
        if ceo2_calc is not None: neuro_resultados['CEO₂'] = f"{ceo2_calc:.2f} %"
        
        # CvJO₂ ELIMINADO

        resultados['Neurocritico'] = {k: v for k, v in neuro_resultados.items() if (isinstance(v, str) and v.startswith('--')) or not (isinstance(v, str) and v.startswith('Error'))}
        
        return json.dumps(resultados), None

    except ZeroDivisionError:
        return None, "Error de Division por Cero. Revisa los campos de PaCO2, CcO2, Peso SDRA, o POCC."
    except Exception as e:
        # Esto capturará cualquier error inesperado en la función y lo mostrará al usuario.
        return None, f"Error inesperado durante el calculo: {e.__class__.__name__}: {e}"
        
# 5. Ruta principal de Flask
@app.route('/', methods=['GET', 'POST'])
def inicio():
    """Maneja las solicitudes GET (mostrar formulario) y POST (calcular)."""
    
    error_calculo = None
    results_json = None
    show_results = False
    
    # Aseguramos que el valor de inicio para el select sea 'Selecciona Arteria'
    user_inputs = {
        'sexo': 'H', 'edad_anos': '', 'peso_kg': '', 'talla_m': '',
        'tas': '', 'tad': '', 'fc': '', 'sato2_sv': '',
        'ph_a': '', 'paco2': '', 'pao2': '', 'sato2_a': '', 'lactato': '', 'hb': '',
        'ph_v': '', 'pvco2': '', 'pvo2': '', 'satvo2': '',
        'vti': '', 'tsvi': '', 'vci': '', 'vci_colaps': 'Selecciona Colapso', 'pvc_medido': '',
        'mapse_l': '', 'mapse_s': '', 'e_onda': '', 'a_onda': '', 'eprim_lat': '', 'eprim_med': '',
        'vfs': '', 'vfd': '', 'long_vi': '',
        'vtmax': '', 'tapse': '', 'vti_pulmonar': '',
        'modo': 'Selecciona Modo', 'vt_protec': '', 'vt_ventilador': '', 'fr': '', 'peco2': '', 'peep': '',
        'fio2': '', 'plateau': '', 'ppico': '', 'cstat_input': '', 'cdin_input': '',
        'v_min': '', 'pocc': '',
        'vs_acm': '', 'vd_acm': '', 'vs_ab': '', 'vd_ab': '', 'vaso_dtc': 'Selecciona Arteria', 'vs_dtc': '', 'vd_dtc': '',
        'vm_aci': '', 'vm_ave': '', 'vno_der': '', 'vno_izq': '', 'vno_dgo': '',
        'ph_jo2': '', 'paco2_jo2': '', 'pao2_jo2': '', 'sato2_jo2': '', 'lactato_jo2': '', 'pvo2_jo2': '' 
    }
    
    if request.method == 'POST':
        if request.form.get('action') == 'calculate':
            show_results = True
        
        for key in user_inputs.keys():
            val = request.form.get(key)
            if val is not None:
                user_inputs[key] = val 
                
        results_json, error_calculo = replicar_formulas(user_inputs)

    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    return render_template_string(
        HTML_TEMPLATE, 
        error_lectura=error_lectura,
        results_json=results_json,
        error_calculo=error_calculo,
        inputs=user_inputs,
        show_results=show_results,
        now=now,
        json=json, 
        BACKGROUND_IMAGES=BACKGROUND_IMAGES 
    )

if __name__ == '__main__':
    HTML_TEMPLATE = re.sub(r'[\s\n\t]+"""$', '"""', HTML_TEMPLATE)
    app.run(debug=True, host='0.0.0.0', port=5002)