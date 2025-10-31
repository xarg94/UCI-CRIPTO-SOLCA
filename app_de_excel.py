# -*- coding: utf-8 -*-
#
# Aplicación Flask para replicar lógica de una hoja de cálculo de Excel.
# Requiere: 'Flask', 'pandas', 'openpyxl'
#
# INSTRUCCIÓN: Coloca tu archivo de Excel nombrado 'datos.xlsx' en la misma carpeta.

from flask import Flask, request, render_template_string, jsonify
import pandas as pd
import json 
import math # Importamos math para la función isfinite
import re # Para limpieza de cadenas HTML

# 1. Configuración de la aplicación Flask
app = Flask(__name__)
EXCEL_FILE_PATH = 'datos.xlsx'
# Usaremos un diccionario global para guardar los datos cargados de Excel
datos_hojas = {}
error_lectura = None

# --- CONSTANTES DE CONFIGURACIÓN ---
HOJAS_PANEL = [
    'Panel', 
    'Microdinamia',
    'Macrodinamia',
    'Ventilatorio',
    'Neurocrítico'
]

# URLs de imágenes temáticas (baja opacidad y desenfoque para fondo)
BACKGROUND_IMAGES = {
    'Panel': "https://placehold.co/1000x300/4a5568/ffffff?text=Monitor", # Fondo neutro para el panel principal
    'Macrodinamia': "https://placehold.co/1000x300/2C5282/ffffff/png?text=ECOCARDIOGRAMA+Y+GC", 
    'Hemodinamia': "https://placehold.co/1000x300/805ad5/ffffff/png?text=ECOCARDIOGRAFIA",
    'Microdinamia': "https://placehold.co/1000x300/4c0519/ffffff/png?text=MICROCIRCULACIÓN", 
    'Ventilatorio': "https://placehold.co/1000x300/104e8b/ffffff/png?text=PULMONES+Y+VENTILACION", 
    'Neurocrítico': "https://placehold.co/1000x300/365314/ffffff/png?text=CEREBRO+Y+PIC"
}


# --- Función para cargar el Excel ---
def cargar_datos_excel():
    """Carga todas las hojas del archivo de Excel usando pandas."""
    global datos_hojas, error_lectura
    try:
        # sheet_name=None carga todas las hojas en un diccionario
        datos_hojas = pd.read_excel(EXCEL_FILE_PATH, sheet_name=HOJAS_PANEL)
        error_lectura = None
        return True
    except FileNotFoundError:
        error_lectura = f"ERROR: El archivo '{EXCEL_FILE_PATH}' no se encontró en la carpeta."
        return False
    except Exception as e:
        # Muestra el nombre de la hoja si es un error de hoja no encontrada
        if "No sheet named" in str(e):
             match = re.search(r"No sheet named '([^']*)'", str(e))
             if match:
                hoja_faltante = match.group(1)
                error_lectura = f"ERROR: Hoja de trabajo llamada '{hoja_faltante}' no encontrada. Revisa tu archivo Excel."
             else:
                error_lectura = f"ERROR al leer el archivo Excel: {e}"
        else:
            error_lectura = f"ERROR al leer el archivo Excel: {e}"
        return False

# Carga inicial de datos al iniciar la aplicación
cargar_datos_excel()


# Definición del template HTML (Diseño responsivo con Tailwind)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitoreo UCI</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        
        /* Estilos para el texto en cursiva */
        .subtitle-italic { font-style: italic; }
        
        /* Clases para la tabla de resultados del panel */
        .resultado-grid {
            display: grid;
            grid-template-columns: 1fr 1fr; /* Dos columnas: Etiqueta y Valor */
            gap: 0px 10px;
        }
        /* Estilo base para todos los inputs y selects */
        .input-base {
            transition: background-color 0.2s;
            border-radius: 0.5rem; /* rounded-lg */
            font-size: 0.875rem; /* text-sm */
        }
        
        /* Reducción de tamaño de texto general */
        .text-base { font-size: 0.875rem; /* Equivalente a text-sm */ }
        .text-xl { font-size: 1.125rem; /* Equivalente a text-lg */ }
        .text-2xl { font-size: 1.5rem; /* Equivalente a text-xl */ }
        .text-4xl { font-size: 2.25rem; /* Equivalente a text-3xl */ }

        /* Estilos de fondo para paneles de resultados */
        .bg-panel {
            background-size: cover;
            background-position: center;
            position: relative;
            overflow: hidden;
            border: 1px solid rgba(255, 255, 255, 0.2);
            min-height: 200px;
        }
        .bg-panel::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: rgba(255, 255, 255, 0.9); /* Superposición blanca para legibilidad */
            backdrop-filter: blur(1px);
        }
        .bg-panel > * {
            position: relative;
            z-index: 10;
        }
        /* ESTILO UNIFICADO PARA TODOS LOS SUBTÍTULOS DE RESULTADOS */
        .subtitulo-resultado {
             color: #4338CA; /* indigo-700 */
             font-weight: 700; /* font-bold */
             font-size: 1rem; /* text-base */
             display: block;
             width: 100%;
             margin-top: 0.5rem; /* pt-2 */
             padding-bottom: 0.25rem; /* pb-1 */
             border-bottom: 1px solid #e5e7eb; /* Para que resalten más */
        }
    </style>
</head>
<body class="p-4 md:p-8">
    <div class="max-w-4xl mx-auto bg-white rounded-2xl shadow-xl p-6 md:p-10">
        <h1 class="text-4xl font-extrabold text-center text-indigo-700 mb-1">
            Monitoreo UCI
        </h1>
        
        <!-- Nuevo Título Cursivo -->
        <p class="text-center text-gray-700 mb-6 text-xl subtitle-italic leading-tight">
            ICU–CRIPTOS| Hemodynamic, Respiratory & Neurocritical Intelligence<br>
            <span class="text-gray-500 text-base">Monitoreo del paciente crítico, en la palma de mi mano</span>
        </p>
        
        <!-- Mensaje de estado de la carga de Excel -->
        {% if error_lectura %}
            <div class="p-4 mb-6 bg-red-100 border border-red-400 text-red-700 rounded-lg text-base">
                <p class="font-bold">Error de Carga:</p>
                <p>{{ error_lectura }}</p>
            </div>
        {% else %}
            <!-- FORMULARIO DE ENTRADA DE DATOS DINÁMICOS -->
            <form method="POST" action="/" id="input-form" 
                class="p-6 border border-gray-200 rounded-xl shadow-inner bg-gray-50 mb-8"
                oninput="submitForm()"> <!-- Cálculo en tiempo real -->
                
                <!-- 1. Datos Antropométricos -->
                <div class="grid md:grid-cols-4 gap-4 mb-6 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Datos Antropométricos</h2>
                    </div>

                    <!-- Sexo -->
                    <div>
                        <label for="sexo" class="block text-sm font-medium text-gray-700">Sexo:</label>
                        <select id="sexo" name="sexo" required 
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               onchange="updateBackground(this)">
                            <option value="">Selecciona</option>
                            <option value="H" {% if inputs.sexo == 'H' %}selected{% endif %}>Hombre</option>
                            <option value="M" {% if inputs.sexo == 'M' %}selected{% endif %}>Mujer</option>
                        </select>
                    </div>

                    <!-- Edad -->
                    <div>
                        <label for="edad_años" class="block text-sm font-medium text-gray-700">Edad:</label>
                        <input type="number" step="1" id="edad_años" name="edad_años" required 
                               value="{{ inputs.edad_años or '' }}"
                               placeholder="Edad en años"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Peso -->
                    <div>
                        <label for="peso_kg" class="block text-sm font-medium text-gray-700">Peso:</label>
                        <input type="number" step="0.1" id="peso_kg" name="peso_kg" required 
                               value="{{ inputs.peso_kg or '' }}"
                               placeholder="Peso en Kg"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Talla -->
                    <div>
                        <label for="talla_m" class="block text-sm font-medium text-gray-700">Talla:</label>
                        <input type="number" step="0.01" id="talla_m" name="talla_m" required 
                               value="{{ inputs.talla_m or '' }}"
                               placeholder="Talla en metros"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                </div>
                
                <!-- 2. Signos Vitales y Hemodinámica -->
                <div class="grid md:grid-cols-4 gap-4 mb-6 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Signos Vitales y Hemodinámica</h2>
                    </div>

                    <!-- TAS -->
                    <div>
                        <label for="tas" class="block text-sm font-medium text-gray-700">TAS:</label>
                        <input type="number" step="1" id="tas" name="tas" required 
                               value="{{ inputs.tas or '' }}"
                               placeholder="TAS (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- TAD -->
                    <div>
                        <label for="tad" class="block text-sm font-medium text-gray-700">TAD:</label>
                        <input type="number" step="1" id="tad" name="tad" required 
                               value="{{ inputs.tad or '' }}"
                               placeholder="TAD (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- FC -->
                    <div>
                        <label for="fc" class="block text-sm font-medium text-gray-700">FC:</label>
                        <input type="number" step="1" id="fc" name="fc" required 
                               value="{{ inputs.fc or '' }}"
                               placeholder="FC (lpm)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- SatO2 (SV) -->
                    <div>
                        <label for="sato2_sv" class="block text-sm font-medium text-gray-700">SatO₂ (SV):</label>
                        <input type="number" step="0.1" id="sato2_sv" name="sato2_sv" required 
                               value="{{ inputs.sato2_sv or '' }}"
                               placeholder="SatO₂ (%)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                </div>

                <!-- 3. Gasometría Arterial -->
                <div class="grid md:grid-cols-3 gap-4 mb-6 text-base">
                    <div class="md:col-span-3">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Gasometría Arterial</h2>
                    </div>

                    <!-- pH (a) -->
                    <div>
                        <label for="ph_a" class="block text-sm font-medium text-gray-700">pH:</label>
                        <input type="number" step="0.01" id="ph_a" name="ph_a" required 
                               value="{{ inputs.ph_a or '' }}"
                               placeholder="pH arterial"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- PaCO2 -->
                    <div>
                        <label for="paco2" class="block text-sm font-medium text-gray-700">PaCO₂:</label>
                        <input type="number" step="0.1" id="paco2" name="paco2" required 
                               value="{{ inputs.paco2 or '' }}"
                               placeholder="PaCO₂ (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- PaO2 -->
                    <div>
                        <label for="pao2" class="block text-sm font-medium text-gray-700">PaO₂:</label>
                        <input type="number" step="0.1" id="pao2" name="pao2" required 
                               value="{{ inputs.pao2 or '' }}"
                               placeholder="PaO₂ (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- SatO2 (a) -->
                    <div>
                        <label for="sato2_a" class="block text-sm font-medium text-gray-700">SatO₂ (a):</label>
                        <input type="number" step="0.1" id="sato2_a" name="sato2_a" required 
                               value="{{ inputs.sato2_a or '' }}"
                               placeholder="SatO₂ (%)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Lactato -->
                    <div>
                        <label for="lactato" class="block text-sm font-medium text-gray-700">Lactato:</label>
                        <input type="number" step="0.01" id="lactato" name="lactato" required 
                               value="{{ inputs.lactato or '' }}"
                               placeholder="Lactato (mmol/L)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Hb -->
                    <div>
                        <label for="hb" class="block text-sm font-medium text-gray-700">Hb:</label>
                        <input type="number" step="0.1" id="hb" name="hb" required 
                               value="{{ inputs.hb or '' }}"
                               placeholder="Hemoglobina (g/dL)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                </div>
                
                <!-- 4. Gasometría Venosa -->
                <div class="grid md:grid-cols-4 gap-4 mb-6 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Gasometría Venosa</h2>
                    </div>

                    <!-- pHv -->
                    <div>
                        <label for="ph_v" class="block text-sm font-medium text-gray-700">pHv:</label>
                        <input type="number" step="0.01" id="ph_v" name="ph_v" required 
                               value="{{ inputs.ph_v or '' }}"
                               placeholder="pH venoso"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- PvCO2 -->
                    <div>
                        <label for="pvco2" class="block text-sm font-medium text-gray-700">PvCO₂:</label>
                        <input type="number" step="0.1" id="pvco2" name="pvco2" required 
                               value="{{ inputs.pvco2 or '' }}"
                               placeholder="PvCO₂ (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- PvO2 -->
                    <div>
                        <label for="pvo2" class="block text-sm font-medium text-gray-700">PvO₂:</label>
                        <input type="number" step="0.1" id="pvo2" name="pvo2" required 
                               value="{{ inputs.pvo2 or '' }}"
                               placeholder="PvO₂ (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- SatvO2 -->
                    <div>
                        <label for="satvo2" class="block text-sm font-medium text-gray-700">SatvO₂:</label>
                        <input type="number" step="0.1" id="satvo2" name="satvo2" required 
                               value="{{ inputs.satvo2 or '' }}"
                               placeholder="SatvO₂ (%)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                </div>

                <!-- 5. Macrodinamia (POCUS) -->
                <div class="grid md:grid-cols-5 gap-4 mb-6 text-base">
                    <div class="md:col-span-5">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">POCUS (Macrodinamia)</h2>
                    </div>

                    <!-- VTI -->
                    <div>
                        <label for="vti" class="block text-sm font-medium text-gray-700">VTI:</label>
                        <input type="number" step="0.1" id="vti" name="vti" required 
                               value="{{ inputs.vti or '' }}"
                               placeholder="VTI (cm)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- TSVI -->
                    <div>
                        <label for="tsvi" class="block text-sm font-medium text-gray-700">TSVI:</label>
                        <input type="number" step="0.1" id="tsvi" name="tsvi" required 
                               value="{{ inputs.tsvi or '' }}"
                               placeholder="TSVI (cm)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- VCI -->
                    <div>
                        <label for="vci" class="block text-sm font-medium text-gray-700">VCI:</label>
                        <input type="number" step="0.1" id="vci" name="vci" required 
                               value="{{ inputs.vci or '' }}"
                               placeholder="VCI (cm)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- VCI Colaps. -->
                    <div>
                        <label for="vci_colaps" class="block text-sm font-medium text-gray-700">VCI Colaps.:</label>
                        <select id="vci_colaps" name="vci_colaps" required 
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               onchange="updateBackground(this)">
                            <option value="">Selecciona Colapso</option>
                            <option value="total" {% if inputs.vci_colaps == 'total' %}selected{% endif %}>Total</option>
                            <option value=">50%" {% if inputs.vci_colaps == '>50%' %}selected{% endif %}> >50% </option>
                            <option value="<50%" {% if inputs.vci_colaps == '<50%' %}selected{% endif %}> <50% </option>
                            <option value="No cambios" {% if inputs.vci_colaps == 'No cambios' %}selected{% endif %}> No cambios </option>
                        </select>
                    </div>

                    <!-- PVC Medido -->
                    <div>
                        <label for="pvc_medido" class="block text-sm font-medium text-gray-700">PVC Medido:</label>
                        <input type="number" step="1" id="pvc_medido" name="pvc_medido" required 
                               value="{{ inputs.pvc_medido or '' }}"
                               placeholder="PVC Medida (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                </div>

                <!-- 6. Hemodinamia (ECHO Avanzado) -->
                <div class="grid md:grid-cols-4 gap-4 mb-6 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Hemodinamia (ECHO)</h2>
                    </div>

                    <!-- VI Inputs -->
                    <div class="md:col-span-2 space-y-4">
                        <!-- Estilo de subtítulo de input VI/VD (se mantiene) -->
                        <h3 class="font-bold text-base text-indigo-700">Ventrículo Izquierdo</h3>
                        {% for label, name, default in [
                            ("MAPSE L:", "mapse_l", "1.0"),
                            ("MAPSE S:", "mapse_s", "1.0"),
                            ("E (onda):", "e_onda", "0.8"),
                            ("A (onda):", "a_onda", "0.6"),
                            ("E' lat:", "eprim_lat", "0.10"),
                            ("E' med:", "eprim_med", "0.08"),
                            ("VFS:", "vfs", "5.0"),
                            ("VFD:", "vfd", "10.0"),
                            ("Long. VI:", "long_vi", "5.0")
                        ] %}
                        <label class="block">
                            <span class="text-gray-600">{{ label }}</span>
                            <input type="number" step="0.01" name="{{ name }}" value="{{ inputs.get(name) or default }}" 
                                   class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm" 
                                   oninput="updateBackground(this)" required>
                        </label>
                        {% endfor %}
                    </div>

                    <!-- VD Inputs -->
                    <div class="md:col-span-2 space-y-4">
                        <!-- Estilo de subtítulo de input VI/VD (se mantiene) -->
                        <h3 class="font-bold text-base text-indigo-700">Ventrículo Derecho</h3>
                        {% for label, name, default in [
                            ("VTmax:", "vtmax", "3.0"),
                            ("TAPSE:", "tapse", "1.8"),
                            ("VTI Pulmonar:", "vti_pulm", "15.0")
                        ] %}
                        <label class="block">
                            <span class="text-gray-600">{{ label }}</span>
                            <input type="number" step="0.1" name="{{ name }}" value="{{ inputs.get(name) or default }}" 
                                   class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm" 
                                   oninput="updateBackground(this)" required>
                        </label>
                        {% endfor %}
                    </div>
                </div>

                <!-- 7. Datos Ventilatorios -->
                <div class="grid md:grid-cols-4 gap-4 mb-6 text-base">
                    <div class="md:col-span-4">
                        <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Datos Ventilatorios</h2>
                    </div>

                    <!-- MODO (Selector) -->
                    <div>
                        <label for="modo" class="block text-sm font-medium text-gray-700">MODO:</label>
                        <select id="modo" name="modo" required 
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               onchange="updateBackground(this)">
                            <option value="">Selecciona Modo</option>
                            <option value="PCV" {% if inputs.modo == 'PCV' %}selected{% endif %}>PCV</option>
                            <option value="VCV" {% if inputs.modo == 'VCV' %}selected{% endif %}>VCV</option>
                        </select>
                    </div>

                    <!-- VT protec. -->
                    <div>
                        <label for="vt_protec" class="block text-sm font-medium text-gray-700">VT protec.:</label>
                        <input type="number" step="0.1" id="vt_protec" name="vt_protec" required 
                               value="{{ inputs.vt_protec or '' }}"
                               placeholder="VT protec. (ml/kg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- VT Ventilador -->
                    <div>
                        <label for="vt_ventilador" class="block text-sm font-medium text-gray-700">VT Ventilador:</label>
                        <input type="number" step="1" id="vt_ventilador" name="vt_ventilador" required 
                               value="{{ inputs.vt_ventilador or '' }}"
                               placeholder="Vt (ml)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- FR -->
                    <div>
                        <label for="fr" class="block text-sm font-medium text-gray-700">FR:</label>
                        <input type="number" step="1" id="fr" name="fr" required 
                               value="{{ inputs.fr or '' }}"
                               placeholder="FR (lpm)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- PeCO2 -->
                    <div>
                        <label for="peco2" class="block text-sm font-medium text-gray-700">PeCO₂:</label>
                        <input type="number" step="0.1" id="peco2" name="peco2" required 
                               value="{{ inputs.peco2 or '' }}"
                               placeholder="PeCO₂ (mmHg)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- PEEP -->
                    <div>
                        <label for="peep" class="block text-sm font-medium text-gray-700">PEEP:</label>
                        <input type="number" step="1" id="peep" name="peep" required 
                               value="{{ inputs.peep or '' }}"
                               placeholder="PEEP (cmH₂O)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- FIO2 -->
                    <div>
                        <label for="fio2" class="block text-sm font-medium text-gray-700">FIO₂:</label>
                        <input type="number" step="0.01" id="fio2" name="fio2" required 
                               value="{{ inputs.fio2 or '' }}"
                               placeholder="FiO₂ (0.x)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Plateau -->
                    <div>
                        <label for="plateau" class="block text-sm font-medium text-gray-700">Plateau:</label>
                        <input type="number" step="1" id="plateau" name="plateau" required 
                               value="{{ inputs.plateau or '' }}"
                               placeholder="Pplat (cmH₂O)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Ppico -->
                    <div>
                        <label for="ppico" class="block text-sm font-medium text-gray-700">Ppico:</label>
                        <input type="number" step="1" id="ppico" name="ppico" required 
                               value="{{ inputs.ppico or '' }}"
                               placeholder="Ppico (cmH₂O)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Cstat -->
                    <div>
                        <label for="cstat_input" class="block text-sm font-medium text-gray-700">Cstat (medida):</label>
                        <input type="number" step="0.1" id="cstat_input" name="cstat_input" required 
                               value="{{ inputs.cstat_input or '' }}"
                               placeholder="Cstat (ml/cmH₂O)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- Cdin -->
                    <div>
                        <label for="cdin_input" class="block text-sm font-medium text-gray-700">Cdin (medida):</label>
                        <input type="number" step="0.1" id="cdin_input" name="cdin_input" required 
                               value="{{ inputs.cdin_input or '' }}"
                               placeholder="Cdin (ml/cmH₂O)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>
                    
                    <!-- V/min -->
                    <div>
                        <label for="v_min" class="block text-sm font-medium text-gray-700">V/min:</label>
                        <input type="number" step="0.1" id="v_min" name="v_min" required 
                               value="{{ inputs.v_min or '' }}"
                               placeholder="Volumen minuto (L/min)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>

                    <!-- POCC -->
                    <div>
                        <label for="pocc" class="block text-sm font-medium text-gray-700">POCC:</label>
                        <input type="number" step="0.1" id="pocc" name="pocc" required 
                               value="{{ inputs.pocc or '' }}"
                               placeholder="PO.1 (cmH₂O)"
                               class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                               oninput="updateBackground(this)">
                    </div>


                </div>
                
                <!-- 8. Monitorización Neurocrítica -->
                <div class="md:col-span-4 mb-6 text-base">
                    <h2 class="text-xl font-semibold text-gray-700 mb-4 border-b pb-2">Monitorización Neurocrítica</h2>

                    <!-- 8.1 DTC - ACM -->
                    <div class="grid md:grid-cols-4 gap-4 mb-4 p-4 border rounded-lg bg-white shadow-sm">
                        <div class="md:col-span-4">
                            <h3 class="font-semibold text-gray-700">DTC (ACM)</h3>
                        </div>
                        
                        <!-- VS ACM -->
                        <div>
                            <label for="vs_acm" class="block text-sm font-medium text-gray-700">VS:</label>
                            <input type="number" step="0.1" id="vs_acm" name="vs_acm" required 
                                    value="{{ inputs.vs_acm or '' }}"
                                    placeholder="VS (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- VD ACM -->
                        <div>
                            <label for="vd_acm" class="block text-sm font-medium text-gray-700">VD:</label>
                            <input type="number" step="0.1" id="vd_acm" name="vd_acm" required 
                                    value="{{ inputs.vd_acm or '' }}"
                                    placeholder="VD (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                    </div>
                    
                    <!-- 8.2 DTC - AB -->
                    <div class="grid md:grid-cols-4 gap-4 mb-4 p-4 border rounded-lg bg-white shadow-sm">
                        <div class="md:col-span-4">
                            <h3 class="font-semibold text-gray-700">DTC (AB)</h3>
                        </div>
                        
                        <!-- VS AB -->
                        <div>
                            <label for="vs_ab" class="block text-sm font-medium text-gray-700">VS:</label>
                            <input type="number" step="0.1" id="vs_ab" name="vs_ab" required 
                                    value="{{ inputs.vs_ab or '' }}"
                                    placeholder="VS (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- VD AB -->
                        <div>
                            <label for="vd_ab" class="block text-sm font-medium text-gray-700">VD:</label>
                            <input type="number" step="0.1" id="vd_ab" name="vd_ab" required 
                                    value="{{ inputs.vd_ab or '' }}"
                                    placeholder="VD (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                    </div>
                    
                    <!-- 8.3 DTC - Genérico -->
                    <div class="grid md:grid-cols-5 gap-4 mb-4 p-4 border rounded-lg bg-white shadow-sm">
                        <div class="md:col-span-5">
                            <h3 class="font-semibold text-gray-700">DTC (Genérico)</h3>
                        </div>

                        <!-- Vaso Select -->
                        <div>
                            <label for="vaso_dtc" class="block text-sm font-medium text-gray-700">Vaso:</label>
                            <select id="vaso_dtc" name="vaso_dtc" required 
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    onchange="updateBackground(this)">
                                <option value="">Selecciona</option>
                                <option value="ACP" {% if inputs.vaso_dtc == 'ACP' %}selected{% endif %}>ACP</option>
                                <option value="ACA" {% if inputs.vaso_dtc == 'ACA' %}selected{% endif %}>ACA</option>
                                <option value="AB" {% if inputs.vaso_dtc == 'AB' %}selected{% endif %}>AB</option>
                                <option value="ACM" {% if inputs.vaso_dtc == 'ACM' %}selected{% endif %}>ACM</option>
                            </select>
                        </div>
                        
                        <!-- VS Genérico -->
                        <div>
                            <label for="vs_dtc" class="block text-sm font-medium text-gray-700">VS:</label>
                            <input type="number" step="0.1" id="vs_dtc" name="vs_dtc" required 
                                    value="{{ inputs.vs_dtc or '' }}"
                                    placeholder="VS (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- VD Genérico -->
                        <div>
                            <label for="vd_dtc" class="block text-sm font-medium text-gray-700">VD:</label>
                            <input type="number" step="0.1" id="vd_dtc" name="vd_dtc" required 
                                    value="{{ inputs.vd_dtc or '' }}"
                                    placeholder="VD (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                    </div>
                    
                    <!-- 8.4 Flujo Vascular Extracraneal -->
                    <div class="grid md:grid-cols-4 gap-4 mb-4 p-4 border rounded-lg bg-white shadow-sm">
                        <div class="md:col-span-4">
                            <h3 class="font-semibold text-gray-700">Flujo Vascular Extracraneal</h3>
                        </div>

                        <!-- VM Art. Carótida interna extracraneal -->
                        <div>
                            <label for="vm_aci" class="block text-sm font-medium text-gray-700">VM Art. Carótida Int.:</label>
                            <input type="number" step="0.1" id="vm_aci" name="vm_aci" required 
                                    value="{{ inputs.vm_aci or '' }}"
                                    placeholder="VM (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- VM Art. Vertebral Extracraneal -->
                        <div>
                            <label for="vm_ave" class="block text-sm font-medium text-gray-700">VM Art. Vertebral:</label>
                            <input type="number" step="0.1" id="vm_ave" name="vm_ave" required 
                                    value="{{ inputs.vm_ave or '' }}"
                                    placeholder="VM (cm/s)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                    </div>

                    <!-- 8.5 VNO -->
                    <div class="grid md:grid-cols-4 gap-4 mb-4 p-4 border rounded-lg bg-white shadow-sm">
                        <div class="md:col-span-4">
                            <h3 class="font-semibold text-gray-700">VNO (Vaina del Nervio Óptico)</h3>
                        </div>

                        <!-- Der. -->
                        <div>
                            <label for="vno_der" class="block text-sm font-medium text-gray-700">Der. (mm):</label>
                            <input type="number" step="0.1" id="vno_der" name="vno_der" required 
                                    value="{{ inputs.vno_der or '' }}"
                                    placeholder="Der. (mm)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                        
                        <!-- Izq. -->
                        <div>
                            <label for="vno_izq" class="block text-sm font-medium text-gray-700">Izq. (mm):</label>
                            <input type="number" step="0.1" id="vno_izq" name="vno_izq" required 
                                    value="{{ inputs.vno_izq or '' }}"
                                    placeholder="Izq. (mm)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- DGO -->
                        <div>
                            <label for="vno_dgo" class="block text-sm font-medium text-gray-700">DGO (mm):</label>
                            <input type="number" step="0.1" id="vno_dgo" name="vno_dgo" required 
                                    value="{{ inputs.vno_dgo or '' }}"
                                    placeholder="DGO (mm)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                    </div>

                    <!-- 8.6 Gasometría jO2 -->
                    <div class="grid md:grid-cols-3 gap-4 mb-4 p-4 border rounded-lg bg-white shadow-sm">
                        <div class="md:col-span-3">
                            <h3 class="font-semibold text-gray-700">Gasometría yugular (jO₂)</h3>
                        </div>
                        
                        <!-- pH -->
                        <div>
                            <label for="ph_jo2" class="block text-sm font-medium text-gray-700">pH:</label>
                            <input type="number" step="0.01" id="ph_jo2" name="ph_jo2" required 
                                    value="{{ inputs.ph_jo2 or '' }}"
                                    placeholder="pH"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- PjCO2 -->
                        <div>
                            <label for="paco2_jo2" class="block text-sm font-medium text-gray-700">PjCO₂:</label>
                            <input type="number" step="0.1" id="paco2_jo2" name="paco2_jo2" required 
                                    value="{{ inputs.paco2_jo2 or '' }}"
                                    placeholder="PjCO₂ (mmHg)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- PjO2 -->
                        <div>
                            <label for="pao2_jo2" class="block text-sm font-medium text-gray-700">PjO₂:</label>
                            <input type="number" step="0.1" id="pao2_jo2" name="pao2_jo2" required 
                                    value="{{ inputs.pao2_jo2 or '' }}"
                                    placeholder="PjO₂ (mmHg)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                        
                        <!-- SatO2 -->
                        <div>
                            <label for="sato2_jo2" class="block text-sm font-medium text-gray-700">SjO₂:</label>
                            <input type="number" step="0.1" id="sato2_jo2" name="sato2_jo2" required 
                                    value="{{ inputs.sato2_jo2 or '' }}"
                                    placeholder="SjO₂ (%)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- Lactato -->
                        <div>
                            <label for="lactato_jo2" class="block text-sm font-medium text-gray-700">Lactato:</label>
                            <input type="number" step="0.01" id="lactato_jo2" name="lactato_jo2" required 
                                    value="{{ inputs.lactato_jo2 or '' }}"
                                    placeholder="Lactato (mmol/L)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>

                        <!-- PvO2 Yugular (H23) -->
                        <div>
                            <label for="pvo2_jo2" class="block text-sm font-medium text-gray-700">PvO₂ Yugular:</label>
                            <input type="number" step="0.1" id="pvo2_jo2" name="pvo2_jo2" required 
                                    value="{{ inputs.pvo2_jo2 or '' }}"
                                    placeholder="PvO₂ Yugular (mmHg)"
                                    class="input-base mt-1 block w-full px-3 py-2 border border-gray-300 shadow-sm"
                                    oninput="updateBackground(this)">
                        </div>
                    </div>
                </div>
                
            </form>

            <!-- 3. Sección de RESULTADOS -->
            {% if resultados_json and not error_calculo %}
                <div class="mt-8">
                    <h2 class="text-2xl font-bold text-gray-800 mb-6">Resultados por Sección</h2>
                    <div class="grid md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                        
                        {% set resultados = json.loads(resultados_json) %}
                        {% for panel_nombre, panel_data in resultados.items() %}
                            {% set bg_img = BACKGROUND_IMAGES[panel_nombre] if panel_nombre in BACKGROUND_IMAGES else '' %}
                            <div class="bg-panel rounded-xl shadow-md p-5 transition duration-200 hover:shadow-lg"
                                style="{% if bg_img %}background-image: url('{{ bg_img }}');{% endif %}">
                                <h3 class="text-xl font-semibold mb-3 text-indigo-800">{{ panel_nombre }}</h3>
                                {% if panel_data.error %}
                                    <div class="p-2 text-sm bg-red-50 text-red-700 font-medium rounded-lg">{{ panel_data.error }}</div>
                                {% else %}
                                    
                                    <!-- Si es el panel principal, usamos la estructura de dos columnas solicitada -->
                                    {% if panel_nombre == 'Panel' %}
                                        <div class="resultado-grid text-sm">
                                            {% for key, valor in panel_data.items() %}
                                                <!-- Aplica estilo subtítulo a los que inician con -- -->
                                                {% if key.startswith('--') %}
                                                    <span class="subtitulo-resultado md:col-span-2">{{ key | replace('--', '') | trim }}:</span>
                                                {% else %}
                                                    <span class="text-gray-600 font-medium text-right pr-2">{{ key }}:</span>
                                                    <span class="text-gray-900 font-bold">{{ valor | safe }}</span>
                                                {% endif %}
                                            {% endfor %}
                                        </div>
                                    {% else %}
                                        <!-- Para Hemodinamia, Neurocrítico, etc., usamos key-value o subtítulo -->
                                        <div class="text-sm">
                                            {% for key, valor in panel_data.items() %}
                                                <!-- Aplica estilo subtítulo a VI, VD, y cualquier otra sección interna que empiece con -- -->
                                                {% if key.startswith('VI') or key.startswith('VD') or key.startswith('--') %}
                                                    <h4 class="subtitulo-resultado">{{ key | replace('--', '') | trim }}</h4>
                                                {% else %}
                                                    <!-- Regular key-value pair -->
                                                    <div class="flex justify-between items-center py-1 border-b border-gray-200 last:border-b-0">
                                                        <span class="text-gray-600 font-medium">{{ key }}:</span>
                                                        <span class="text-gray-900 font-bold">{{ valor | safe }}</span>
                                                    </div>
                                                {% endif %}
                                            {% endfor %}
                                        </div>
                                    {% endif %}
                                {% endif %}
                            </div>
                        {% endfor %}

                    </div>
                    
                    <!-- Cuadro de Abreviaturas Ventilatorias -->
                    {% if resultados.Ventilatorio and not resultados.Ventilatorio.error %}
                        <div class="mt-8 p-5 bg-blue-50 border-l-4 border-blue-400 rounded-lg shadow-inner text-base">
                            <h4 class="text-lg font-semibold text-blue-800 mb-3">Abreviaturas Ventilatorias</h4>
                            <div class="grid grid-cols-2 sm:grid-cols-3 gap-2 text-sm text-gray-700">
                                <div><span class="font-bold">EM:</span> Espacio Muerto</div>
                                <div><span class="font-bold">EV:</span> Eficiencia Ventilatoria</div>
                                <div><span class="font-bold">PpMt:</span> Presión transpulmonar muscular</div>
                                <div><span class="font-bold">PM:</span> Poder Mecánico</div>
                                <div><span class="font-bold">Raw:</span> Resistencia de Vía Aérea</div>
                            </div>
                        </div>
                    {% endif %}

                </div>
            {% elif error_calculo %}
                    <div class="mt-8 p-4 bg-red-100 border border-red-400 text-red-700 rounded-lg text-base" role="alert">
                        <p class="font-bold">Error en el Cálculo de Fórmulas</p>
                        <p class="text-sm">
                            {{ error_calculo }}
                            <br>
                            <span class="font-semibold text-xs text-red-600 block mt-1">
                                El error ha pintado los campos inválidos de rojo tenue.
                            </span>
                        </p>
                    </div>
            {% endif %}
            
        {% endif %}
    </div>

    <!-- Script para enviar el formulario en tiempo real -->
    <script>
        let timeout = null;
        const form = document.getElementById('input-form');
        
        // Esta función envía el formulario después de un breve retraso
        function submitForm() {
            clearTimeout(timeout);
            timeout = setTimeout(() => {
                form.submit();
            }, 300); // 300ms de retraso para evitar envíos constantes mientras se escribe
        }

        // Script para colorear dinámicamente los campos de entrada
        document.addEventListener('DOMContentLoaded', () => {
            const inputs = document.querySelectorAll('.input-base');
            
            // Función que aplica el color inicial basado en el valor
            const applyInitialColor = (input) => {
                const isSelect = input.tagName === 'SELECT';
                let hasValue;

                if (isSelect) {
                    hasValue = input.value !== '' && input.value !== null && input.value !== 'Selecciona' && input.value !== 'Selecciona Modo' && input.value !== 'Selecciona Colapso';
                } else {
                    // Verifica si es un número válido (ignora si está vacío)
                    const numValue = parseFloat(input.value);
                    hasValue = input.value.trim() !== '' && !isNaN(numValue) && isFinite(numValue);
                }

                // Limpia clases de color
                input.classList.remove('bg-white', 'bg-green-100', 'bg-red-100');
                
                // Aplica el color basado en el valor actual
                if (hasValue) {
                    input.classList.add('bg-green-100');
                } else {
                    input.classList.add('bg-white');
                }
            };

            // Aplica el color inicial a todos los inputs
            inputs.forEach(applyInitialColor);

            // Si hubo un error de cálculo (POST fallido), pinta de rojo los inputs que se enviaron vacíos o con error.
            {% if error_calculo %}
                inputs.forEach(input => {
                    const isSelect = input.tagName === 'SELECT';
                    const value = input.value.trim();
                    let hasError = false;

                    // Error 1: Campo requerido vacío
                    if (input.required && value === '' || (isSelect && value === 'Selecciona')) {
                        hasError = true;
                    }
                    // Error 2: Valor numérico inválido (NaN)
                    if (input.type === 'number' && value !== '' && (isNaN(parseFloat(value)) || !isFinite(parseFloat(value)))) {
                        hasError = true;
                    }

                    if (hasError) {
                        input.classList.remove('bg-white', 'bg-green-100');
                        input.classList.add('bg-red-100');
                    }
                });
            {% endif %}

            // Función de escucha para actualizar el color al escribir o cambiar
            window.updateBackground = (input) => {
                const isSelect = input.tagName === 'SELECT';
                let hasValue;
                
                if (isSelect) {
                    hasValue = input.value !== '' && input.value !== null && input.value !== 'Selecciona' && input.value !== 'Selecciona Modo' && input.value !== 'Selecciona Colapso';
                } else {
                    const numValue = parseFloat(input.value);
                    hasValue = input.value.trim() !== '' && !isNaN(numValue) && isFinite(numValue);
                }

                input.classList.remove('bg-white', 'bg-green-100', 'bg-red-100');

                if (hasValue) {
                    input.classList.add('bg-green-100');
                } else {
                    input.classList.add('bg-white');
                }
            };
        });
    </script>
</body>
</html>
"""

# 4. --- Lógica de Replicación de Fórmulas ---
def replicar_formulas(user_inputs):
    """
    Función que replica la lógica de las fórmulas de Excel.
    Recibe la entrada dinámica del usuario.
    """
    resultados = {}

    if error_lectura or not datos_hojas:
        return {"error": "No se cargaron los datos de Excel."}, None

    # --- ENTRADAS DEL USUARIO ---
    
    # Antropométricos (D5 a D8)
    sexo = user_inputs.get('sexo')
    edad = user_inputs.get('edad_años')
    peso_kg = user_inputs.get('peso_kg')
    talla_m = user_inputs.get('talla_m')
    
    # Signos Vitales (D18 a D22)
    tas = user_inputs.get('tas')
    tad = user_inputs.get('tad')
    fc = user_inputs.get('fc')
    sato2_sv = user_inputs.get('sato2_sv') # SatO2 de Signos Vitales
    
    # Gasometría Arterial (D27, D30 a D34)
    ph_a = user_inputs.get('ph_a')
    paco2 = user_inputs.get('paco2')
    pao2 = user_inputs.get('pao2')
    sato2_a = user_inputs.get('sato2_a') # SatO2 arterial (D33)
    lactato = user_inputs.get('lactato')
    hb = user_inputs.get('hb') # Hemoglobina (D27)
    
    # Gasometría Venosa (D37 a D40)
    ph_v = user_inputs.get('ph_v')
    pvco2 = user_inputs.get('pvco2')
    pvo2 = user_inputs.get('pvo2')
    satvo2 = user_inputs.get('satvo2') # SatvO2 (D40)

    # Macrodinamia / POCUS (D7, D8, D13, D14, D18)
    vti = user_inputs.get('vti') # D7
    tsvi = user_inputs.get('tsvi') # D8
    vci = user_inputs.get('vci') # D13
    vci_colaps = user_inputs.get('vci_colaps') # D14 (string)
    pvc_medido = user_inputs.get('pvc_medido') # D18
    
    # Hemodinamia (ECHO Avanzado) (VI: D21-D24, D26-D27, D30-D33, D35-D36 / VD: D29, D41-D44, D46)
    mapse_l = user_inputs.get('mapse_l') # D21
    mapse_s = user_inputs.get('mapse_s') # D22
    e_onda = user_inputs.get('e_onda') # D23
    a_onda = user_inputs.get('a_onda') # D24
    eprim_lat = user_inputs.get('eprim_lat') # D26
    eprim_med = user_inputs.get('eprim_med') # D27
    vfs = user_inputs.get('vfs') # D30
    vfd = user_inputs.get('vfd') # D31
    long_vi = user_inputs.get('long_vi') # D33

    vtmax = user_inputs.get('vtmax') # D41
    tapse = user_inputs.get('tapse') # D43
    vti_pulm = user_inputs.get('vti_pulm') # D44
    
    # Ventilatorio Inputs
    modo = user_inputs.get('modo') 
    vt_protec_ml_kg = user_inputs.get('vt_protec') 
    vt_ventilador = user_inputs.get('vt_ventilador') 
    fr = user_inputs.get('fr') 
    peco2 = user_inputs.get('peco2') 
    peep = user_inputs.get('peep') 
    fio2 = user_inputs.get('fio2') 
    plateau = user_inputs.get('plateau') 
    ppico = user_inputs.get('ppico') 
    cstat_input = user_inputs.get('cstat_input') 
    cdin_input = user_inputs.get('cdin_input') 
    v_min = user_inputs.get('v_min') 
    pocc = user_inputs.get('pocc') 

    # Neurocrítico Inputs (D, G, J, H)
    vs_acm = user_inputs.get('vs_acm')
    vd_acm = user_inputs.get('vd_acm')
    vs_ab = user_inputs.get('vs_ab')
    vd_ab = user_inputs.get('vd_ab')
    vaso_dtc = user_inputs.get('vaso_dtc')
    vs_dtc = user_inputs.get('vs_dtc')
    vd_dtc = user_inputs.get('vd_dtc')
    vm_aci = user_inputs.get('vm_aci')
    vm_ave = user_inputs.get('vm_ave')
    vno_der = user_inputs.get('vno_der')
    vno_izq = user_inputs.get('vno_izq')
    vno_dgo = user_inputs.get('vno_dgo')
    ph_jo2 = user_inputs.get('ph_jo2') 
    paco2_jo2 = user_inputs.get('paco2_jo2') 
    pao2_jo2 = user_inputs.get('pao2_jo2') 
    sato2_jo2 = user_inputs.get('sato2_jo2') # H24 (SjO2)
    lactato_jo2 = user_inputs.get('lactato_jo2') 
    pvo2_jo2 = user_inputs.get('pvo2_jo2') # H23

    
    # --- INICIO DE LA LÓGICA DE TUS 6 HOJAS ---
    # *****************************************************************
    
    try:
        # --- CÁLCULOS BASE DEL PANEL PRINCIPAL (D5 a D40) ---
        
        # D20: TAM = (D18+(2*D19))/3
        tam = (tas + (2 * tad)) / 3
        
        # SCT (D10)
        if talla_m == 0:
            sct = 0.0 # Usamos 0.0 para poder chequear en cálculos de división
        else:
            # D10: SCT = (0.020247*(D7^0.425)*(D8^0.725))*10
            sct = (0.020247 * (peso_kg ** 0.425) * (talla_m ** 0.725)) * 10
            
        # D11: PI (Peso Ideal) - Usado en Ventilatorio!D6
        talla_cm_por_2_54 = (sct * 100) / 2.54 if sct != 0.0 else 0.0
        if sexo == "H":
            pi = 56.2 + 1.41 * (talla_cm_por_2_54 - 60)
        elif sexo == "M":
            pi = 53.1 + 1.36 * (talla_cm_por_2_54 - 60)
        else:
            pi = 0.0 # Usamos 0.0 para evitar errores en otros cálculos

        
        # --- 1. PANEL (Cálculos de la hoja Panel) ---
        panel_resultados = {}
        
        # 1.1 Antropométricos (D5 a D12)
        panel_resultados['-- Datos Antropométricos --'] = " "
        
        panel_resultados['Sexo'] = sexo
        panel_resultados['Edad'] = f"{edad:.0f} años"
        panel_resultados['Peso'] = f"{peso_kg:.0f} Kg" # Sin decimales
        panel_resultados['Talla'] = f"{talla_m:.2f} m"
        
        if talla_m == 0 or sct == 0.0:
            imc = "Error (Talla 0)"
            pi_display = "N/A"
            act = "N/A"
        else:
            # D9: IMC = D7/D8^2
            imc = peso_kg / (talla_m ** 2)
            pi_display = pi # Mostrar el valor ya calculado de PI (D11)
            
            # D12: ACT
            talla_cm = talla_m * 100
            if sexo == "H":
                act = 2.447 - (0.09156 * edad) + (0.3362 * peso_kg) + (0.1074 * talla_cm)
            elif sexo == "M":
                act = -2.097 + (0.1069 * talla_cm) + (0.2466 * peso_kg)
            else:
                act = "Error (Sexo)"

            
        panel_resultados['IMC'] = f"{imc:.2f}" if isinstance(imc, float) and math.isfinite(imc) else imc
        panel_resultados['SCT'] = f"{sct:.2f} m²"
        panel_resultados['PI'] = f"{pi_display:.2f} Kg" if isinstance(pi_display, float) and math.isfinite(pi_display) else pi_display
        panel_resultados['ACT'] = f"{act:.2f} L" if isinstance(act, float) and math.isfinite(act) else act
        
        # 1.2 Signos Vitales y Gasometrías
        panel_resultados['-- Signos Vitales --'] = " "
        
        panel_resultados['TAS'] = f"{tas:.0f} mmHg" # Sin decimales
        panel_resultados['TAD'] = f"{tad:.0f} mmHg" # Sin decimales
        panel_resultados['TAM'] = f"{tam:.0f} mmHg" # Sin decimales
        panel_resultados['FC'] = f"{fc:.0f} lpm" # Sin decimales
        panel_resultados['SatO₂ (SV)'] = f"{sato2_sv:.0f} %" # Sin decimales
        
        panel_resultados['-- Gasometría Arterial --'] = " "
        
        panel_resultados['pH (a)'] = f"{ph_a:.2f}" # Dos decimales
        panel_resultados['PaCO₂'] = f"{paco2:.1f} mmHg"
        panel_resultados['PaO₂'] = f"{pao2:.1f} mmHg"
        panel_resultados['SatO₂ (a)'] = f"{sato2_a:.1f} %"
        panel_resultados['Lactato'] = f"{lactato:.2f} mmol/L" # Dos decimales
        panel_resultados['Hb'] = f"{hb:.1f} g/dL"
        
        panel_resultados['-- Gasometría Venosa --'] = " "
        
        panel_resultados['pHv'] = f"{ph_v:.2f}" # Dos decimales
        panel_resultados['PvCO₂'] = f"{pvco2:.1f} mmHg"
        panel_resultados['PvO₂'] = f"{pvo2:.1f} mmHg"
        panel_resultados['SatvO₂'] = f"{satvo2:.1f} %"
        
        # Asignar resultados al panel
        resultados['Panel'] = panel_resultados
        
        
        # --- CÁLCULOS BASE DE MACRO/MICRO DINAMIA ---
        
        # D11 Macrodinamia: GC (Usado en Microdinamia)
        # GC (D11) = ((TSVI^2 * 0.785) * VTI / 1000) * FC
        vs_macro = ((tsvi ** 2) * 0.785) * vti 
        gc = (vs_macro / 1000) * fc
        
        # D15 Macrodinamia: PVC ECO (Para RVS)
        pvc_eco = "Error"
        if vci < 1.5:
            pvc_eco = 5
        elif vci >= 1.5 and vci <= 2.5:
            if vci_colaps in ["total", ">50%"]:
                pvc_eco = 8
            elif vci_colaps == "<50%":
                pvc_eco = 13
        elif vci > 2.5:
            if vci_colaps == "<50%":
                pvc_eco = 18
            elif vci_colaps == "No cambios":
                pvc_eco = 20
                
        # D5 Microdinamia: CaO2
        sato2_frac = sato2_a / 100.0
        satvo2_frac = satvo2 / 100.0
        cao2 = (1.36 * hb * sato2_frac) + (0.0031 * pao2)
        
        # D14 Microdinamia: ExtO2
        cvo2 = (1.36 * hb * satvo2_frac) + (0.0031 * pvo2)
        davo2 = cao2 - cvo2
        exto2 = (davo2 / cao2) * 100 if cao2 != 0.0 else "Error (CaO₂ 0)"


        # --- 2. MACRODINAMIA (POCUS Central) ---
        macrodinamia_resultados = {}
        
        # D9: TSVI Inferido
        tsvi_inf = (0.01 * (talla_m * 100)) + 0.25
        
        # D10: VS (Calculado arriba)
        
        # D12: IC
        ic = gc / sct if sct != 0.0 else "Error (SCT 0)"
        
        # D17: RVS
        if isinstance(pvc_eco, str) or gc == 0.0:
            rvs = "Error (GC/PVC)"
        else:
            rvs = ((tam - pvc_eco) * 80) / gc
            
        # D18: RVSI
        if sct == 0.0 or isinstance(rvs, str):
            rvsi = "Error (SCT/RVS)"
        else:
            rvsi = rvs / sct

        # Preparar resultados de Macrodinamia 
        macrodinamia_resultados['VTI'] = f"{vti:.2f} cm"
        macrodinamia_resultados['TSVI'] = f"{tsvi:.2f} cm"
        macrodinamia_resultados['TSVI Inferido'] = f"{tsvi_inf:.2f} cm"
        macrodinamia_resultados['VS'] = f"{vs_macro:.0f} ml" # Sin decimales
        macrodinamia_resultados['GC'] = f"{gc:.2f} L/min"
        ic_format = f"{ic:.2f} L/min/m²" if isinstance(ic, float) and math.isfinite(ic) else ic
        macrodinamia_resultados['IC'] = ic_format
        macrodinamia_resultados['VCI'] = f"{vci:.2f} cm"
        macrodinamia_resultados['VCI Colaps.'] = f"{vci_colaps}"
        pvc_eco_format = f"{pvc_eco:.0f} mmHg" if isinstance(pvc_eco, int) else pvc_eco
        macrodinamia_resultados['PVC ECO'] = pvc_eco_format
        macrodinamia_resultados['PVC Medido'] = f"{pvc_medido:.0f} mmHg" # Sin decimales
        rvs_format = f"{rvs:.0f} dyn.s/cm⁵" if isinstance(rvs, float) and math.isfinite(rvs) else rvs
        macrodinamia_resultados['RVS'] = rvs_format
        rvsi_format = f"{rvsi:.0f} dyn.s/cm⁵/m²" if isinstance(rvsi, float) and math.isfinite(rvsi) else rvsi
        macrodinamia_resultados['RVSI'] = rvsi_format
        
        # Asignar resultados al panel
        resultados['Macrodinamia'] = macrodinamia_resultados


        # --- 3. HEMODINAMIA (Cálculos de ECHO Avanzado) ---
        hemodinamia_resultados = {}
        
        # 3.1 Ventrículo Izquierdo
        
        # E/A: =D23/D24
        ea_ratio = e_onda / a_onda if a_onda != 0.0 else "Error (A onda 0)"
        # E' Prom: =(D26+D27)/2
        eprim_prom = (eprim_lat + eprim_med) / 2
        # E/E': =D23/D28
        ee_ratio = e_onda / eprim_prom if eprim_prom != 0.0 else "Error (E' Prom 0)"
        # FEVI SIMP: =((D31-D30)/D31)*100
        fevi_simp = ((vfd - vfs) / vfd) * 100 if vfd != 0.0 else "Error (VFD 0)"
        # Strain MAPSE: =((D21+D22)/2)/D33*100
        strain_mapse = (((mapse_l + mapse_s) / 2) / long_vi) * 100 if long_vi != 0.0 else "Error (Long VI 0)"
        # Ea: =(0.9*Panel!D18)/D10 (0.9 * FC) / VS
        ea = (0.9 * fc) / vs_macro if vs_macro != 0.0 else "Error (VS 0)"
        # Ee: =(0.9*Panel!D18)/D30 (0.9 * FC) / VFS
        ee = (0.9 * fc) / vfs if vfs != 0.0 else "Error (VFS 0)"
        # Power C: =(Panel!D20*D11)/451 (TAM * GC) / 451
        power_c = (tam * gc) / 451 if isinstance(gc, (int, float)) else "Error (GC)"
        
        hemodinamia_resultados['VI'] = "" # Subtítulo
        hemodinamia_resultados['MAPSE L'] = f"{mapse_l:.2f} cm"
        hemodinamia_resultados['MAPSE S'] = f"{mapse_s:.2f} cm"
        hemodinamia_resultados['E (onda)'] = f"{e_onda:.2f}"
        hemodinamia_resultados['A (onda)'] = f"{a_onda:.2f}"
        hemodinamia_resultados['E/A'] = f"{ea_ratio:.2f}" if isinstance(ea_ratio, float) and math.isfinite(ea_ratio) else ea_ratio
        hemodinamia_resultados['E\' lat'] = f"{eprim_lat:.2f}"
        hemodinamia_resultados['E\' med'] = f"{eprim_med:.2f}"
        hemodinamia_resultados['E\' Prom'] = f"{eprim_prom:.2f}"
        hemodinamia_resultados['E/E\''] = f"{ee_ratio:.2f}" if isinstance(ee_ratio, float) and math.isfinite(ee_ratio) else ee_ratio
        hemodinamia_resultados['VFS'] = f"{vfs:.2f}"
        hemodinamia_resultados['VFD'] = f"{vfd:.2f}"
        hemodinamia_resultados['FEVI SIMP'] = f"{fevi_simp:.1f} %" if isinstance(fevi_simp, float) and math.isfinite(fevi_simp) else fevi_simp
        hemodinamia_resultados['Long. VI'] = f"{long_vi:.2f} cm"
        hemodinamia_resultados['Strain MAPSE'] = f"{strain_mapse:.1f} %" if isinstance(strain_mapse, float) and math.isfinite(strain_mapse) else strain_mapse
        hemodinamia_resultados['Ea'] = f"{ea:.2f}" if isinstance(ea, float) and math.isfinite(ea) else ea
        hemodinamia_resultados['Ee'] = f"{ee:.2f}" if isinstance(ee, float) and math.isfinite(ee) else ee
        power_c_format = f"{power_c:.2f} W" if isinstance(power_c, float) and math.isfinite(power_c) else power_c
        hemodinamia_resultados['Power C'] = power_c_format
        
        # 3.2 Ventrículo Derecho
        # Gradiente IT: =4*(D41^2)
        gradiente_it = 4 * (vtmax ** 2)
        # PSAP: =D42+D15 (Gradiente IT + PVC ECO)
        psap = gradiente_it + pvc_eco if isinstance(pvc_eco, (int, float)) else "Error (PVC ECO)"
        # PMAP: =((0.6*D45)+2) (0.6 * PSAP + 2)
        pmap = (0.6 * psap) + 2 if isinstance(psap, (int, float)) else "Error (PSAP)"
        # RVSPulm.: =(((D41/D44)*10)+0.16) (VTmax / VTI Pulmonar * 10 + 0.16)
        rvs_pulm = (((vtmax / vti_pulm) * 10) + 0.16) if vti_pulm != 0.0 else "Error (VTI Pulm 0)"
        # RVSPulm. In.: =(((D46-D40)/D12)*80) (PMAP - SatvO2 / IC * 80)
        rvs_pulm_in = (((pmap - satvo2) / ic) * 80) if isinstance(ic, (int, float)) and ic != 0.0 else "Error (IC 0)"
        # AVD: =D43/D45 (TAPSE / VTI Pulmonar)
        avd = tapse / vti_pulm if vti_pulm != 0.0 else "Error (VTI Pulm 0)"

        hemodinamia_resultados['VD'] = "" # Subtítulo
        hemodinamia_resultados['VTmax'] = f"{vtmax:.2f} m/s"
        hemodinamia_resultados['Gradiente IT'] = f"{gradiente_it:.1f} mmHg"
        hemodinamia_resultados['TAPSE'] = f"{tapse:.2f} cm"
        hemodinamia_resultados['VTI Pulmonar'] = f"{vti_pulm:.1f} cm"
        psap_format = f"{psap:.0f} mmHg" if isinstance(psap, (int, float)) else psap
        hemodinamia_resultados['PSAP'] = psap_format
        pmap_format = f"{pmap:.1f} mmHg" if isinstance(pmap, (int, float)) else pmap
        hemodinamia_resultados['PMAP'] = pmap_format
        rvs_pulm_format = f"{rvs_pulm:.2f}" if isinstance(rvs_pulm, float) and math.isfinite(rvs_pulm) else rvs_pulm
        hemodinamia_resultados['RVSPulm.'] = rvs_pulm_format
        rvs_pulm_in_format = f"{rvs_pulm_in:.2f}" if isinstance(rvs_pulm_in, float) and math.isfinite(rvs_pulm_in) else rvs_pulm_in
        hemodinamia_resultados['RVSPulm. In.'] = rvs_pulm_in_format
        avd_format = f"{avd:.2f}" if isinstance(avd, float) and math.isfinite(avd) else avd
        hemodinamia_resultados['AVD'] = avd_format
        
        # Asignar resultados al panel
        resultados['Hemodinamia'] = hemodinamia_resultados


        # --- 4. MICRODINAMIA (Cálculos de la hoja Microdinamia) ---
        
        micro_resultados = {}
        
        # CÁLCULOS INTERMEDIOS DE MICRODINAMIA (Usan GC, Hb, Sat)
        
        # D6: CvO2
        cvo2 = (1.36 * hb * satvo2_frac) + (0.0031 * pvo2)
        # D7: CcO2
        sato2_sv_frac = sato2_sv / 100.0
        cco2 = ((hb * 1.36) * sato2_sv_frac) + (pao2 * 0.0031)
        # D8: DavO2
        davo2 = cao2 - cvo2
        
        # D9: VO2 (Consumo de O2) - USA GC
        vo2 = gc * (cao2 - cvo2) * 10 if isinstance(gc, (int, float)) else "Error (GC)"
        
        # D10: VO2I (Índice de Consumo de O2) - USA GC y SCT
        vo2i = vo2 / sct if sct != 0.0 and isinstance(vo2, (int, float)) else "Error (SCT/VO₂)"
        
        # D11: DO2 (Transporte/Entrega de O2) - USA GC
        do2 = (gc * cao2) * 10 if isinstance(gc, (int, float)) else "Error (GC)"
        
        # D12: DO2I (Índice de Transporte/Entrega de O2) - USA GC y SCT
        do2i = do2 / sct if sct != 0.0 and isinstance(do2, (int, float)) else "Error (SCT/DO₂)"
        
        # D15: DavCO2
        davco2 = pvco2 - paco2
        
        # D17: GC Fick - Fórmula original
        gc_fick_calc = "Error (DavO₂ 0)"
        if isinstance(davo2, (int, float)) and isinstance(cao2, (int, float)) and davo2 != 0.0 and cao2 != 0.0:
            # La fórmula original de Excel es inusual, replicaremos la estructura: ((DavO2 * 100) / CaO2) / DavO2
            gc_fick_calc = ((davo2 * 100.0) / cao2) / davo2
        elif isinstance(davo2, (int, float)) and davo2 == 0.0:
            gc_fick_calc = "Error (DavO₂ 0)"


        # Preparar resultados de Microdinamia 
        micro_resultados['-- Parámetros de Extracción --'] = ""
        micro_resultados['CaO₂'] = f"{cao2:.2f} ml/dL"
        micro_resultados['CvO₂'] = f"{cvo2:.2f} ml/dL"
        micro_resultados['CcO₂'] = f"{cco2:.2f} ml/dL"
        micro_resultados['DavO₂'] = f"{davo2:.2f} ml/dL"
        vo2_format = f"{vo2:.2f} ml/min" if isinstance(vo2, float) and math.isfinite(vo2) else vo2
        micro_resultados['VO₂'] = vo2_format
        vo2i_format = f"{vo2i:.2f} ml/min/m²" if isinstance(vo2i, float) and math.isfinite(vo2i) else vo2i
        micro_resultados['VO₂I'] = vo2i_format
        do2_format = f"{do2:.2f} ml/min" if isinstance(do2, float) and math.isfinite(do2) else do2
        micro_resultados['DO₂'] = do2_format
        do2i_format = f"{do2i:.2f} ml/min/m²" if isinstance(do2i, float) and math.isfinite(do2i) else do2i
        micro_resultados['DO₂I'] = do2i_format
        exto2_format = f"{exto2:.2f} %" if isinstance(exto2, float) and math.isfinite(exto2) else exto2
        micro_resultados['ExtO₂'] = exto2_format
        micro_resultados['SatvO₂'] = f"{satvo2:.1f} %"
        micro_resultados['DavCO₂'] = f"{davco2:.1f} mmHg"
        micro_resultados['Lactato'] = f"{lactato:.2f} mmol/L"
        gc_fick_calc_format = f"{gc_fick_calc:.2f}" if isinstance(gc_fick_calc, float) and math.isfinite(gc_fick_calc) else gc_fick_calc
        micro_resultados['GC Fick (Calc.)'] = gc_fick_calc_format
        
        resultados['Microdinamia'] = micro_resultados
        
        
        # --- 5. VENTILATORIO (Implementación completa) ---
        ventilatorio_resultados = {}

        # D6: Peso SDRA (PI)
        peso_sdra = pi
        
        # D8: VT protec. C. = D7 * D8 (D7 = VT protec. ml/kg, D8 = Peso SDRA)
        if isinstance(peso_sdra, float) and math.isfinite(peso_sdra):
            vt_protec_calc = vt_protec_ml_kg * peso_sdra
        else:
            vt_protec_calc = "Error (PI no válido)"
            
        # D13: PaCO2 (del panel principal)
        paco2_vent = paco2
        
        # D16: Driving P. = Plateau - PEEP
        driving_p = plateau - peep
        
        # D17: Cstat Calc = D10 / (D16 - D14) (D10=VT Ventilador, D16=Plateau, D14=PEEP)
        if driving_p == 0:
             cstat_calc = "Error (DP 0)"
        else:
            cstat_calc = vt_ventilador / driving_p
            
        # D19: Cdin Calc = D10 / (D18 - D14) (D10=VT Ventilador, D18=Ppico, D14=PEEP)
        ppico_menos_peep = ppico - peep
        if ppico_menos_peep == 0:
             cdin_calc = "Error (Ppico=PEEP)"
        else:
            cdin_calc = vt_ventilador / ppico_menos_peep
            
        # D20: Raw = D18 - D16 (Ppico - Plateau)
        raw = ppico - plateau

        
        # Preparar resultados de Ventilatorio 
        ventilatorio_resultados['MODO'] = modo
        peso_sdra_format = f"{peso_sdra:.2f} Kg" if isinstance(peso_sdra, float) and math.isfinite(peso_sdra) else peso_sdra
        ventilatorio_resultados['-- Peso y VT --'] = ""
        ventilatorio_resultados['Peso SDRA (PI)'] = peso_sdra_format
        ventilatorio_resultados['VT protec.'] = f"{vt_protec_ml_kg:.1f} ml/Kg"
        vt_protec_calc_format = f"{vt_protec_calc:.0f} ml" if isinstance(vt_protec_calc, float) and math.isfinite(vt_protec_calc) else vt_protec_calc
        ventilatorio_resultados['VT protec. C.'] = vt_protec_calc_format
        ventilatorio_resultados['VT Ventilador'] = f"{vt_ventilador:.0f} ml"
        ventilatorio_resultados['FR'] = f"{fr:.0f} lpm"
        ventilatorio_resultados['V/min'] = f"{v_min:.1f} L/min"

        ventilatorio_resultados['-- Presiones y Gases --'] = ""
        ventilatorio_resultados['PEEP'] = f"{peep:.0f} cmH₂O"
        ventilatorio_resultados['FIO₂'] = f"{fio2:.2f}"
        ventilatorio_resultados['PaCO₂'] = f"{paco2_vent:.1f} mmHg"
        ventilatorio_resultados['PeCO₂'] = f"{peco2:.1f} mmHg"
        ventilatorio_resultados['Plateau'] = f"{plateau:.0f} cmH₂O"
        driving_p_format = f"{driving_p:.0f} cmH₂O" if isinstance(driving_p, float) and math.isfinite(driving_p) else driving_p
        ventilatorio_resultados['Driving P.'] = driving_p_format
        ventilatorio_resultados['Ppico'] = f"{ppico:.0f} cmH₂O"
        ventilatorio_resultados['POCC'] = f"{pocc:.1f} cmH₂O"

        ventilatorio_resultados['-- Complianza y Resistencia --'] = ""
        ventilatorio_resultados['Cstat (medida)'] = f"{cstat_input:.1f} ml/cmH₂O"
        cstat_calc_format = f"{cstat_calc:.1f} ml/cmH₂O" if isinstance(cstat_calc, float) and math.isfinite(cstat_calc) else cstat_calc
        ventilatorio_resultados['Cstat Calc'] = cstat_calc_format
        ventilatorio_resultados['Cdin (medida)'] = f"{cdin_input:.1f} ml/cmH₂O"
        cdin_calc_format = f"{cdin_calc:.1f} ml/cmH₂O" if isinstance(cdin_calc, float) and math.isfinite(cdin_calc) else cdin_calc
        ventilatorio_resultados['Cdin Calc'] = cdin_calc_format
        raw_format = f"{raw:.1f} cmH₂O/L/s" if isinstance(raw, float) and math.isfinite(raw) else raw # Unidades Raw ajustadas
        ventilatorio_resultados['Raw'] = raw_format
        
        # Asignar resultados al panel
        resultados['Ventilatorio'] = ventilatorio_resultados
        
        
        # --- 6. NEUROCRÍTICO (Implementación) ---
        
        # Neurocrítico Inputs (D, G, J, H)
        
        
        neuro_resultados = {}
        
        # --- 6.1 DTC - ACM (D9, D10, D11, D12) ---
        
        # D11: VM ACM
        vm_acm = (vs_acm + (2 * vd_acm)) / 3
        
        # D12: IP ACM
        ip_acm = (vs_acm - vd_acm) / vm_acm if vm_acm != 0.0 else "Error (VM 0)"
        
        # D13: IR ACM
        ir_acm = (vs_acm - vd_acm) / vs_acm if vs_acm != 0.0 else "Error (VS 0)"
        
        # D14: PIC (Asumiendo que D12 es IP ACM)
        # PIC: =(10.93*D12)-1.28
        if isinstance(ip_acm, float) and math.isfinite(ip_acm):
            pic = (10.93 * ip_acm) - 1.28
        else:
            pic = "Error (IP ACM)"
        
        # D15: PPC: =Panel!D20 - Neurocrítico!D14
        if isinstance(pic, float) and math.isfinite(pic):
            ppc = tam - pic
        else:
            ppc = "Error (PIC no válido)"
        
        # --- 6.2 DTC - AB (G9, G10, G11, G12) ---

        # G11: VM AB
        vm_ab = (vs_ab + (2 * vd_ab)) / 3

        # G12: IP AB
        ip_ab = (vs_ab - vd_ab) / vm_ab if vm_ab != 0.0 else "Error (VM 0)"
        
        # G13: IR AB
        ir_ab = (vs_ab - vd_ab) / vs_ab if vs_ab != 0.0 else "Error (VS 0)"

        # --- 6.3 DTC - Genérico (J9, J10, J11, J12) ---
        
        # J11: VM Genérico
        vm_dtc = (vs_dtc + (2 * vd_dtc)) / 3

        # J12: IP Genérico
        ip_dtc = (vs_dtc - vd_dtc) / vm_dtc if vm_dtc != 0.0 else "Error (VM 0)"
        
        # J13: IR Genérico
        ir_dtc = (vs_dtc - vd_dtc) / vs_dtc if vs_dtc != 0.0 else "Error (VS 0)"

        # --- 6.5 Índices (D, G, H) ---
        
        # Indice Lindergard: =D11/H6 (H6=VM Art. Carótida Int. = vm_aci)
        il = vm_acm / vm_aci if vm_aci != 0.0 else "Error (ACI 0)"
        
        # Indice de Soustiel: =G11/H7 (H7=VM Art. Vertebral = vm_ave)
        isou = vm_ab / vm_ave if vm_ave != 0.0 else "Error (AVE 0)"
        
        # --- 6.6 VNO (H23) ---
        
        # H23: VNO/DGO
        vno_dgo_calc = (vno_der + vno_izq) / (2 * vno_dgo) if vno_dgo != 0.0 else "Error (DGO 0)"

        # --- 6.8 Neuromonitoreo Yugular (H23, H24, D33, D14) ---
        
        # SjO2: =H24 (Es una entrada)
        sjo2_calc = sato2_jo2
        
        # AVDO2: =Microdinamia!D5 - D33 (CaO2 - PvO2 yugular)
        # Usamos pvo2_jo2 (PvO2 Yugular) como D33.
        avdo2 = cao2 - pvo2_jo2 if isinstance(cao2, (int, float)) and math.isfinite(cao2) else "Error (CaO₂)"
        
        # CEO2: =Microdinamia!D14 - H24 (ExtO2 - SjO2)
        exto2_val = exto2 if isinstance(exto2, (int, float)) else 0.0 # Usamos 0.0 si es error para evitar doble error
        ceo2 = exto2_val - sjo2_calc if isinstance(exto2_val, (int, float)) and math.isfinite(exto2_val) else "Error (ExtO₂)"

        # CvJO2: =(1.36 * Hb * SjO2) + (0.0031 * PjO2)
        cvjo2 = (1.36 * hb * (sjo2_calc / 100.0)) + (0.0031 * pao2_jo2)
        
        
        # --- PREPARACIÓN DE RESULTADOS ---
        
        neuro_resultados = {}
        
        neuro_resultados['-- DTC (ACM) --'] = " "
        neuro_resultados['VS (ACM)'] = f"{vs_acm:.1f} cm/s"
        neuro_resultados['VD (ACM)'] = f"{vd_acm:.1f} cm/s"
        vm_acm_format = f"{vm_acm:.1f} cm/s" if isinstance(vm_acm, float) and math.isfinite(vm_acm) else vm_acm
        neuro_resultados['VM (ACM)'] = vm_acm_format
        neuro_resultados['IP (ACM)'] = f"{ip_acm:.2f}" if isinstance(ip_acm, float) and math.isfinite(ip_acm) else ip_acm
        neuro_resultados['IR (ACM)'] = f"{ir_acm:.2f}" if isinstance(ir_acm, float) and math.isfinite(ir_acm) else ir_acm
        neuro_resultados['PIC (Calc.)'] = f"{pic:.1f} mmHg" if isinstance(pic, float) and math.isfinite(pic) else pic
        neuro_resultados['PPC (Calc.)'] = f"{ppc:.1f} mmHg" if isinstance(ppc, float) and math.isfinite(ppc) else ppc

        neuro_resultados['-- DTC (AB) --'] = " "
        neuro_resultados['VS (AB)'] = f"{vs_ab:.1f} cm/s"
        neuro_resultados['VD (AB)'] = f"{vd_ab:.1f} cm/s"
        vm_ab_format = f"{vm_ab:.1f} cm/s" if isinstance(vm_ab, float) and math.isfinite(vm_ab) else vm_ab
        neuro_resultados['VM (AB)'] = vm_ab_format
        neuro_resultados['IP (AB)'] = f"{ip_ab:.2f}" if isinstance(ip_ab, float) and math.isfinite(ip_ab) else ip_ab
        neuro_resultados['IR (AB)'] = f"{ir_ab:.2f}" if isinstance(ir_ab, float) and math.isfinite(ir_ab) else ir_ab

        neuro_resultados['-- DTC (Genérico) --'] = " "
        neuro_resultados['Vaso Medido'] = vaso_dtc
        neuro_resultados['VS'] = f"{vs_dtc:.1f} cm/s"
        neuro_resultados['VD'] = f"{vd_dtc:.1f} cm/s"
        vm_dtc_format = f"{vm_dtc:.1f} cm/s" if isinstance(vm_dtc, float) and math.isfinite(vm_dtc) else vm_dtc
        neuro_resultados['VM'] = vm_dtc_format
        neuro_resultados['IP'] = f"{ip_dtc:.2f}" if isinstance(ip_dtc, float) and math.isfinite(ip_dtc) else ip_dtc
        neuro_resultados['IR'] = f"{ir_dtc:.2f}" if isinstance(ir_dtc, float) and math.isfinite(ir_dtc) else ir_dtc

        neuro_resultados['-- Flujo Vascular Extracraneal --'] = " "
        neuro_resultados['VM Art. Carótida Int.'] = f"{vm_aci:.1f} cm/s"
        neuro_resultados['VM Art. Vertebral'] = f"{vm_ave:.1f} cm/s"

        neuro_resultados['-- Índices Combinados --'] = " "
        il_format = f"{il:.2f}" if isinstance(il, float) and math.isfinite(il) else il
        neuro_resultados['Índice Lindergard'] = il_format
        isou_format = f"{isou:.2f}" if isinstance(isou, float) and math.isfinite(isou) else isou
        neuro_resultados['Índice de Soustiel'] = isou_format
        
        neuro_resultados['-- VNO (Vaina Nervio Óptico) --'] = " "
        neuro_resultados['Der.'] = f"{vno_der:.1f} mm"
        neuro_resultados['Izq.'] = f"{vno_izq:.1f} mm"
        neuro_resultados['DGO'] = f"{vno_dgo:.1f} mm"
        vno_dgo_calc_format = f"{vno_dgo_calc:.2f}" if isinstance(vno_dgo_calc, float) and math.isfinite(vno_dgo_calc) else vno_dgo_calc
        neuro_resultados['VNO/DGO'] = vno_dgo_calc_format

        neuro_resultados['-- Gasometría yugular (jO₂) --'] = " "
        neuro_resultados['pH'] = f"{ph_jo2:.2f}"
        neuro_resultados['PjCO₂'] = f"{paco2_jo2:.1f} mmHg"
        neuro_resultados['PjO₂'] = f"{pao2_jo2:.1f} mmHg"
        neuro_resultados['SjO₂'] = f"{sato2_jo2:.1f} %"
        neuro_resultados['Lactato'] = f"{lactato_jo2:.2f} mmol/L"
        neuro_resultados['PvO₂ Yugular'] = f"{pvo2_jo2:.1f} mmHg"
        
        neuro_resultados['-- Neuromonitoreo --'] = " "
        neuro_resultados['SjO₂ (Monit.)'] = f"{sjo2_calc:.1f} %"
        avdo2_format = f"{avdo2:.2f}" if isinstance(avdo2, float) and math.isfinite(avdo2) else avdo2
        neuro_resultados['AVDO₂'] = avdo2_format
        ceo2_format = f"{ceo2:.2f}" if isinstance(ceo2, float) and math.isfinite(ceo2) else ceo2
        neuro_resultados['CEO₂'] = ceo2_format
        cvjo2_format = f"{cvjo2:.2f}" if isinstance(cvjo2, float) and math.isfinite(cvjo2) else cvjo2
        neuro_resultados['CvJO₂'] = cvjo2_format

        resultados['Neurocrítico'] = neuro_resultados
        
    except ZeroDivisionError:
        return None, "Error de División por Cero. Revisa los campos de Talla, VM, ACI, AVE, y DGO."
    except KeyError as e:
        return None, f"Error: Una columna necesaria para el cálculo no fue encontrada: {e}."
    except Exception as e:
        return None, f"Error inesperado durante el cálculo: {e.__class__.__name__}: {e}"
        
    # *****************************************************************
    # --- FIN DE LA LÓGICA DE TUS 6 HOJAS ---

    # Devolvemos los resultados como una cadena JSON para ser leída por el template HTML
    return json.dumps(resultados), None


# 5. Ruta principal de Flask
@app.route('/', methods=['GET', 'POST'])
def inicio():
    """Maneja las solicitudes GET (mostrar formulario) y POST (calcular)."""
    
    # Intentamos limpiar la plantilla HTML de caracteres residuales antes de usarla
    try:
        clean_template = re.sub(r'[\r\n\s#]*$', '', HTML_TEMPLATE)
    except Exception:
        clean_template = HTML_TEMPLATE

    error_calculo = None
    resultados_json = None
    
    # Inicialización de inputs (asegura que las variables existan en el primer GET)
    user_inputs = {
        # Antropométricos (D5-D8)
        'sexo': None, 'edad_años': None, 'peso_kg': None, 'talla_m': None,
        # Signos Vitales (D18-D22)
        'tas': None, 'tad': None, 'fc': None, 'sato2_sv': None,
        # Gasometría Arterial (D27, D30-D34)
        'ph_a': None, 'paco2': None, 'pao2': None, 'sato2_a': None, 'lactato': None, 'hb': None,
        # Gasometría Venosa (D37-D40)
        'ph_v': None, 'pvco2': None, 'pvo2': None, 'satvo2': None,
        # Macrodinamia / POCUS (D7, D8, D13, D14, D18)
        'vti': None, 'tsvi': None, 'vci': None, 'vci_colaps': None, 'pvc_medido': None,
        # Hemodinamia (ECHO Avanzado) (VI/VD)
        'mapse_l': None, 'mapse_s': None, 'e_onda': None, 'a_onda': None, 'eprim_lat': None, 'eprim_med': None,
        'vfs': None, 'vfd': None, 'long_vi': None, 'vtmax': None, 'tapse': None, 'vti_pulm': None,
        # Ventilatorio Inputs
        'modo': None, 'vt_protec': None, 'vt_ventilador': None, 'fr': None, 'peco2': None, 'peep': None, 
        'fio2': None, 'plateau': None, 'ppico': None, 'cstat_input': None, 'cdin_input': None, 
        'v_min': None, 'pocc': None,
        # Neurocrítico Inputs (D, G, J, H)
        'vs_acm': None, 'vd_acm': None, 'vs_ab': None, 'vd_ab': None, 'vaso_dtc': None, 'vs_dtc': None, 'vd_dtc': None, 
        'vm_aci': None, 'vm_ave': None, 'vno_der': None, 'vno_izq': None, 'vno_dgo': None,
        'ph_jo2': None, 'paco2_jo2': None, 'pao2_jo2': None, 'sato2_jo2': None, 'lactato_jo2': None, 'pvo2_jo2': None
    }
    
    if request.method == 'POST':
        # Captura los datos del formulario enviado
        try:
            # Captura de entradas (Antropométricas)
            user_inputs['sexo'] = request.form.get('sexo')
            user_inputs['edad_años'] = float(request.form.get('edad_años') or 0.0)
            user_inputs['peso_kg'] = float(request.form.get('peso_kg') or 0.0)
            user_inputs['talla_m'] = float(request.form.get('talla_m') or 0.0)
            
            # Captura de entradas (Signos Vitales y Hemodinámica)
            user_inputs['tas'] = float(request.form.get('tas') or 0.0)
            user_inputs['tad'] = float(request.form.get('tad') or 0.0)
            user_inputs['fc'] = float(request.form.get('fc') or 0.0)
            user_inputs['sato2_sv'] = float(request.form.get('sato2_sv') or 0.0)
            
            # Captura de entradas (Gasometría Arterial)
            user_inputs['ph_a'] = float(request.form.get('ph_a') or 0.0)
            user_inputs['paco2'] = float(request.form.get('paco2') or 0.0)
            user_inputs['pao2'] = float(request.form.get('pao2') or 0.0)
            user_inputs['sato2_a'] = float(request.form.get('sato2_a') or 0.0)
            user_inputs['lactato'] = float(request.form.get('lactato') or 0.0)
            user_inputs['hb'] = float(request.form.get('hb') or 0.0)
            
            # Captura de entradas (Gasometría Venosa)
            user_inputs['ph_v'] = float(request.form.get('ph_v') or 0.0)
            user_inputs['pvco2'] = float(request.form.get('pvco2') or 0.0)
            user_inputs['pvo2'] = float(request.form.get('pvo2') or 0.0)
            user_inputs['satvo2'] = float(request.form.get('satvo2') or 0.0)
            
            # Captura de entradas (Macrodinamia / POCUS)
            user_inputs['vti'] = float(request.form.get('vti') or 0.0)
            user_inputs['tsvi'] = float(request.form.get('tsvi') or 0.0)
            user_inputs['vci'] = float(request.form.get('vci') or 0.0)
            user_inputs['vci_colaps'] = request.form.get('vci_colaps')
            user_inputs['pvc_medido'] = float(request.form.get('pvc_medido') or 0.0)

            # Captura de entradas (Hemodinamia)
            user_inputs['mapse_l'] = float(request.form.get('mapse_l') or 0.0)
            user_inputs['mapse_s'] = float(request.form.get('mapse_s') or 0.0)
            user_inputs['e_onda'] = float(request.form.get('e_onda') or 0.0)
            user_inputs['a_onda'] = float(request.form.get('a_onda') or 0.0)
            user_inputs['eprim_lat'] = float(request.form.get('eprim_lat') or 0.0)
            user_inputs['eprim_med'] = float(request.form.get('eprim_med') or 0.0)
            user_inputs['vfs'] = float(request.form.get('vfs') or 0.0)
            user_inputs['vfd'] = float(request.form.get('vfd') or 0.0)
            user_inputs['long_vi'] = float(request.form.get('long_vi') or 0.0)
            user_inputs['vtmax'] = float(request.form.get('vtmax') or 0.0)
            user_inputs['tapse'] = float(request.form.get('tapse') or 0.0)
            user_inputs['vti_pulm'] = float(request.form.get('vti_pulm') or 0.0)

            # Captura de entradas (Ventilatorio)
            user_inputs['modo'] = request.form.get('modo')
            user_inputs['vt_protec'] = float(request.form.get('vt_protec') or 0.0)
            user_inputs['vt_ventilador'] = float(request.form.get('vt_ventilador') or 0.0)
            user_inputs['fr'] = float(request.form.get('fr') or 0.0)
            user_inputs['peco2'] = float(request.form.get('peco2') or 0.0)
            user_inputs['peep'] = float(request.form.get('peep') or 0.0)
            user_inputs['fio2'] = float(request.form.get('fio2') or 0.0)
            user_inputs['plateau'] = float(request.form.get('plateau') or 0.0)
            user_inputs['ppico'] = float(request.form.get('ppico') or 0.0)
            user_inputs['cstat_input'] = float(request.form.get('cstat_input') or 0.0)
            user_inputs['cdin_input'] = float(request.form.get('cdin_input') or 0.0)
            user_inputs['v_min'] = float(request.form.get('v_min') or 0.0)
            user_inputs['pocc'] = float(request.form.get('pocc') or 0.0)
            
            # Captura de entradas (Neurocrítico)
            user_inputs['vs_acm'] = float(request.form.get('vs_acm') or 0.0)
            user_inputs['vd_acm'] = float(request.form.get('vd_acm') or 0.0)
            user_inputs['vs_ab'] = float(request.form.get('vs_ab') or 0.0)
            user_inputs['vd_ab'] = float(request.form.get('vd_ab') or 0.0)
            user_inputs['vaso_dtc'] = request.form.get('vaso_dtc')
            user_inputs['vs_dtc'] = float(request.form.get('vs_dtc') or 0.0)
            user_inputs['vd_dtc'] = float(request.form.get('vd_dtc') or 0.0)
            user_inputs['vm_aci'] = float(request.form.get('vm_aci') or 0.0)
            user_inputs['vm_ave'] = float(request.form.get('vm_ave') or 0.0)
            user_inputs['vno_der'] = float(request.form.get('vno_der') or 0.0)
            user_inputs['vno_izq'] = float(request.form.get('vno_izq') or 0.0)
            user_inputs['vno_dgo'] = float(request.form.get('vno_dgo') or 0.0)
            user_inputs['ph_jo2'] = float(request.form.get('ph_jo2') or 0.0)
            user_inputs['paco2_jo2'] = float(request.form.get('paco2_jo2') or 0.0)
            user_inputs['pao2_jo2'] = float(request.form.get('pao2_jo2') or 0.0)
            user_inputs['sato2_jo2'] = float(request.form.get('sato2_jo2') or 0.0)
            user_inputs['lactato_jo2'] = float(request.form.get('lactato_jo2') or 0.0)
            user_inputs['pvo2_jo2'] = float(request.form.get('pvo2_jo2') or 0.0)
            
            # Llama a la lógica de negocio (Python)
            resultados_json, error_calculo = replicar_formulas(user_inputs)
            
        except (TypeError, ValueError):
            error_calculo = "Por favor, introduce valores numéricos válidos en todos los campos requeridos."
            
    # Renderiza el template HTML
    return render_template_string(
        clean_template, # Usa la plantilla autolimpiada
        error_lectura=error_lectura,
        resultados_json=resultados_json,
        error_calculo=error_calculo,
        inputs=user_inputs,
        json=json # Pasamos el módulo 'json' al template
    )

# 6. Ejecución del servidor
if __name__ == '__main__':
    # print("-----------------------------------------------------------------------")
    # print("APLICACIÓN DE EXCEL INICIADA:")
    # print(f"Accede a la app en tu navegador: http://172.17.3.182:5002/ (usando tu IP)")
    # print("-----------------------------------------------------------------------")
    # Aseguramos que Flask escuche en todas las interfaces de red (0.0.0.0)
    app.run(debug=True, host='0.0.0.0', port=5002)