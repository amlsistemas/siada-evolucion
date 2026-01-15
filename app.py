from ortools.sat.python import cp_model
import pandas as pd
import streamlit as st
import requests
import datetime
from io import BytesIO

# ============================================================
# SIADA+ Evolución v6.0 - Optimizador de Horarios con CP-SAT
# Código completo y corregido
# ============================================================

st.set_page_config(page_title="SIADA+ Evolución v6.0", layout="wide")


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def safe_text(value, default="N/A"):
    """Convierte valores a texto limpio, cuidando NaN y espacios."""
    if pd.isna(value):
        return default
    text = str(value).strip()
    return default if text == "" or text.lower() == "nan" else text

def parse_trimester(value):
    """Limpia y normaliza el valor de trimestre para usarlo como clave/valor."""
    if pd.isna(value):
        return "1" # Valor por defecto si está vacío
    text = str(value).strip()
    if not text:
        return "1"

    # Intenta convertir a número (int o float)
    try:
        num = float(text.replace(",", "."))
        # Si es un entero (ej: 2.0), lo guardamos como entero para claves, pero lo devolvemos como string si es necesario
        return str(int(num)) if num.is_integer() else str(num)
    except ValueError:
        # Si es texto (ej: Semestre 1), lo devolvemos tal cual, limpio.
        return safe_text(text, default="1")


def get_dia_semana(fecha):
    """Obtiene el nombre del día de la semana en español."""
    dias_esp = {
        0: "Lunes",
        1: "Martes",
        2: "Miércoles",
        3: "Jueves",
        4: "Viernes",
        5: "Sábado",
        6: "Domingo"
    }
    return dias_esp[fecha.weekday()]


def obtener_festivos_colombia(año_inicio, año_fin):
    """Obtiene los días festivos de Colombia para los años especificados."""
    festivos = set()
    try:
        for year in range(año_inicio, año_fin + 1):
            response = requests.get(
                f"https://date.nager.at/api/v3/publicholidays/{year}/CO",
                timeout=10
            )
            if response.status_code == 200:
                for f in response.json():
                    festivos.add(
                        datetime.datetime.strptime(f["date"], "%Y-%m-%d").date()
                    )
            else:
                st.warning(f"⚠️ No se pudieron cargar festivos para {year}")
    except Exception as e:
        st.warning(f"⚠️ No se pudieron cargar festivos: {e}. Continuando sin ellos.")
    return festivos


def calcular_dias_no_laborables(fecha_inicio, num_dias, festivos):
    """Calcula los índices de días no laborables (fines de semana + festivos)."""
    dias_no_laborables = set()
    for d in range(num_dias):
        fecha_dia = fecha_inicio + datetime.timedelta(days=d)
        if fecha_dia.weekday() >= 5 or fecha_dia in festivos:
            dias_no_laborables.add(d)
    return dias_no_laborables


# ============================================================
# CARGA DE ARCHIVOS
# ============================================================

def cargar_archivos_basic(uploaded_grupos, uploaded_instructores, uploaded_ambientes, uploaded_curriculo):
    """Carga los archivos Excel y parsea a estructuras ricas (listas de dicts)."""
    try:
        grupos_df = pd.read_excel(uploaded_grupos, sheet_name=0)
        instructores_df = pd.read_excel(uploaded_instructores, sheet_name=0)
        ambientes_df = pd.read_excel(uploaded_ambientes, sheet_name=0)
        curriculo_df = pd.read_excel(uploaded_curriculo, sheet_name=0)

        st.info(f"📋 Columnas GRUPOS: {list(grupos_df.columns)}")
        st.info(f"📋 Columnas INSTRUCTORES: {list(instructores_df.columns)}")
        st.info(f"📋 Columnas AMBIENTES: {list(ambientes_df.columns)}")
        st.info(f"📋 Columnas CURRÍCULO: {list(curriculo_df.columns)}")

        # =====================================================
        # PARSEO DE GRUPOS/FICHAS
        # =====================================================
        col_ficha = next((col for col in grupos_df.columns if any(kw in col.lower() for kw in ["grupo", "ficha"])), grupos_df.columns[0])
        if col_ficha != grupos_df.columns[0]:
            st.warning(f"⚠️ Usando '{col_ficha}' como Ficha/Grupo.")

        grupos = []
        col_map_grupos = {
            'programa': ['programa', 'nombre programa'],
            'trimestre': ['trimestre', 'nivel'],
            'jornada': ['jornada', 'turno'],
            'municipio': ['municipio', 'ciudad', 'sede', 'localidad']
        }
        for _, row in grupos_df.iterrows():
            grupo = {'Ficha': safe_text(row.get(col_ficha))}
            for key, candidates in col_map_grupos.items():
                value = "N/A"
                for cand in candidates:
                    matches = [col for col in grupos_df.columns if cand.lower() in col.lower()]
                    if matches:
                        raw_value = row.get(matches[0])
                        if key == 'trimestre':
                            # Parseamos el valor real del trimestre
                            parsed_value = parse_trimester(raw_value)
                            grupo['Trimestre'] = parsed_value # Guardamos el valor real (como string parseado)
                            value = parsed_value
                            break
                        else:
                            candidate_value = safe_text(raw_value)
                            if candidate_value != "N/A":
                                value = candidate_value
                                break
                if key != 'trimestre':
                    grupo[key.capitalize()] = value
            
            grupo.setdefault('Trimestre', "1") # Asegurar que siempre existe la clave
            grupos.append(grupo)

        # =====================================================
        # PARSEO DE INSTRUCTORES
        # =====================================================
        col_instructor = next((col for col in instructores_df.columns if any(kw in col.lower() for kw in ["nombre", "instructor", "docente"])), instructores_df.columns[0])
        if col_instructor != instructores_df.columns[0]:
            st.warning(f"⚠️ Usando '{col_instructor}' como Instructor.")

        instructores = []
        for _, row in instructores_df.iterrows():
            inst = {'Nombre': safe_text(row.get(col_instructor))}
            inst['Jornada del Instructor'] = safe_text(next((row.get(c) for c in instructores_df.columns if any(kw in c.lower() for kw in ['jornada', 'turno'])), None))
            inst['Exclusiones del Instructor'] = safe_text(next((row.get(c) for c in instructores_df.columns if any(kw in c.lower() for kw in ['exclusión', 'restriccion', 'no disponible', 'exclu'])), None))
            inst['Municipio'] = safe_text(next((row.get(c) for c in instructores_df.columns if any(kw in c.lower() for kw in ['municipio', 'ciudad', 'sede'])), None))
            
            inst.setdefault('Jornada del Instructor', "N/A")
            inst.setdefault('Exclusiones del Instructor', "N/A")
            inst.setdefault('Municipio', "N/A")
            instructores.append(inst)

        # =====================================================
        # PARSEO DE AMBIENTES
        # =====================================================
        col_ambiente = next((col for col in ambientes_df.columns if any(kw in col.lower() for kw in ['ambiente', 'aula', 'salón', 'sala', 'laboratorio'])), ambientes_df.columns[0])
        if col_ambiente != ambientes_df.columns[0]:
            st.warning(f"⚠️ Usando '{col_ambiente}' como Ambiente.")

        ambientes = [safe_text(x, default=None) for x in ambientes_df[col_ambiente].dropna().tolist()]
        ambientes = [a for a in ambientes if a]
        if not ambientes:
            ambientes = ["Aula General"]
            st.warning("⚠️ No hay ambientes. Usando 'Aula General'.")

        # =====================================================
        # PARSEO DE CURRÍCULO
        # =====================================================
        curriculo_sessions = []
        col_map_curr = {
            'asignatura': ['asignatura', 'materia', 'curso'],
            'competencia': ['competencia'],
            'resultados': ['resultado', 'aprendizaje', 'ra', 'resultados de aprendizaje'],
            'hora_inicio': ['inicio', 'hora inicio', 'desde'],
            'hora_fin': ['fin', 'hora fin', 'hasta'],
            'horas': ['horas', 'duracion', 'duración'],
            'trimestre': ['trimestre', 'nivel', 'semestre']
        }
        for _, row in curriculo_df.iterrows():
            sess = {}
            has_data = False
            for key, candidates in col_map_curr.items():
                value = None
                for cand in candidates:
                    matches = [col for col in curriculo_df.columns if cand.lower() in col.lower()]
                    if matches:
                        value = safe_text(row.get(matches[0]), default=None)
                        if value:
                            sess[key] = value
                            has_data = True
                            break
                if not value and key not in ['hora_inicio', 'hora_fin', 'horas']:
                    sess[key] = "N/A"

            if not has_data and not any(k in sess for k in ['asignatura', 'competencia']):
                continue

            sess.setdefault('hora_inicio', "08:00")
            sess.setdefault('hora_fin', "12:00")
            sess.setdefault('trimestre', "1") # Default trimester for session lookup if column is missing
            curriculo_sessions.append(sess)

        if not curriculo_sessions:
            curriculo_sessions = [{
                "asignatura": "Clase General",
                "competencia": "N/A",
                "resultados": "N/A",
                "hora_inicio": "08:00",
                "hora_fin": "12:00",
                "trimestre": "1"
            }]
            st.warning("⚠️ Currículo vacío. Usando sesión por defecto.")

        # Cálculo de Horas por Asignación
        horas_por_asignacion = 4
        for sess in curriculo_sessions:
            if 'horas' in sess and sess['horas'] not in (None, "N/A"):
                try:
                    horas_por_asignacion = max(1, int(round(float(sess['horas']))))
                    break
                except Exception:
                    continue
        else:
            for sess in curriculo_sessions:
                try:
                    hora_inicio = datetime.datetime.strptime(sess['hora_inicio'], "%H:%M")
                    hora_fin = datetime.datetime.strptime(sess['hora_fin'], "%H:%M")
                    duracion = (hora_fin - hora_inicio).seconds / 3600
                    if duracion >= 1:
                        horas_por_asignacion = int(duracion)
                        break
                except Exception:
                    continue

        # =====================================================
        # Agrupamos las sesiones del currículo por Trimestre
        # =====================================================
        curriculo_por_trimestre = {}
        for sess in curriculo_sessions:
            # Usamos el valor parseado del trimestre como clave de búsqueda
            tri_key = str(sess.get('trimestre', "1"))
            curriculo_por_trimestre.setdefault(tri_key, []).append(sess)

        for tri, sesiones in curriculo_por_trimestre.items():
            st.info(f"Trimestre Clave '{tri}': {len(sesiones)} sesiones encontradas")

        num_grupos = len(grupos)
        num_instructores = len(instructores)

        st.success(
            f"✅ Cargados: {num_grupos} grupos/fichas, "
            f"{num_instructores} instructores, "
            f"{len(ambientes)} ambientes, "
            f"{len(curriculo_sessions)} sesiones currículo"
        )

        if num_grupos == 0 or num_instructores == 0:
            st.error("❌ Debe haber al menos 1 grupo/ficha y 1 instructor.")
            return None

        return {
            "num_instructores": num_instructores,
            "instructores": instructores,
            "num_grupos": num_grupos,
            "grupos": grupos,
            "ambientes": ambientes,
            "curriculo_sessions": curriculo_sessions,
            "curriculo_por_trimestre": curriculo_por_trimestre,
            "horas_por_asignacion": horas_por_asignacion,
        }

    except Exception as e:
        st.error(f"⚠️ Error al cargar datos: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None


# ============================================================
# GENERADOR DE HORARIO ÓPTIMO
# ============================================================

def generar_horario_optimo(
    num_instructores,
    instructores,
    num_grupos,
    grupos,
    dias_no_laborables,
    ambientes,
    curriculo_sessions,
    curriculo_por_trimestre,
    horas_por_asignacion=4,
    num_dias=75,
    max_horas_semana=40,
    max_dias_semana=6,
    fecha_inicio=None,
    forzar_equidad=True,
):
    """Genera el horario óptimo con estructura completa de columnas."""
    
    if fecha_inicio is None:
        fecha_inicio = datetime.date.today()
    
    dias_laborables = num_dias - len(dias_no_laborables)
    if dias_laborables <= 0:
        return pd.DataFrame(), "NO_DIAS_LABORABLES", {}

    horas_por_asignacion = max(1, horas_por_asignacion)
    model = cp_model.CpModel()

    # ============================================================
    # VARIABLES DE DECISIÓN
    # ============================================================
    asignacion = {}
    for i in range(num_instructores):
        for d in range(num_dias):
            for g in range(num_grupos):
                asignacion[(i, d, g)] = model.NewBoolVar(f"asig_i{i}_d{d}_g{g}")

    # ============================================================
    # RESTRICCIONES
    # ============================================================

    # 1. Días no laborables: nadie trabaja
    for d in dias_no_laborables:
        for g in range(num_grupos):
            for i in range(num_instructores):
                model.Add(asignacion[(i, d, g)] == 0)

    # 2. Cada grupo tiene exactamente 1 instructor por día laborable
    for d in range(num_dias):
        if d not in dias_no_laborables:
            for g in range(num_grupos):
                model.AddExactlyOne([asignacion[(i, d, g)] for i in range(num_instructores)])

    # 3. Un instructor no puede enseñar más de 1 grupo al mismo día
    for i in range(num_instructores):
        for d in range(num_dias):
            model.AddAtMostOne([asignacion[(i, d, g)] for g in range(num_grupos)])

    # 4. Límites semanales
    max_asignaciones_semana = max(1, max_horas_semana // horas_por_asignacion)
    num_semanas = (num_dias + 6) // 7

    for i in range(num_instructores):
        for w in range(num_semanas):
            dia_inicio = w * 7
            dia_fin = min(dia_inicio + 7, num_dias)

            asignaciones_semana = [
                asignacion[(i, d, g)]
                for d in range(dia_inicio, dia_fin)
                for g in range(num_grupos)
            ]
            model.Add(sum(asignaciones_semana) <= max_asignaciones_semana)

            dias_trabajados_semana = []
            for d in range(dia_inicio, dia_fin):
                if d not in dias_no_laborables:
                    trabajo_dia = model.NewBoolVar(f"trabajo_i{i}_d{d}")
                    asignaciones_dia = [asignacion[(i, d, g)] for g in range(num_grupos)]
                    model.Add(sum(asignaciones_dia) >= 1).OnlyEnforceIf(trabajo_dia)
                    model.Add(sum(asignaciones_dia) == 0).OnlyEnforceIf(trabajo_dia.Not())
                    dias_trabajados_semana.append(trabajo_dia)
            if dias_trabajados_semana:
                model.Add(sum(dias_trabajados_semana) <= max_dias_semana)

    # 5. Equidad en la distribución de carga
    total_asignaciones = num_grupos * dias_laborables
    min_asign_por_inst = total_asignaciones // num_instructores if num_instructores > 0 else 0
    max_asign_por_inst = min_asign_por_inst + (1 if num_instructores > 0 and total_asignaciones % num_instructores else 0)
    margen = max(1, min_asign_por_inst // 10) if min_asign_por_inst > 0 else 1

    for i in range(num_instructores):
        total_instructor = sum(
            asignacion[(i, d, g)]
            for d in range(num_dias)
            for g in range(num_grupos)
        )
        if forzar_equidad and num_instructores > 0:
            model.Add(total_instructor >= max(0, min_asign_por_inst - margen))
            model.Add(total_instructor <= max_asign_por_inst + margen)

    # ============================================================
    # OBJETIVO: Minimizar desbalance de carga
    # ============================================================
    cargas = []
    for i in range(num_instructores):
        carga_i = model.NewIntVar(0, total_asignaciones, f'carga_{i}')
        model.Add(carga_i == sum(
            asignacion[(i, d, g)]
            for d in range(num_dias)
            for g in range(num_grupos)
        ))
        cargas.append(carga_i)

    max_carga = model.NewIntVar(0, total_asignaciones, 'max_carga')
    min_carga = model.NewIntVar(0, total_asignaciones, 'min_carga')
    model.AddMaxEquality(max_carga, cargas)
    model.AddMinEquality(min_carga, cargas)
    model.Minimize(max_carga - min_carga)

    # ============================================================
    # RESOLVER
    # ============================================================
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    solver.parameters.num_search_workers = 8
    solver.parameters.log_search_progress = True

    status = solver.Solve(model)

    status_names = {
        cp_model.OPTIMAL: "OPTIMAL",
        cp_model.FEASIBLE: "FEASIBLE",
        cp_model.INFEASIBLE: "INFEASIBLE",
        cp_model.MODEL_INVALID: "MODEL_INVALID",
        cp_model.UNKNOWN: "UNKNOWN",
    }
    status_name = status_names.get(status, "UNKNOWN")

    stats = {
        "status": status_name,
        "tiempo_solver": solver.WallTime(),
        "conflictos": solver.NumConflicts(),
        "ramas": solver.NumBranches(),
        "dias_laborables": dias_laborables,
        "total_asignaciones_esperadas": total_asignaciones,
    }

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        resultados = []
        num_ambientes = len(ambientes)

        for d in range(num_dias):
            fecha_real = fecha_inicio + datetime.timedelta(days=d)
            dia_semana = get_dia_semana(fecha_real)
            es_festivo = d in dias_no_laborables

            for g in range(num_grupos):
                grupo_info = grupos[g]
                # *** 1. Usamos el trimestre REAL del grupo ***
                grupo_trimestre_key = str(grupo_info.get('Trimestre', "1"))

                for i in range(num_instructores):
                    if solver.Value(asignacion[(i, d, g)]) == 1:
                        
                        instructor_info = instructores[i]
                        amb_idx = (d * num_grupos + g) % num_ambientes
                        ambiente = ambientes[amb_idx]

                        # Buscamos las sesiones que corresponden a la clave del trimestre del GRUPO
                        sesiones_trimestre = curriculo_por_trimestre.get(
                            grupo_trimestre_key,
                            curriculo_sessions
                        )
                        sesiones_trimestre = sesiones_trimestre or curriculo_sessions
                        
                        if not sesiones_trimestre:
                            sesiones_trimestre = [{
                                "asignatura": "Sin asignatura",
                                "competencia": "N/A",
                                "resultados": "N/A",
                                "hora_inicio": "08:00",
                                "hora_fin": "12:00"
                            }]

                        # Rotamos cíclicamente SOLO dentro de las sesiones del trimestre correspondiente
                        num_sesiones_trim = len(sesiones_trimestre)
                        curr_idx = (d * num_grupos + g) % num_sesiones_trim
                        curr_info = sesiones_trimestre[curr_idx]

                        # Definición de Jornada (manteniendo la lógica anterior)
                        jornada_options = ["Mañana", "Tarde", "Noche"]
                        jornada_idx = (d + g) % len(jornada_options)
                        jornada_dia = jornada_options[jornada_idx]

                        if jornada_dia == "Mañana":
                            hora_inicio_str = "08:00"
                            hora_fin_str = "13:00"
                        elif jornada_dia == "Tarde":
                            hora_inicio_str = "14:00"
                            hora_fin_str = "18:00"
                        else:
                            hora_inicio_str = "18:00"
                            hora_fin_str = "22:00"

                        resultados.append({
                            "Fecha": fecha_real.strftime("%Y-%m-%d"),
                            "Día": dia_semana,
                            "Jornada": jornada_dia,
                            "Ficha": grupo_info.get('Ficha', 'N/A'),
                            "Programa": grupo_info.get('Programa', 'N/A'),
                            # *** 2. Aquí se muestra el trimestre REAL del grupo ***
                            "Trimestre": grupo_info.get('Trimestre', "N/A"), 
                            "Hora Inicio": hora_inicio_str,
                            "Hora Fin": hora_fin_str,
                            "Municipio": grupo_info.get('Municipio', 'N/A'),
                            "Asignatura": curr_info.get('asignatura', 'N/A'),
                            "Competencia": curr_info.get('competencia', 'N/A'),
                            "Resultados de Aprendizaje": curr_info.get('resultados', 'N/A'),
                            "Instructor": instructor_info.get('Nombre', 'N/A'),
                            "Jornada del Instructor": instructor_info.get('Jornada del Instructor', 'N/A'),
                            "Ambiente": ambiente,
                            "Estado": "Programado",
                            "Notas": "",
                            "Festivo": "Sí" if es_festivo else "No",
                            "Exclusiones del Instructor": instructor_info.get('Exclusiones del Instructor', 'N/A'),
                        })

        df_resultados = pd.DataFrame(resultados)

        if not df_resultados.empty:
            try:
                df_resultados = df_resultados.sort_values(
                    ['Fecha', 'Hora Inicio', 'Programa']
                ).reset_index(drop=True)
            except Exception:
                df_resultados = df_resultados.sort_values('Fecha').reset_index(drop=True)

        return df_resultados, status_name, stats

    return pd.DataFrame(), status_name, stats


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.title("📅 SIADA+ Evolución v6.0")
st.markdown("### Optimizador de Horarios con CP-SAT (Google OR-Tools)")

st.markdown("""
**Estructura del horario generado:**

Fecha, Día, Jornada, Ficha, Programa, Trimestre, Hora Inicio, Hora Fin,
Municipio, Asignatura, Competencia, Resultados de Aprendizaje,
Instructor, Jornada del Instructor, Ambiente, Estado, Notas, Festivo,
Exclusiones del Instructor.
""")

# ============================================================
# SELECTOR DE FECHA
# ============================================================

st.subheader("📅 Configuración de Fecha")

col_fecha1, col_fecha2 = st.columns([1, 2])
with col_fecha1:
    fecha_inicio = st.date_input(
        "📆 Fecha inicial del horario",
        value=datetime.date.today(),
        min_value=datetime.date(2024, 1, 1),
        max_value=datetime.date(2026, 12, 31),
        help="Selecciona la fecha en la que comienza la planificación"
    )

with col_fecha2:
    st.markdown("### ")
    st.markdown(f"""
    <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; text-align: center;">
        <h2 style="margin: 0; color: #1f77b4;">{fecha_inicio.strftime('%d')}</h2>
        <p style="margin: 0; color: #666; font-size: 18px;">{fecha_inicio.strftime('%B').upper()}</p>
        <p style="margin: 0; color: #888; font-size: 16px;">{fecha_inicio.strftime('%Y')}</p>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# CARGA DE ARCHIVOS
# ============================================================

st.markdown("---")
st.subheader("📁 Carga de Archivos Excel")

col1, col2 = st.columns(2)
with col1:
    uploaded_grupos = st.file_uploader(
        "📋 GRUPOS/FICHAS (obligatorio)",
        type=["xlsx", "xls"],
        help="Debe tener al menos una columna: Grupo o Ficha"
    )
    uploaded_instructores = st.file_uploader(
        "👥 INSTRUCTORES (obligatorio)",
        type=["xlsx", "xls"],
        help="Debe tener al menos una columna: Nombre del instructor"
    )
with col2:
    uploaded_ambientes = st.file_uploader(
        "🏫 AMBIENTES (obligatorio)",
        type=["xlsx", "xls"],
        help="Debe tener al menos una columna: Ambiente"
    )
    uploaded_curriculo = st.file_uploader(
        "📚 CURRÍCULO (obligatorio)",
        type=["xlsx", "xls"],
        help="Debe tener columnas: Asignatura, Competencia, Resultados, Horas (opcional)"
    )

# ============================================================
# AYUDA SOBRE ESTRUCTURA DE ARCHIVOS
# ============================================================

with st.expander("📖 Estructura sugerida de archivos Excel"):
    st.markdown("""
    ### GRUPOS (Excel)
    | Columna | Ejemplo |
    |---------|---------|
    | Ficha o Grupo | 2658235 |
    | Programa | Análisis y Desarrollo de Software |
    | Trimestre | 2 |
    | Municipio | Bogotá |
    | Jornada | Mañana |

    ### INSTRUCTORES (Excel)
    | Columna | Ejemplo |
    |---------|---------|
    | Nombre | Juan Pérez |
    | Jornada del Instructor | Mañana |
    | Exclusiones del Instructor | Lunes, Miércoles |

    ### AMBIENTES (Excel)
    | Columna | Ejemplo |
    |---------|--------- |
    | Ambiente | A-101 |

    ### CURRÍCULO (Excel)
    | Columna | Ejemplo |
    |---------|---------|
    | Asignatura | Programación Web |
    | Competencia | Desarrollar aplicaciones web |
    | Resultados de Aprendizaje | Implementa servicios web |
    | Hora Inicio | 08:00 |
    | Hora Fin | 12:00 |
    | Horas (opcional) | 4 |
    """)

# ============================================================
# CONFIGURACIÓN AVANZADA
# ============================================================

with st.expander("⚙️ Configuración avanzada"):
    col_a, col_b = st.columns(2)
    with col_a:
        max_horas_semana = st.slider(
            "Máx. horas/semana por instructor",
            min_value=20, max_value=48, value=40, step=4,
            help="Horas máximas que un instructor puede trabajar por semana"
        )
        num_dias = st.slider(
            "Número total de días a planificar",
            min_value=7, max_value=90, value=75, step=1,
            help="Cantidad de días en el horizonte de planificación"
        )
    with col_b:
        max_dias_semana = st.slider(
            "Máx. días laborales/semana",
            min_value=1, max_value=7, value=6,
            help="Días máximos que un instructor puede trabajar por semana"
        )
        forzar_equidad = st.checkbox(
            "Forzar equidad en asignaciones",
            value=True,
            help="Distribuir las cargas de trabajo de manera equilibrada"
        )

# ============================================================
# BOTÓN PRINCIPAL
# ============================================================

if st.button("🚀 Generar Horario Óptimo", type="primary"):
    if all([uploaded_grupos, uploaded_instructores, uploaded_ambientes, uploaded_curriculo]):
        datos_basic = cargar_archivos_basic(
            uploaded_grupos, uploaded_instructores, uploaded_ambientes, uploaded_curriculo
        )
        if datos_basic is None:
            st.stop()

        año_fin = (fecha_inicio + datetime.timedelta(days=num_dias)).year
        try:
            with st.spinner("📅 Cargando festivos de Colombia..."):
                festivos = obtener_festivos_colombia(fecha_inicio.year, año_fin)
        except Exception as e:
            st.warning(f"⚠️ No se pudieron cargar festivos: {e}")
            festivos = set()

        dias_no_laborables = calcular_dias_no_laborables(fecha_inicio, num_dias, festivos)

        datos = {
            **datos_basic,
            "dias_no_laborables": dias_no_laborables,
            "fecha_inicio": fecha_inicio,
        }

        st.markdown("---")
        st.markdown("### 📊 Resumen de Datos Cargados")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Fichas/Grupos", datos["num_grupos"])
        col2.metric("Instructores", datos["num_instructores"])
        col3.metric("Ambientes", len(datos["ambientes"]))
        col4.metric("Sesiones Currículo", len(datos["curriculo_sessions"]))

        col5, col6, col7 = st.columns(3)
        col5.metric("Días no laborables", len(dias_no_laborables))
        col6.metric("Días laborables", num_dias - len(dias_no_laborables))
        col7.metric("Horas por sesión", datos["horas_por_asignacion"])

        fecha_fin = fecha_inicio + datetime.timedelta(days=num_dias - 1)
        st.info(f"📅 Período: **{fecha_inicio.strftime('%d/%m/%Y')}** al **{fecha_fin.strftime('%d/%m/%Y')}** ({num_dias} días)")

        st.markdown("---")
        with st.spinner("⏳ Optimizando con CP-SAT..."):
            horario_df, status, stats = generar_horario_optimo(
                num_instructores=datos["num_instructores"],
                instructores=datos["instructores"],
                num_grupos=datos["num_grupos"],
                grupos=datos["grupos"],
                dias_no_laborables=datos["dias_no_laborables"],
                ambientes=datos["ambientes"],
                curriculo_sessions=datos["curriculo_sessions"],
                curriculo_por_trimestre=datos["curriculo_por_trimestre"],
                horas_por_asignacion=datos["horas_por_asignacion"],
                num_dias=num_dias,
                max_horas_semana=max_horas_semana,
                max_dias_semana=max_dias_semana,
                fecha_inicio=fecha_inicio,
                forzar_equidad=forzar_equidad,
            )

        st.markdown("### 🔧 Estadísticas del Solver")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Estado", status)
        col2.metric("Tiempo", f"{stats['tiempo_solver']:.2f}s")
        col3.metric("Conflictos", stats["conflictos"])
        col4.metric("Ramas", stats["ramas"])

        if not horario_df.empty:
            st.success(f"✅ ¡Horario generado exitosamente! ({len(horario_df)} asignaciones)")

            tab1, tab2, tab3 = st.tabs(["📊 Horario Completo", "📈 Análisis", "📥 Descargar"])

            columnas_ordenadas = [
                "Fecha", "Día", "Jornada", "Hora Inicio", "Hora Fin",
                "Ficha", "Programa", "Trimestre", "Asignatura",
                "Competencia", "Resultados de Aprendizaje",
                "Instructor", "Jornada del Instructor",
                "Ambiente", "Municipio",
                "Estado", "Notas", "Festivo", "Exclusiones del Instructor"
            ]

            with tab1:
                st.dataframe(
                    horario_df[columnas_ordenadas],
                    use_container_width=True,
                    height=600
                )

            with tab2:
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    st.subheader("📊 Carga por Instructor")
                    cargas_inst = horario_df['Instructor'].value_counts()
                    st.bar_chart(cargas_inst)
                with col_g2:
                    st.subheader("📅 Asignaciones por Día")
                    dias_count = horario_df['Día'].value_counts()
                    st.bar_chart(dias_count)

                col_g3, col_g4 = st.columns(2)
                with col_g3:
                    st.subheader("🏫 Ambientes más utilizados")
                    ambientes_count = horario_df['Ambiente'].value_counts()
                    st.bar_chart(ambientes_count)
                with col_g4:
                    st.subheader("📚 Asignaturas programadas")
                    asig_count = horario_df['Asignatura'].value_counts()
                    st.bar_chart(asig_count)

                st.subheader("📋 Resumen de Cargas por Instructor")
                resumen_cargas = horario_df.groupby('Instructor').agg({
                    'Asignatura': 'count',
                    'Ambiente': 'nunique',
                    'Programa': lambda x: ', '.join(x.unique()[:3])
                }).rename(columns={
                    'Asignatura': 'Sesiones',
                    'Ambiente': 'Ambientes Únicos',
                    'Programa': 'Programas'
                })
                st.dataframe(resumen_cargas.sort_values('Sesiones', ascending=False))

            with tab3:
                st.markdown("#### 📥 Descargar Horario")

                csv = horario_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8')
                st.download_button(
                    label="💾 Descargar CSV",
                    data=csv,
                    file_name=f"horario_siada_{fecha_inicio.strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )

                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    horario_df.to_excel(writer, index=False, sheet_name="Horario Completo")

                    resumen_inst = horario_df.groupby('Instructor').agg({
                        'Asignatura': 'count',
                        'Programa': lambda x: ', '.join(x.unique()[:3])
                    }).reset_index()
                    resumen_inst.columns = ['Instructor', 'Sesiones', 'Programas']
                    resumen_inst.to_excel(writer, index=False, sheet_name="Resumen Instructores")

                    resumen_prog = horario_df.groupby('Programa').agg({
                        'Asignatura': 'count',
                        'Instructor': lambda x: ', '.join(x.unique()[:3])
                    }).reset_index()
                    resumen_prog.columns = ['Programa', 'Sesiones', 'Instructores']
                    resumen_prog.to_excel(writer, index=False, sheet_name="Resumen Programas")

                st.download_button(
                    label="📊 Descargar Excel (3 hojas)",
                    data=buffer.getvalue(),
                    file_name=f"horario_siada_{fecha_inicio.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        else:
            st.error(f"❌ No se encontró solución. Estado: {status}")

            with st.expander("🔍 Diagnóstico detallado"):
                st.write("### Posibles causas:")
                st.write("- Restricciones muy estrictas.")
                st.write("- Pocos instructores para la cantidad de grupos.")
                st.write("- Datos inconsistentes en los archivos Excel.")

                st.write("### Datos actuales:")
                st.json({
                    "Grupos": datos["num_grupos"],
                    "Instructores": datos["num_instructores"],
                    "Ambientes": len(datos["ambientes"]),
                    "Días totales": num_dias,
                    "Días no laborables": len(dias_no_laborables),
                    "Días laborables": num_dias - len(dias_no_laborables),
                    "Horas/semana máx.": max_horas_semana,
                    "Días/semana máx.": max_dias_semana,
                })

                st.write("### Sugerencias:")
                st.write("1. Aumenta el número máximo de días laborales por semana.")
                st.write("2. Aumenta el máximo de horas por semana.")
                st.write("3. Desactiva 'Forzar equidad en asignaciones'.")
                st.write("4. Reduce el número de días a planificar.")

    else:
        st.warning("⚠️ Por favor, sube **todos** los archivos Excel requeridos.")

# ============================================================
# FOOTER
# ============================================================

st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #888;">
        <small>SIADA+ Evolución v6.0 - Optimizador de Horarios | Powered by Google OR-Tools CP-SAT</small>
    </div>
    """,
    unsafe_allow_html=True
)