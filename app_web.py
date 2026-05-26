#!/usr/bin/env python3
"""
Generador de Horarios Semanales / Anuales — Versión Web
========================================================
Powered by Streamlit · Accesible desde cualquier navegador
"""

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import random
from datetime import datetime, timedelta
from collections import defaultdict
from io import BytesIO


# ─────────────────────────────────────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────────────────────────────────────

DAYS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado']

MONTHS_ES = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
]


# ─────────────────────────────────────────────────────────────────────────────
# Utilidades de fecha
# ─────────────────────────────────────────────────────────────────────────────

def get_week_dates(reference_date=None):
    if reference_date is None:
        reference_date = datetime.now()
    monday = reference_date - timedelta(days=reference_date.weekday())
    return [
        (DAYS[i], (monday + timedelta(days=i)).strftime('%d/%m/%Y'))
        for i in range(6)
    ]


def get_year_weeks(year=None):
    if year is None:
        year = datetime.now().year
    jan1 = datetime(year, 1, 1)
    first_monday = jan1 - timedelta(days=jan1.weekday())
    if first_monday.year < year:
        first_monday += timedelta(weeks=1)
    weeks = []
    current = first_monday
    while current.year == year:
        week_dates = [
            (DAYS[i], (current + timedelta(days=i)).strftime('%d/%m/%Y'))
            for i in range(6)
        ]
        weeks.append((current, week_dates))
        current += timedelta(weeks=1)
    return weeks


# ─────────────────────────────────────────────────────────────────────────────
# Normalización de sexo
# ─────────────────────────────────────────────────────────────────────────────

def normalize_sex(value):
    if pd.isna(value):
        return 'Indiferente'
    v = str(value).strip().lower()
    if v in ('h', 'hombre', 'm', 'masculino', 'male', 'masc', 'hom'):
        return 'Hombre'
    if v in ('mujer', 'f', 'femenino', 'female', 'fem', 'w', 'muj'):
        return 'Mujer'
    return 'Indiferente'


def student_can_do(student_sex: str, activity_sex: str) -> bool:
    return activity_sex == 'Indiferente' or student_sex == activity_sex


# ─────────────────────────────────────────────────────────────────────────────
# Algoritmo de generación
# ─────────────────────────────────────────────────────────────────────────────

def generate_week_schedule(activities: list, students: list) -> dict:
    """
    schedule[dia][actividad] = nombre_alumno (o '—')
    - Actividades de sexo específico → prioridad (menos candidatos)
    - Cada alumno: máximo una actividad por día
    - Sin repetición actividad-alumno dentro de la semana (en lo posible)
    """
    schedule = {day: {} for day in DAYS}
    history  = {s['Nombre Alumnos']: set() for s in students}

    for day in DAYS:
        available = {s['Nombre Alumnos']: s['Sexo'] for s in students}
        specific  = [a for a in activities if a['Sexo'] != 'Indiferente']
        general   = [a for a in activities if a['Sexo'] == 'Indiferente']
        random.shuffle(specific)
        random.shuffle(general)

        for activity in specific + general:
            act_name = activity['Nombre actividad']
            act_sex  = activity['Sexo']
            eligible = [sn for sn, ss in available.items()
                        if student_can_do(ss, act_sex)]
            if not eligible:
                schedule[day][act_name] = '—'
                continue
            not_done = [s for s in eligible if act_name not in history[s]]
            pool     = not_done if not_done else eligible
            chosen   = random.choice(pool)
            schedule[day][act_name] = chosen
            del available[chosen]
            history[chosen].add(act_name)

    return schedule


def _prepare_dfs(activities_df, students_df):
    acts = activities_df.copy()
    stus = students_df.copy()
    acts['Sexo'] = acts['Sexo'].apply(normalize_sex)
    stus['Sexo'] = stus['Sexo'].apply(normalize_sex)
    return acts.to_dict('records'), stus.to_dict('records')


def generate_schedule(activities_df, students_df):
    acts, stus = _prepare_dfs(activities_df, students_df)
    return generate_week_schedule(acts, stus)


def generate_annual_schedule(activities_df, students_df, year=None):
    acts, stus = _prepare_dfs(activities_df, students_df)
    result = []
    for monday_date, week_dates in get_year_weeks(year):
        sched = generate_week_schedule(acts, stus)
        result.append((monday_date, week_dates, sched))
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Estilos Excel
# ─────────────────────────────────────────────────────────────────────────────

C_BLUE_DARK  = "1F4E79"
C_BLUE_MED   = "2E75B6"
C_BLUE_LIGHT = "D6E4F0"
C_WHITE      = "FFFFFF"
C_GRAY_LIGHT = "F5F5F5"

THIN = Border(
    left=Side(style='thin', color='AAAAAA'),   right=Side(style='thin', color='AAAAAA'),
    top=Side(style='thin', color='AAAAAA'),    bottom=Side(style='thin', color='AAAAAA'),
)
MED = Border(
    left=Side(style='medium', color='666666'), right=Side(style='medium', color='666666'),
    top=Side(style='medium', color='666666'),  bottom=Side(style='medium', color='666666'),
)

def _fill(h):
    return PatternFill(start_color=h, end_color=h, fill_type='solid')
def _font(bold=False, color="1A1A1A", size=10):
    return Font(bold=bold, color=color, size=size)
def _align(h='center', wrap=True):
    return Alignment(horizontal=h, vertical='center', wrap_text=wrap)


def _write_week_block(ws, start_row, week_dates, schedule, activity_names,
                      show_week_hdr=True, row_h=26, fsz=10):
    n_cols   = len(DAYS) + 1
    last_col = get_column_letter(n_cols)
    row      = start_row

    if show_week_hdr:
        ws.merge_cells(f'A{row}:{last_col}{row}')
        c = ws.cell(row=row, column=1,
                    value=f"Semana del {week_dates[0][1]} al {week_dates[-1][1]}")
        c.font = _font(bold=True, color=C_WHITE, size=fsz)
        c.fill = _fill(C_BLUE_MED)
        c.alignment = _align('center')
        ws.row_dimensions[row].height = 20
        row += 1

    c = ws.cell(row=row, column=1, value="Actividad")
    c.font = _font(bold=True, color=C_WHITE, size=fsz)
    c.fill = _fill(C_BLUE_DARK)
    c.alignment = _align('center')
    c.border = THIN
    ws.row_dimensions[row].height = 40 if fsz >= 10 else 32

    for col_idx, (day, date) in enumerate(week_dates, start=2):
        c = ws.cell(row=row, column=col_idx, value=f"{day}\n{date}")
        c.font = _font(bold=True, color=C_WHITE, size=fsz)
        c.fill = _fill(C_BLUE_MED)
        c.alignment = _align('center')
        c.border = THIN
    row += 1

    for off, act_name in enumerate(activity_names):
        r        = row + off
        row_fill = C_BLUE_LIGHT if off % 2 == 0 else C_WHITE
        c = ws.cell(row=r, column=1, value=act_name)
        c.font = _font(bold=True, size=fsz)
        c.fill = _fill(row_fill)
        c.alignment = _align('left')
        c.border = THIN
        ws.row_dimensions[r].height = row_h

        for col_idx, (day, _) in enumerate(week_dates, start=2):
            student  = schedule[day].get(act_name, '')
            is_empty = student in ('—', '')
            c = ws.cell(row=r, column=col_idx,
                        value='—' if is_empty else student)
            c.font = _font(size=fsz, color='AAAAAA' if is_empty else '1A1A1A')
            c.fill = _fill(C_GRAY_LIGHT if is_empty else row_fill)
            c.alignment = _align('center')
            c.border = THIN

    return row + len(activity_names) + 1


def _set_col_widths(ws, act_w=26, day_w=18):
    ws.column_dimensions['A'].width = act_w
    for col in range(2, len(DAYS) + 2):
        ws.column_dimensions[get_column_letter(col)].width = day_w


# ─────────────────────────────────────────────────────────────────────────────
# Guardar Excel → BytesIO (para descarga web)
# ─────────────────────────────────────────────────────────────────────────────

def save_weekly_excel(schedule, activities_df, week_dates):
    """Retorna un BytesIO con el Excel semanal."""
    act_names = list(activities_df['Nombre actividad'])
    n_cols    = len(DAYS) + 1
    last_col  = get_column_letter(n_cols)

    wb = Workbook()
    ws = wb.active
    ws.title = "Horario Semanal"

    ws.merge_cells(f'A1:{last_col}1')
    c = ws['A1']
    c.value = "HORARIO SEMANAL DE ACTIVIDADES"
    c.font = _font(bold=True, color=C_WHITE, size=15)
    c.fill = _fill(C_BLUE_DARK)
    c.alignment = _align('center')
    c.border = MED
    ws.row_dimensions[1].height = 36

    ws.merge_cells(f'A2:{last_col}2')
    c = ws['A2']
    c.value = f"Semana del {week_dates[0][1]} al {week_dates[-1][1]}"
    c.font = _font(color=C_WHITE, size=10)
    c.fill = _fill(C_BLUE_MED)
    c.alignment = _align('center')
    ws.row_dimensions[2].height = 20

    _write_week_block(ws, 3, week_dates, schedule, act_names,
                      show_week_hdr=False, row_h=32, fsz=10)
    _set_col_widths(ws, act_w=30, day_w=22)
    ws.freeze_panes = 'B4'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def save_annual_excel(annual_data, activities_df, year):
    """Retorna un BytesIO con el Excel anual (12 hojas)."""
    act_names = list(activities_df['Nombre actividad'])
    n_cols    = len(DAYS) + 1
    last_col  = get_column_letter(n_cols)

    wb = Workbook()
    months: dict = defaultdict(list)
    for monday_date, week_dates, schedule in annual_data:
        months[monday_date.month].append((week_dates, schedule))

    first = True
    for month_num in range(1, 13):
        if month_num not in months:
            continue
        month_name = MONTHS_ES[month_num - 1]
        if first:
            ws = wb.active
            ws.title = month_name
            first = False
        else:
            ws = wb.create_sheet(month_name)

        ws.merge_cells(f'A1:{last_col}1')
        c = ws['A1']
        c.value = f"{month_name.upper()} {year}"
        c.font = _font(bold=True, color=C_WHITE, size=13)
        c.fill = _fill(C_BLUE_DARK)
        c.alignment = _align('center')
        c.border = MED
        ws.row_dimensions[1].height = 30

        cur_row = 2
        for week_dates, schedule in months[month_num]:
            cur_row = _write_week_block(ws, cur_row, week_dates, schedule,
                                        act_names, show_week_hdr=True,
                                        row_h=24, fsz=9)
        _set_col_widths(ws, act_w=26, day_w=17)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# Interfaz Web — Streamlit
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Generador de Horarios",
    page_icon="📅",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Estilos CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Encabezado personalizado */
    .app-header {
        background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
        padding: 2rem 2rem 1.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .app-header h1 {
        color: white !important;
        font-size: 1.9rem !important;
        margin: 0 0 0.3rem 0 !important;
    }
    .app-header p {
        color: #BDD7EE !important;
        font-size: 0.95rem !important;
        margin: 0 !important;
    }

    /* Tarjetas de sección */
    .section-card {
        background: #F8FAFD;
        border: 1px solid #D0DEF0;
        border-radius: 10px;
        padding: 1.4rem 1.6rem;
        margin-bottom: 1.2rem;
    }
    .section-title {
        color: #1F4E79;
        font-size: 1rem;
        font-weight: 700;
        margin-bottom: 0.8rem;
    }

    /* Botón de descarga */
    .stDownloadButton > button {
        background-color: #1D6A38 !important;
        color: white !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        padding: 0.7rem 1.5rem !important;
        border-radius: 8px !important;
        border: none !important;
        width: 100% !important;
    }
    .stDownloadButton > button:hover {
        background-color: #145228 !important;
    }

    /* Botón primario */
    .stButton > button[kind="primary"] {
        background-color: #1F4E79 !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        padding: 0.7rem 1.5rem !important;
        border-radius: 8px !important;
        width: 100% !important;
    }

    /* Métricas */
    [data-testid="metric-container"] {
        background: white;
        border: 1px solid #D0DEF0;
        border-radius: 8px;
        padding: 0.8rem 1rem;
    }

    /* Quitar padding extra del contenedor principal */
    .block-container { padding-top: 1.5rem !important; }

    /* Reglas en footer */
    .rules-box {
        background: #EBF3FB;
        border-left: 4px solid #2E75B6;
        border-radius: 0 8px 8px 0;
        padding: 0.9rem 1.2rem;
        font-size: 0.88rem;
        color: #1A1A1A;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Encabezado ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📅 Generador de Horarios</h1>
    <p>Distribución aleatoria &nbsp;·&nbsp; Lunes a Sábado &nbsp;·&nbsp; Restricciones por sexo</p>
</div>
""", unsafe_allow_html=True)

# ── Sección 1 — Subir archivo ────────────────────────────────────────────────
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">1 · Sube tu archivo Excel</div>',
            unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "El archivo debe tener las hojas **Actividades** y **Alumnos**",
    type=['xlsx', 'xls'],
    help="Hoja 'Actividades': columnas 'Nombre actividad' y 'Sexo'.\n"
         "Hoja 'Alumnos': columnas 'Nombre Alumnos' y 'Sexo'."
)

# Link para descargar el Excel de ejemplo
st.caption("¿No tienes un archivo? Usa el de ejemplo que viene en el proyecto.")
st.markdown('</div>', unsafe_allow_html=True)

# ── Procesar archivo ─────────────────────────────────────────────────────────
if uploaded_file is not None:
    try:
        xl = pd.ExcelFile(uploaded_file)

        # Validar hojas
        for sheet in ('Actividades', 'Alumnos'):
            if sheet not in xl.sheet_names:
                st.error(
                    f"❌ No se encontró la hoja **'{sheet}'**.\n\n"
                    f"Hojas encontradas: {', '.join(xl.sheet_names)}"
                )
                st.stop()

        acts = xl.parse('Actividades').dropna(subset=['Nombre actividad'])
        stus = xl.parse('Alumnos').dropna(subset=['Nombre Alumnos'])

        # Validar columnas
        for df, sheet, cols in [
            (acts, 'Actividades', ('Nombre actividad', 'Sexo')),
            (stus, 'Alumnos',     ('Nombre Alumnos',   'Sexo')),
        ]:
            miss = [c for c in cols if c not in df.columns]
            if miss:
                st.error(f"❌ A la hoja **'{sheet}'** le falta(n): **{', '.join(miss)}**")
                st.stop()

        # Métricas del archivo
        col1, col2 = st.columns(2)
        with col1:
            st.metric("✅ Actividades", len(acts))
        with col2:
            st.metric("✅ Alumnos", len(stus))

        st.divider()

        # ── Sección 2 — Tipo de horario ──────────────────────────────────────
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">2 · Tipo de horario</div>',
                    unsafe_allow_html=True)

        schedule_type = st.radio(
            "Selecciona el periodo a generar:",
            ["📅  Esta semana", "📆  Todo el año"],
            horizontal=True,
            label_visibility="collapsed"
        )

        year = None
        if "año" in schedule_type:
            year = st.number_input(
                "Año a generar:",
                min_value=2020, max_value=2060,
                value=datetime.now().year, step=1
            )

        st.markdown('</div>', unsafe_allow_html=True)

        # ── Sección 3 — Generar ──────────────────────────────────────────────
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">3 · Genera y descarga</div>',
                    unsafe_allow_html=True)

        if st.button("⚙  Generar Horario", type="primary", use_container_width=True):
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')

            if "semana" in schedule_type:
                with st.spinner("Generando horario semanal…"):
                    week_dates = get_week_dates()
                    schedule   = generate_schedule(acts, stus)
                    excel_buf  = save_weekly_excel(schedule, acts, week_dates)
                    filename   = f"Horario_Semanal_{ts}.xlsx"
                    n_weeks    = 1
            else:
                with st.spinner(f"Generando horario anual {int(year)} (52 semanas)…"):
                    annual    = generate_annual_schedule(acts, stus, int(year))
                    excel_buf = save_annual_excel(annual, acts, int(year))
                    filename  = f"Horario_Anual_{int(year)}_{ts}.xlsx"
                    n_weeks   = len(annual)

            st.success(
                f"✅ ¡Horario generado!  "
                f"({'1 semana' if n_weeks == 1 else f'{n_weeks} semanas · 12 hojas'})"
            )

            st.download_button(
                label="📥  Descargar Horario Excel",
                data=excel_buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as exc:
        st.error(f"❌ Error inesperado: {exc}")
        st.exception(exc)

# ── Reglas y ayuda ───────────────────────────────────────────────────────────
with st.expander("ℹ️  Reglas de asignación y formato del Excel"):
    st.markdown("""
**Formato del Excel de entrada:**

| Hoja | Columna A | Columna B |
|---|---|---|
| Actividades | Nombre actividad | Sexo (Hombre / Mujer / Indiferente) |
| Alumnos | Nombre Alumnos | Sexo (Hombre / Mujer) |

**Reglas de asignación:**
- 🔵 Actividades **Hombre** → solo alumnos masculinos
- 🔴 Actividades **Mujer** → solo alumnos femeninos
- ⚪ Actividades **Indiferente** → cualquier alumno
- Cada alumno hace **máximo una actividad por día**
- Se evita repetir la misma actividad al mismo alumno **dentro de la semana**
- La distribución es **aleatoria** en cada generación (usa "Generar" nuevamente para obtener otra distribución)
- El símbolo **—** indica que no hay alumnos disponibles con el género requerido
""")

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Generador de Horarios Semanales · Distribución aleatoria con restricciones de género")
