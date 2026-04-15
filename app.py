"""
═══════════════════════════════════════════════════════════════
  App de Gestion de Viaticos y Fondos - Metrica Spa v2
  Dashboard + Wizard de Solicitud + Configuracion
═══════════════════════════════════════════════════════════════
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os, sys, json
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(__file__))

from utils.data_manager import (
    load_default_data, save_default_data,
    load_solicitud, save_solicitud, get_empty_solicitud,
    get_ciudades_destino, get_km_for_ciudad,
    save_to_historial, list_historial, load_historial_entry,
    generar_excel, generar_pdf
)
from utils.calculator import calcular_todo

# ─── Page Config ───
st.set_page_config(
    page_title="Viaticos y Fondos | Metrica Spa",
    page_icon="briefcase",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load CSS
css_path = os.path.join(os.path.dirname(__file__), 'assets', 'style.css')
if os.path.exists(css_path):
    with open(css_path, 'r', encoding='utf-8') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)


# ─── Init Session State ───
def init():
    if 'ref_data' not in st.session_state:
        st.session_state.ref_data = load_default_data()
    if 'wizard_step' not in st.session_state:
        st.session_state.wizard_step = 1
    if 'wizard_data' not in st.session_state:
        st.session_state.wizard_data = get_empty_solicitud()
    if 'wizard_complete' not in st.session_state:
        st.session_state.wizard_complete = False
    if 'show_modal' not in st.session_state:
        st.session_state.show_modal = False

init()


def fmt(value):
    """Formatea como moneda CLP."""
    if value is None or value == 0:
        return "$0"
    return f"${int(value):,}".replace(",", ".")


# ═══════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:12px 0 6px 0;">
        <div style="font-size:36px;">&#128188;</div>
        <h2 style="margin:6px 0 0 0; font-size:18px;
            background:linear-gradient(135deg,#64b5f6,#1e88e5);
            -webkit-background-clip:text; -webkit-text-fill-color:transparent;">
            Viaticos & Fondos
        </h2>
        <p style="color:rgba(180,200,230,0.45); font-size:11px; margin:2px 0 0 0;">Metrica Spa</p>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    page = st.radio("Nav", [
        "Dashboard",
        "Nueva Solicitud",
        "Configuracion"
    ], label_visibility="collapsed")
    st.divider()
    st.markdown('<p class="footer-text">Metrica Spa 2026</p>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════
def page_dashboard():
    st.markdown("# Dashboard de Viaticos")

    historial = list_historial()
    if not historial:
        st.info("No hay solicitudes guardadas. Crea una en 'Nueva Solicitud'.")
        return

    # Selector de solicitud
    options = [f"{h['cliente'] or 'Sin cliente'} - {h['destino'] or '?'} - {h['timestamp'][:10]} - {fmt(h['total'])}" for h in historial]
    sel_idx = st.selectbox("Seleccionar solicitud", range(len(options)), format_func=lambda i: options[i], key="dash_sel")
    entry = load_historial_entry(historial[sel_idx]["filepath"])
    sol = entry.get("solicitud", {})
    calc = entry.get("calculos", {})

    st.divider()

    # KPI row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("FONDO A RENDIR", fmt(calc.get("total_fondo_rendir", 0)))
    with c2:
        st.metric("VIATICO NO RENDIBLE", fmt(calc.get("total_viatico_no_rendible", 0)))
    with c3:
        st.metric("KM TOTALES", f"{calc.get('combustible', {}).get('km_totales', 0):,}".replace(",", "."))
    with c4:
        st.markdown('<div class="pulse-card">', unsafe_allow_html=True)
        st.metric("TOTAL GENERAL", fmt(calc.get("total_general", 0)))
        st.markdown('</div>', unsafe_allow_html=True)

    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs(["Resumen", "Desglose", "Tecnicos", "Descargas"])

    with tab1:
        cl, cr = st.columns([1, 1])
        with cl:
            st.markdown("### Datos del Viaje")
            destinos = sol.get("destinos", [""])
            dest_str = ", ".join([d for d in destinos if d]) or "---"
            info = {
                "Cliente": sol.get("cliente", "---"),
                "Motivo": sol.get("motivo", "---"),
                "ACN": sol.get("acn", "---"),
                "Destino": dest_str,
                "Vehiculo": sol.get("tipo_vehiculo", "---"),
                "Auto/Camion": sol.get("tipo_auto_camion", "---"),
                "Ida y Vuelta": "Si" if sol.get("ida_vuelta", True) else "Solo ida",
                "Tipo viaje": sol.get("tipo_viaje", "Nacional"),
                "Fecha inicio": sol.get("fecha_inicio", "---"),
                "Fecha termino": sol.get("fecha_termino", "---"),
                "Dias": str(sol.get("dias", 0)),
                "Noches": str(sol.get("noches", 0)),
            }
            for k, v in info.items():
                st.markdown(f"""<div style="display:flex; justify-content:space-between; padding:5px 10px;
                    border-bottom:1px solid rgba(68,114,196,0.08); font-size:13px;">
                    <span style="color:rgba(180,200,230,0.55); font-weight:500;">{k}</span>
                    <span style="color:#b0c4de;">{v}</span></div>""", unsafe_allow_html=True)

        with cr:
            st.markdown("### Distribucion de Gastos")
            desglose = calc.get("desglose", {})
            labels = [k for k, v in desglose.items() if v > 0]
            values = [v for v in desglose.values() if v > 0]
            if values:
                colors = ['#1565c0', '#1e88e5', '#42a5f5', '#64b5f6', '#90caf9']
                fig = go.Figure(data=[go.Pie(
                    labels=labels, values=values, hole=0.55,
                    marker=dict(colors=colors[:len(labels)], line=dict(color='#0a1628', width=2)),
                    textfont=dict(size=12, color='white'),
                    textposition='outside', textinfo='label+percent',
                    hovertemplate='<b>%{label}</b><br>%{value:,.0f} CLP<extra></extra>'
                )])
                fig.update_layout(
                    showlegend=False, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    margin=dict(t=10, b=10, l=10, r=10), height=280,
                    annotations=[dict(text=f'<b>{fmt(calc.get("total_general",0))}</b>',
                                      x=0.5, y=0.5, font_size=16, font_color='#64b5f6', showarrow=False)]
                )
                st.plotly_chart(fig, key="dash_pie")
            else:
                st.info("Sin gastos calculados.")

    with tab2:
        st.markdown("### Desglose Detallado")
        rows = []
        categories = {
            "Peajes": calc.get("peaje", {}).get("peaje_ida_vuelta", 0),
            "Alojamiento": calc.get("alojamiento", {}).get("total", 0),
            "Combustible": calc.get("combustible", {}).get("costo_total", 0),
            "Imprevistos": calc.get("imprevistos", {}).get("total_imprevistos", 0),
            "Viaticos Tecnicos": calc.get("total_viatico_no_rendible", 0),
        }
        total = calc.get("total_general", 0) or 1
        for cat, val in categories.items():
            rows.append({"Concepto": cat, "Monto": fmt(val), "% Total": f"{val / total * 100:.1f}%"})
        rows.append({"Concepto": "TOTAL GENERAL", "Monto": fmt(calc.get("total_general", 0)), "% Total": "100%"})
        st.dataframe(pd.DataFrame(rows), hide_index=True, width=800)

        # Details
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**Combustible**")
            comb = calc.get("combustible", {})
            st.markdown(f"- Km totales: **{comb.get('km_totales', 0):,}**")
            st.markdown(f"- Litros: **{comb.get('litros_totales', 0)}**")
            st.markdown(f"- Precio/Lt: **{fmt(comb.get('precio_litro', 0))}**")
        with c2:
            st.markdown("**Alojamiento**")
            aloj = calc.get("alojamiento", {})
            st.markdown(f"- Noches: **{aloj.get('noches', 0)}**")
            st.markdown(f"- Habitaciones: **{aloj.get('num_habitaciones', 0)}**")
            st.markdown(f"- Precio/noche: **{fmt(aloj.get('precio_noche', 0))}**")
        with c3:
            st.markdown("**Imprevistos**")
            imp = calc.get("imprevistos", {})
            st.markdown(f"- Porcentaje: **{imp.get('porcentaje', 0) * 100:.0f}%**")
            st.markdown(f"- Subtotal base: **{fmt(imp.get('subtotal_gastos', 0))}**")

    with tab3:
        st.markdown("### Tecnicos Asignados")
        tecs = sol.get("tecnicos", [])
        if tecs:
            tec_data = [{"Tecnico": t["nombre"], "Viatico No Rendible": fmt(t["monto"])} for t in tecs]
            tec_data.append({"Tecnico": "TOTAL", "Viatico No Rendible": fmt(sum(t["monto"] for t in tecs))})
            st.dataframe(pd.DataFrame(tec_data), hide_index=True, width=600)
        else:
            st.info("Sin tecnicos registrados.")
        st.markdown(f"**Responsable:** {sol.get('responsable', '---')}")
        st.markdown(f"**Jefe de Area:** {sol.get('jefe_area', '---')}")
        st.markdown(f"**Solicitado por:** {sol.get('solicitado_por', '---')}")

    with tab4:
        st.markdown("### Descargar Solicitud")
        c1, c2 = st.columns(2)
        with c1:
            try:
                ref = entry.get("ref_data_snapshot", st.session_state.ref_data)
                excel_bytes = generar_excel(sol, calc, ref)
                st.markdown('<div class="download-btn">', unsafe_allow_html=True)
                st.download_button(
                    "Descargar Excel", data=excel_bytes,
                    file_name=f"Solicitud_Viaticos_{sol.get('cliente', 'sin_cliente').replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dash_dl_xls"
                )
                st.markdown('</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error generando Excel: {e}")
        with c2:
            try:
                ref = entry.get("ref_data_snapshot", st.session_state.ref_data)
                pdf_bytes = generar_pdf(sol, calc, ref)
                st.markdown('<div class="download-btn">', unsafe_allow_html=True)
                st.download_button(
                    "Descargar PDF", data=pdf_bytes,
                    file_name=f"Solicitud_Viaticos_{sol.get('cliente', 'sin_cliente').replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    key="dash_dl_pdf"
                )
                st.markdown('</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error generando PDF: {e}")


# ═══════════════════════════════════════════════════════════
# NUEVA SOLICITUD - WIZARD
# ═══════════════════════════════════════════════════════════
def render_step_indicator(current):
    labels = ["Datos Generales", "Datos del Viaje", "Confirmar"]
    html = '<div class="wizard-step-indicator">'
    for i, label in enumerate(labels, 1):
        if i < current:
            cls = "step-done"
            icon = "&#10003;"
        elif i == current:
            cls = "step-active"
            icon = str(i)
        else:
            cls = "step-pending"
            icon = str(i)
        html += f'<div class="step-dot {cls}">{icon}</div>'
        if i < len(labels):
            line_cls = "step-line-done" if i < current else ""
            html += f'<div class="step-line {line_cls}"></div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def page_nueva_solicitud():
    st.markdown("# Nueva Solicitud de Viaticos")

    if st.session_state.get("show_modal", False):
        show_success_modal()

    step = st.session_state.wizard_step
    render_step_indicator(step)
    st.divider()

    if step == 1:
        wizard_step_1()
    elif step == 2:
        wizard_step_2()
    elif step == 3:
        wizard_step_3()


def wizard_step_1():
    """Paso 1: Datos generales."""
    d = st.session_state.wizard_data
    st.markdown("### Informacion General")

    c1, c2 = st.columns(2)
    with c1:
        d["solicitado_por"] = st.text_input("Solicitado por", value=d.get("solicitado_por", ""), key="w1_sol")
        d["acn"] = st.text_input("ACN", value=d.get("acn", ""), key="w1_acn")
        d["cliente"] = st.text_input("Cliente", value=d.get("cliente", ""), key="w1_cli")
    with c2:
        d["responsable"] = st.text_input("Responsable", value=d.get("responsable", ""), key="w1_resp")
        d["orden_compra"] = st.text_input("Orden de Compra", value=d.get("orden_compra", ""), key="w1_oc")
        d["jefe_area"] = st.text_input("Jefe de Area", value=d.get("jefe_area", "Cristian Sanchez Rojas"), key="w1_jefe")

    d["motivo"] = st.text_area("Motivo del viaje", value=d.get("motivo", ""), height=80, key="w1_motivo")

    st.divider()
    c1, c2, c3 = st.columns([1, 1, 1])
    with c3:
        st.markdown('<div class="wizard-next">', unsafe_allow_html=True)
        if st.button("Siguiente  >>>", key="w1_next", use_container_width=True):
            if not d["cliente"].strip():
                st.error("Debes ingresar un Cliente.")
            else:
                st.session_state.wizard_step = 2
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with c1:
        st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
        if st.button("Limpiar formulario", key="w1_clear", use_container_width=True):
            st.session_state.wizard_data = get_empty_solicitud()
            st.session_state.wizard_step = 1
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)


def wizard_step_2():
    """Paso 2: Datos del viaje."""
    d = st.session_state.wizard_data
    ref = st.session_state.ref_data
    ciudades = get_ciudades_destino(ref)

    st.markdown("### Tecnicos Asignados")
    tecnicos = d.get("tecnicos", [])

    c1, c2, c3 = st.columns([3, 2, 1])
    with c1:
        nuevo_nombre = st.text_input("Nombre del tecnico", key="w2_tn", placeholder="Ej: Jose Rojas")
    with c2:
        nuevo_monto = st.number_input("Viatico no rendible ($)", min_value=0, value=20000, step=5000, key="w2_tm")
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Agregar", key="w2_add"):
            if nuevo_nombre.strip():
                tecnicos.append({"nombre": nuevo_nombre.strip(), "monto": nuevo_monto})
                d["tecnicos"] = tecnicos
                st.rerun()

    if tecnicos:
        for i, t in enumerate(tecnicos):
            c1, c2, c3 = st.columns([4, 2, 1])
            with c1:
                st.markdown(f"**{t['nombre']}**")
            with c2:
                st.markdown(f"*{fmt(t['monto'])}*")
            with c3:
                if st.button("X", key=f"w2_del_{i}"):
                    tecnicos.pop(i)
                    d["tecnicos"] = tecnicos
                    st.rerun()

    st.divider()
    st.markdown("### Datos del Viaje")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        destinos = d.get("destinos", ["", "", "", ""])
        while len(destinos) < 4:
            destinos.append("")
        dest_idx = ciudades.index(destinos[0]) if destinos[0] in ciudades else 0
        destinos[0] = st.selectbox("Destino 1 (principal)", ciudades, index=dest_idx, key="w2_d1")

    with c2:
        iv_val = d.get("ida_vuelta", True)
        d["ida_vuelta"] = st.selectbox("Ida y Vuelta?", [True, False],
                                        index=0 if iv_val else 1,
                                        format_func=lambda x: "Si, Ida y Vuelta" if x else "Solo Ida",
                                        key="w2_iv")
    with c3:
        veh_opts = ["Vehiculo Empresa", "Vehiculo Arrendado", "Vehiculo Propio"]
        veh_idx = veh_opts.index(d.get("tipo_vehiculo", "Vehiculo Empresa")) if d.get("tipo_vehiculo") in veh_opts else 0
        d["tipo_vehiculo"] = st.selectbox("Tipo de vehiculo", veh_opts, index=veh_idx, key="w2_tv")

    with c4:
        ac_opts = ["Auto/Camioneta", "Camion"]
        ac_idx = ac_opts.index(d.get("tipo_auto_camion", "Auto/Camioneta")) if d.get("tipo_auto_camion") in ac_opts else 0
        d["tipo_auto_camion"] = st.selectbox("Auto o Camion", ac_opts, index=ac_idx, key="w2_ac")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        viaje_opts = ["Nacional", "Internacional"]
        viaje_idx = viaje_opts.index(d.get("tipo_viaje", "Nacional")) if d.get("tipo_viaje") in viaje_opts else 0
        d["tipo_viaje"] = st.selectbox("Tipo de viaje", viaje_opts, index=viaje_idx, key="w2_tj")
    with c2:
        d["dias"] = st.number_input("Dias", min_value=1, value=int(d.get("dias", 1)), key="w2_dias")
    with c3:
        d["noches"] = st.number_input("Noches", min_value=0, value=int(d.get("noches", 0)), key="w2_noches")
    with c4:
        hab_opts = ["habitacion_doble", "habitacion_single"]
        hab_labels = ["Habitacion Doble", "Habitacion Single"]
        hab_idx = hab_opts.index(d.get("tipo_habitacion", "habitacion_doble"))
        sel_hab = st.selectbox("Tipo habitacion", hab_labels, index=hab_idx, key="w2_hab")
        d["tipo_habitacion"] = hab_opts[hab_labels.index(sel_hab)]

    c1, c2 = st.columns(2)
    with c1:
        fi = d.get("fecha_inicio", date.today().isoformat())
        if isinstance(fi, str):
            try: fi = date.fromisoformat(fi)
            except: fi = date.today()
        d["fecha_inicio"] = st.date_input("Fecha de Inicio", value=fi, key="w2_fi").isoformat()
    with c2:
        ft = d.get("fecha_termino", date.today().isoformat())
        if isinstance(ft, str):
            try: ft = date.fromisoformat(ft)
            except: ft = date.today()
        d["fecha_termino"] = st.date_input("Fecha de Termino", value=ft, key="w2_ft").isoformat()

    # Destinos adicionales (opcionales)
    with st.expander("Destinos adicionales (opcionales)", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            destinos[1] = st.text_input("Destino 2", value=destinos[1], key="w2_d2")
        with c2:
            destinos[2] = st.text_input("Destino 3", value=destinos[2], key="w2_d3")
        with c3:
            destinos[3] = st.text_input("Destino 4", value=destinos[3], key="w2_d4")
    d["destinos"] = destinos

    st.divider()
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.markdown('<div class="wizard-back">', unsafe_allow_html=True)
        if st.button("<<< Atras", key="w2_back", use_container_width=True):
            st.session_state.wizard_step = 1
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="wizard-next">', unsafe_allow_html=True)
        if st.button("Siguiente  >>>", key="w2_next", use_container_width=True):
            if not tecnicos:
                st.error("Agrega al menos un tecnico.")
            elif not destinos[0]:
                st.error("Selecciona un destino principal.")
            else:
                st.session_state.wizard_step = 3
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)


def wizard_step_3():
    """Paso 3: Confirmacion y calculo."""
    d = st.session_state.wizard_data
    ref = st.session_state.ref_data

    # Set rango alojamiento a promedio
    d["rango_precio_alojamiento"] = "promedio"
    d["fecha_solicitud"] = date.today().isoformat()

    calc = calcular_todo(ref, d)

    st.markdown("### Resumen de la Solicitud")

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("FONDO A RENDIR", fmt(calc["total_fondo_rendir"]))
    with c2:
        st.metric("VIATICO NO RENDIBLE", fmt(calc["total_viatico_no_rendible"]))
    with c3:
        st.metric("KM TOTALES", f"{calc['combustible']['km_totales']:,}".replace(",", "."))
    with c4:
        st.markdown('<div class="pulse-card">', unsafe_allow_html=True)
        st.metric("TOTAL GENERAL", fmt(calc["total_general"]))
        st.markdown('</div>', unsafe_allow_html=True)

    # Data summary in 2 cols
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Datos Generales**")
        for k, v in [("Cliente", d["cliente"]), ("ACN", d["acn"]),
                      ("Motivo", d["motivo"]), ("Responsable", d["responsable"]),
                      ("Solicitado por", d["solicitado_por"])]:
            st.markdown(f"- **{k}:** {v or '---'}")

        st.markdown("**Datos del Viaje**")
        dest_str = ", ".join([x for x in d["destinos"] if x])
        for k, v in [("Destino", dest_str), ("Vehiculo", d["tipo_vehiculo"]),
                      ("Auto/Camion", d["tipo_auto_camion"]),
                      ("Ida y Vuelta", "Si" if d["ida_vuelta"] else "Solo ida"),
                      ("Dias", d["dias"]), ("Noches", d["noches"]),
                      ("Habitacion", "Doble" if d["tipo_habitacion"]=="habitacion_doble" else "Single"),
                      ("Fechas", f"{d['fecha_inicio']} a {d['fecha_termino']}")]:
            st.markdown(f"- **{k}:** {v}")

    with c2:
        st.markdown("**Tecnicos**")
        for t in d.get("tecnicos", []):
            st.markdown(f"- {t['nombre']}: **{fmt(t['monto'])}**")

        st.markdown("**Desglose de Costos**")
        for concept, val in calc["desglose"].items():
            st.markdown(f"- {concept}: **{fmt(val)}**")

    st.divider()

    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.markdown('<div class="wizard-back">', unsafe_allow_html=True)
        if st.button("<<< Atras", key="w3_back", use_container_width=True):
            st.session_state.wizard_step = 2
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="wizard-submit">', unsafe_allow_html=True)
        if st.button("GENERAR SOLICITUD", key="w3_submit", use_container_width=True):
            # Save
            save_solicitud(d)
            save_to_historial(d, ref, calc)
            st.session_state.wizard_complete = True
            st.session_state.last_calc = calc
            st.session_state.show_modal = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)


@st.dialog("Resumen de Solicitud")
def show_success_modal():
    """Modal de solicitud completada con opciones de descarga."""
    d = st.session_state.wizard_data
    ref = st.session_state.ref_data
    calc = st.session_state.get("last_calc", {})

    st.markdown(f"""
    <div style="text-align:center;">
        <h3 style="color:#64b5f6; margin-bottom: 5px;">Solicitud Generada Exitosamente</h3>
        <p style="font-size:32px; font-weight:bold; color:#ffb300; margin-top: 5px;">{fmt(calc.get('total_general', 0))}</p>
        <p style="color:rgba(180,200,230,0.5); font-size:13px; margin-top:6px;">
            {d.get('cliente', '')} | {', '.join([x for x in d.get('destinos',[]) if x])} | {d.get('fecha_solicitud', '')}
        </p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Fondo a Rendir:** <span style='color:#64b5f6; font-size:16px;'>{fmt(calc.get('total_fondo_rendir', 0))}</span>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"**Viático No Rendible:** <span style='color:#64b5f6; font-size:16px;'>{fmt(calc.get('total_viatico_no_rendible', 0))}</span>", unsafe_allow_html=True)

    st.divider()

    st.markdown("<p style='text-align:center; margin-bottom:10px; font-weight:600;'>Descargar Documentos</p>", unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        try:
            excel_bytes = generar_excel(d, calc, ref)
            st.download_button(
                "📥 Formato Excel (.xlsx)", data=excel_bytes,
                file_name=f"Solicitud_Viaticos_{d.get('cliente','').replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="modal_dl_xls", use_container_width=True
            )
        except Exception as e:
            st.error(f"Error: {e}")

    with c2:
        try:
            pdf_bytes = generar_pdf(d, calc, ref)
            st.download_button(
                "📥 Formato PDF", data=pdf_bytes,
                file_name=f"Solicitud_Viaticos_{d.get('cliente','').replace(' ','_')}.pdf",
                mime="application/pdf",
                key="modal_dl_pdf", use_container_width=True
            )
        except Exception as e:
            st.error(f"Error: {e}")

    st.divider()
    if st.button("Aceptar y Cerrar", type="primary", use_container_width=True):
        st.session_state.show_modal = False
        st.session_state.wizard_step = 1
        st.session_state.wizard_data = get_empty_solicitud()
        st.rerun()


# ═══════════════════════════════════════════════════════════
# CONFIGURACION
# ═══════════════════════════════════════════════════════════
def page_configuracion():
    st.markdown("# Configuracion")

    ref = st.session_state.ref_data

    tabs = st.tabs(["Kilometraje", "Peajes", "Combustible", "Alojamiento", "Imprevistos"])

    # ── Kilometraje ──
    with tabs[0]:
        st.markdown("### Tabla de Kilometraje (desde Santiago)")
        km_data = ref.get("kilometraje", [])
        df_km = pd.DataFrame(km_data)
        if not df_km.empty:
            df_km.columns = ["Ciudad", "Km"]
            edited = st.data_editor(df_km, hide_index=True, num_rows="dynamic", key="cfg_km", height=400)
            st.markdown('<div class="success-btn">', unsafe_allow_html=True)
            if st.button("Guardar Kilometraje", key="cfg_km_save", use_container_width=True):
                new_km = [{"ciudad": str(r["Ciudad"]), "km": float(r["Km"]) if pd.notna(r["Km"]) else 0} for _, r in edited.iterrows()]
                ref["kilometraje"] = new_km
                save_default_data(ref)
                st.success("Kilometraje guardado.")
            st.markdown('</div>', unsafe_allow_html=True)

    # ── Peajes ──
    with tabs[1]:
        st.markdown("### Peajes por Destino")
        sub1, sub2 = st.tabs(["Norte", "Sur"])
        with sub1:
            norte = ref.get("peajes_por_destino", {}).get("norte", [])
            df_n = pd.DataFrame(norte)
            if not df_n.empty:
                df_n.columns = ["Ciudad", "Auto (ida)", "Camion (ida)"]
                edited_n = st.data_editor(df_n, hide_index=True, num_rows="dynamic", key="cfg_pn", height=300)
                st.markdown('<div class="success-btn">', unsafe_allow_html=True)
                if st.button("Guardar Norte", key="cfg_pn_save", use_container_width=True):
                    ref["peajes_por_destino"]["norte"] = [
                        {"ciudad": str(r["Ciudad"]), "auto_ida": int(r["Auto (ida)"]) if pd.notna(r["Auto (ida)"]) else 0,
                         "camion_ida": int(r["Camion (ida)"]) if pd.notna(r["Camion (ida)"]) else 0}
                        for _, r in edited_n.iterrows()
                    ]
                    save_default_data(ref)
                    st.success("Peajes Norte guardados.")
                st.markdown('</div>', unsafe_allow_html=True)
        with sub2:
            sur = ref.get("peajes_por_destino", {}).get("sur", [])
            df_s = pd.DataFrame(sur)
            if not df_s.empty:
                df_s.columns = ["Ciudad", "Auto (ida)", "Camion (ida)"]
                edited_s = st.data_editor(df_s, hide_index=True, num_rows="dynamic", key="cfg_ps", height=300)
                st.markdown('<div class="success-btn">', unsafe_allow_html=True)
                if st.button("Guardar Sur", key="cfg_ps_save", use_container_width=True):
                    ref["peajes_por_destino"]["sur"] = [
                        {"ciudad": str(r["Ciudad"]), "auto_ida": int(r["Auto (ida)"]) if pd.notna(r["Auto (ida)"]) else 0,
                         "camion_ida": int(r["Camion (ida)"]) if pd.notna(r["Camion (ida)"]) else 0}
                        for _, r in edited_s.iterrows()
                    ]
                    save_default_data(ref)
                    st.success("Peajes Sur guardados.")
                st.markdown('</div>', unsafe_allow_html=True)

    # ── Combustible ──
    with tabs[2]:
        st.markdown("### Precios de Combustible")
        comb = ref.get("combustible", {})
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Auto / Camioneta")
            comb["precio_litro_auto"] = st.number_input("Precio por Litro ($)", min_value=0, value=int(comb.get("precio_litro_auto", 1450)), step=50, key="cfg_pla")
            comb["consumo_km_lt_auto"] = st.number_input("Consumo (Km/Lt)", min_value=1, value=int(comb.get("consumo_km_lt_auto", 10)), key="cfg_cka")
        with c2:
            st.markdown("#### Camion")
            comb["precio_litro_camion"] = st.number_input("Precio por Litro ($)", min_value=0, value=int(comb.get("precio_litro_camion", 1100)), step=50, key="cfg_plc")
            comb["consumo_km_lt_camion"] = st.number_input("Consumo (Km/Lt)", min_value=1, value=int(comb.get("consumo_km_lt_camion", 6)), key="cfg_ckc")
        ref["combustible"] = comb
        st.markdown('<div class="success-btn">', unsafe_allow_html=True)
        if st.button("Guardar Combustible", key="cfg_comb_save", use_container_width=True):
            save_default_data(ref)
            st.success("Combustible guardado.")
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Alojamiento ──
    with tabs[3]:
        st.markdown("### Precios de Alojamiento (por noche)")
        aloj = ref.get("alojamiento", {})
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Habitacion Doble con Bano")
            doble = aloj.get("habitacion_doble", {})
            doble["bajo"] = st.number_input("Rango Bajo ($)", min_value=0, value=int(doble.get("bajo", 20000)), step=5000, key="cfg_db")
            doble["promedio"] = st.number_input("Rango Promedio ($)", min_value=0, value=int(doble.get("promedio", 70000)), step=5000, key="cfg_dp")
            doble["alto"] = st.number_input("Rango Alto ($)", min_value=0, value=int(doble.get("alto", 60000)), step=5000, key="cfg_da")
            aloj["habitacion_doble"] = doble
        with c2:
            st.markdown("#### Habitacion Single con Bano")
            single = aloj.get("habitacion_single", {})
            single["bajo"] = st.number_input("Rango Bajo ($)", min_value=0, value=int(single.get("bajo", 0)), step=5000, key="cfg_sb")
            single["promedio"] = st.number_input("Rango Promedio ($)", min_value=0, value=int(single.get("promedio", 0)), step=5000, key="cfg_sp")
            single["alto"] = st.number_input("Rango Alto ($)", min_value=0, value=int(single.get("alto", 0)), step=5000, key="cfg_sa")
            aloj["habitacion_single"] = single
        ref["alojamiento"] = aloj
        st.markdown('<div class="success-btn">', unsafe_allow_html=True)
        if st.button("Guardar Alojamiento", key="cfg_aloj_save", use_container_width=True):
            save_default_data(ref)
            st.success("Alojamiento guardado.")
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Imprevistos ──
    with tabs[4]:
        st.markdown("### Porcentaje de Imprevistos")
        st.markdown("Se aplica sobre la suma de Peajes + Combustible + Alojamiento.")
        pct = ref.get("imprevistos_porcentaje", 0.20)
        new_pct = st.slider("Porcentaje (%)", min_value=0, max_value=50, value=int(pct * 100), step=5, key="cfg_pct")
        ref["imprevistos_porcentaje"] = new_pct / 100

        st.markdown(f"""
        <div class="result-card" style="margin:16px 0;">
            <p style="color:rgba(180,200,230,0.5); font-size:12px; text-transform:uppercase;">Porcentaje configurado</p>
            <p style="color:#64b5f6; font-size:32px; font-weight:700;">{new_pct}%</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="success-btn">', unsafe_allow_html=True)
        if st.button("Guardar Imprevistos", key="cfg_imp_save", use_container_width=True):
            save_default_data(ref)
            st.success("Imprevistos guardado.")
        st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
# ROUTER
# ═══════════════════════════════════════════════════════════
if page == "Dashboard":
    page_dashboard()
elif page == "Nueva Solicitud":
    page_nueva_solicitud()
elif page == "Configuracion":
    page_configuracion()
