import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
import os

# Configuración
st.set_page_config(page_title="Bitácora", page_icon="🔧", layout="wide")

st.markdown("""
<style>
    .stButton button { width: 100%; border-radius: 10px; padding: 0.5rem; }
</style>
""", unsafe_allow_html=True)

DB_NAME = "herramientas.db"

# Inicializar BD
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS herramientas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE,
                    cantidad_total INTEGER,
                    descripcion TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS movimientos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    herramienta_id INTEGER,
                    usuario TEXT,
                    fecha_prestamo TEXT,
                    fecha_devolucion TEXT)''')
    conn.commit()
    conn.close()
init_db()

# Funciones
def get_herramientas():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT * FROM herramientas", conn)
    conn.close()
    return df

def get_movimientos():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT * FROM movimientos ORDER BY id DESC", conn)
    conn.close()
    return df

def get_prestamos_activos():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT m.id, h.nombre, m.usuario, m.fecha_prestamo
        FROM movimientos m
        JOIN herramientas h ON m.herramienta_id = h.id
        WHERE m.fecha_devolucion IS NULL
    """, conn)
    conn.close()
    return df

# Importar Excel automáticamente si está vacío
if len(get_herramientas()) == 0:
    archivos = ["herramientas.xlsx", "herramientas.csv", "inventario.xlsx"]
    for arch in archivos:
        if os.path.exists(arch):
            try:
                if arch.endswith('.csv'):
                    df = pd.read_csv(arch)
                else:
                    df = pd.read_excel(arch, engine='openpyxl')
                df.columns = df.columns.str.strip().str.lower()
                if 'nombre' in df.columns and 'cantidad_total' in df.columns:
                    conn = sqlite3.connect(DB_NAME)
                    c = conn.cursor()
                    for _, row in df.iterrows():
                        try:
                            c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                                      (row['nombre'], int(row['cantidad_total']), row.get('descripcion', '')))
                        except:
                            pass
                    conn.commit()
                    conn.close()
                    st.success(f"✅ Importado {arch}")
                    break
            except:
                pass

# INTERFAZ
st.title("🔧 Bitácora")

menu = st.sidebar.radio("Menú", ["➕ Agregar", "📤 Préstamo", "📥 Devolución", "📊 Inventario", "🔍 Diagnóstico"])

# ----------------------------------------
if menu == "➕ Agregar":
    st.subheader("Agregar herramienta")
    with st.form("add"):
        nom = st.text_input("Nombre")
        cant = st.number_input("Cantidad", min_value=1, value=1)
        desc = st.text_area("Descripción")
        if st.form_submit_button("Guardar"):
            if nom:
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                try:
                    c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                              (nom, cant, desc))
                    conn.commit()
                    st.success(f"✅ {nom} agregada")
                except:
                    st.error("Ya existe")
                conn.close()

# ----------------------------------------
elif menu == "📤 Préstamo":
    st.subheader("Registrar préstamo")
    herramientas = get_herramientas()
    if herramientas.empty:
        st.warning("No hay herramientas")
    else:
        # Calcular disponibles
        df_stock = get_herramientas()
        prestamos = get_movimientos()
        prestados = prestamos[prestamos['fecha_devolucion'].isnull()]
        prestados_count = prestados.groupby('herramienta_id').size().reset_index(name='prestados')
        df_stock = df_stock.merge(prestados_count, left_on='id', right_on='herramienta_id', how='left')
        df_stock['prestados'] = df_stock['prestados'].fillna(0).astype(int)
        df_stock['disponible'] = df_stock['cantidad_total'] - df_stock['prestados']
        disponibles = df_stock[df_stock['disponible'] > 0]
        
        if disponibles.empty:
            st.error("No hay herramientas disponibles")
        else:
            herramienta = st.selectbox("Herramienta", disponibles['nombre'])
            usuario = st.text_input("Nombre de quien retira")
            fecha = st.date_input("Fecha", date.today())
            if st.button("🔓 Registrar préstamo", use_container_width=True):
                if not usuario:
                    st.error("Escribe el nombre")
                else:
                    id_her = disponibles[disponibles['nombre'] == herramienta]['id'].values[0]
                    conn = sqlite3.connect(DB_NAME)
                    c = conn.cursor()
                    try:
                        c.execute("INSERT INTO movimientos (herramienta_id, usuario, fecha_prestamo) VALUES (?, ?, ?)",
                                  (id_her, usuario, fecha.isoformat()))
                        conn.commit()
                        st.success(f"✅ Préstamo registrado (ID: {c.lastrowid})")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Error: {e}")
                    conn.close()

# ----------------------------------------
elif menu == "📥 Devolución":
    st.subheader("Devolver herramienta")
    prestamos_activos = get_prestamos_activos()
    if prestamos_activos.empty:
        st.info("No hay préstamos activos")
        # Mostrar todos los movimientos para depuración
        with st.expander("Ver todos los movimientos"):
            st.dataframe(get_movimientos())
    else:
        st.dataframe(prestamos_activos)
        opciones = prestamos_activos.apply(lambda x: f"ID:{x['id']} - {x['nombre']} - {x['usuario']}", axis=1)
        seleccion = st.selectbox("Préstamo a devolver", opciones)
        fecha_dev = st.date_input("Fecha devolución", date.today())
        if st.button("🔒 Marcar devuelto", use_container_width=True):
            idx = opciones[opciones == seleccion].index[0]
            id_mov = prestamos_activos.loc[idx, 'id']
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("UPDATE movimientos SET fecha_devolucion = ? WHERE id = ?", (fecha_dev.isoformat(), id_mov))
                conn.commit()
                st.success(f"✅ Devuelto ID {id_mov}")
                st.balloons()
            except Exception as e:
                st.error(f"Error: {e}")
            conn.close()

# ----------------------------------------
elif menu == "📊 Inventario":
    st.subheader("Inventario")
    df = get_herramientas()
    st.dataframe(df)

# ----------------------------------------
elif menu == "🔍 Diagnóstico":
    st.subheader("Diagnóstico")
    st.write("**Herramientas:**")
    st.dataframe(get_herramientas())
    st.write("**Movimientos:**")
    st.dataframe(get_movimientos())
    st.write("**Préstamos activos:**")
    st.dataframe(get_prestamos_activos())
