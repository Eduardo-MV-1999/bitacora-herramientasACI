import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
import os

# ========== CONFIGURACIÓN ==========
st.set_page_config(page_title="Bitácora de Herramientas", page_icon="🔧", layout="wide")

st.markdown("""
<style>
    .stButton button { width: 100%; border-radius: 10px; padding: 0.5rem; font-size: 1.2rem; }
    @media (max-width: 768px) { .stDataFrame { font-size: 12px; } }
    .diagnostico { background-color: #f0f2f6; padding: 10px; border-radius: 5px; margin: 10px 0; }
</style>
""", unsafe_allow_html=True)

# ========== BASE DE DATOS ==========
DB_NAME = "herramientas.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS herramientas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE,
                    cantidad_total INTEGER,
                    descripcion TEXT
                )''')
    c.execute('''CREATE TABLE IF NOT EXISTS movimientos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    herramienta_id INTEGER,
                    usuario TEXT,
                    fecha_prestamo TEXT,
                    fecha_devolucion TEXT,
                    FOREIGN KEY(herramienta_id) REFERENCES herramientas(id)
                )''')
    conn.commit()
    conn.close()

init_db()

# ========== FUNCIONES ==========
def get_herramientas():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT id, nombre, cantidad_total FROM herramientas", conn)
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
        ORDER BY m.fecha_prestamo DESC
    """, conn)
    conn.close()
    return df

def get_stock_disponible():
    conn = sqlite3.connect(DB_NAME)
    df_herramientas = pd.read_sql("SELECT id, nombre, cantidad_total FROM herramientas", conn)
    prestamos = pd.read_sql("""
        SELECT herramienta_id, COUNT(*) as prestados
        FROM movimientos
        WHERE fecha_devolucion IS NULL
        GROUP BY herramienta_id
    """, conn)
    conn.close()
    df_herramientas['prestados'] = 0
    for _, row in prestamos.iterrows():
        df_herramientas.loc[df_herramientas['id'] == row['herramienta_id'], 'prestados'] = row['prestados']
    df_herramientas['disponible'] = df_herramientas['cantidad_total'] - df_herramientas['prestados']
    return df_herramientas

# ========== IMPORTAR/EXPORTAR ==========
def importar_dataframe(df, fuente):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    insertados = 0
    for _, row in df.iterrows():
        try:
            c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                      (row['nombre'], int(row['cantidad_total']), row.get('descripcion', '')))
            insertados += 1
        except sqlite3.IntegrityError:
            st.warning(f"⚠️ Herramienta '{row['nombre']}' ya existe, omitida")
    conn.commit()
    conn.close()
    st.success(f"✅ Importadas {insertados} herramientas desde {fuente}")
    return insertados

def exportar_excel():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT nombre, cantidad_total, descripcion FROM herramientas ORDER BY nombre", conn)
    conn.close()
    return df

def importar_desde_archivo_local():
    posibles = ["herramientas.xlsx", "herramientas.csv", "inventario.xlsx", "inventario.csv"]
    archivo = None
    for f in posibles:
        if os.path.exists(f):
            archivo = f
            break
    if archivo is None:
        st.info("📂 No se encontró archivo de inventario en la carpeta. Usa 'Subir archivo manual'.")
        return False
    try:
        if archivo.endswith('.csv'):
            df = pd.read_csv(archivo, encoding='utf-8')
        else:
            df = pd.read_excel(archivo, engine='openpyxl')
        df.columns = df.columns.str.strip().str.lower()
        if 'nombre' not in df.columns or 'cantidad_total' not in df.columns:
            st.error(f"El archivo debe tener columnas 'nombre' y 'cantidad_total'. Encontradas: {list(df.columns)}")
            return False
        df['nombre'] = df['nombre'].astype(str)
        df['cantidad_total'] = pd.to_numeric(df['cantidad_total'], errors='coerce').fillna(1).astype(int)
        if 'descripcion' not in df.columns:
            df['descripcion'] = ''
        else:
            df['descripcion'] = df['descripcion'].fillna('').astype(str)
        return importar_dataframe(df, archivo)
    except Exception as e:
        st.error(f"Error al leer {archivo}: {e}")
        return False

# ========== INICIALIZACIÓN ==========
if 'datos_cargados' not in st.session_state:
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM herramientas")
    count = c.fetchone()[0]
    conn.close()
    st.session_state.datos_cargados = count > 0

if not st.session_state.datos_cargados:
    with st.spinner("Buscando archivo de inventario..."):
        if importar_desde_archivo_local():
            st.session_state.datos_cargados = True
            st.rerun()
        else:
            st.warning("Sube tu archivo Excel/CSV para comenzar.")
            archivo = st.file_uploader("Selecciona archivo", type=['xlsx', 'xls', 'csv'])
            if archivo is not None:
                try:
                    if archivo.name.endswith('.csv'):
                        df = pd.read_csv(archivo)
                    else:
                        df = pd.read_excel(archivo, engine='openpyxl')
                    df.columns = df.columns.str.strip().str.lower()
                    if 'nombre' in df.columns and 'cantidad_total' in df.columns:
                        df['nombre'] = df['nombre'].astype(str)
                        df['cantidad_total'] = pd.to_numeric(df['cantidad_total'], errors='coerce').fillna(1).astype(int)
                        if 'descripcion' not in df.columns:
                            df['descripcion'] = ''
                        importar_dataframe(df, archivo.name)
                        st.session_state.datos_cargados = True
                        st.rerun()
                    else:
                        st.error("Columnas requeridas: 'nombre' y 'cantidad_total'")
                except Exception as e:
                    st.error(f"Error: {e}")
            st.stop()

# ========== MENÚ PRINCIPAL ==========
st.title("🔧 Bitácora de Herramientas")

menu = st.sidebar.radio(
    "📋 Menú",
    ["📊 Inventario", "➕ Nueva herramienta", "📤 Registrar préstamo", "📥 Registrar devolución", 
     "📜 Historial", "📤 Exportar inventario", "🔄 Subir archivo", "🔍 Diagnóstico"]
)

# ---------- INVENTARIO ----------
if menu == "📊 Inventario":
    st.subheader("Inventario actual")
    df_stock = get_stock_disponible()
    if df_stock.empty:
        st.info("No hay herramientas registradas.")
    else:
        if st.checkbox("Ver como tarjetas (mejor en móvil)"):
            for _, row in df_stock.iterrows():
                with st.container():
                    col1, col2, col3 = st.columns([3,1,1])
                    col1.markdown(f"**{row['nombre']}**")
                    col2.write(f"Total: {row['cantidad_total']}")
                    col3.write(f"Disponible: {row['disponible']}")
                    st.divider()
        else:
            st.dataframe(df_stock[['nombre', 'cantidad_total', 'prestados', 'disponible']], width='stretch')

# ---------- NUEVA HERRAMIENTA ----------
elif menu == "➕ Nueva herramienta":
    st.subheader("Agregar herramienta manualmente")
    with st.form("add_tool"):
        nombre = st.text_input("Nombre *")
        cantidad = st.number_input("Cantidad total", min_value=1, step=1, value=1)
        descripcion = st.text_area("Descripción")
        submitted = st.form_submit_button("Guardar")
        if submitted and nombre:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                          (nombre, cantidad, descripcion))
                conn.commit()
                st.success(f"✅ '{nombre}' agregada correctamente")
            except sqlite3.IntegrityError:
                st.error("❌ Ya existe una herramienta con ese nombre")
            conn.close()

# ---------- REGISTRAR PRÉSTAMO ----------
elif menu == "📤 Registrar préstamo":
    st.subheader("Salida de herramienta")
    df_stock = get_stock_disponible()
    disponibles = df_stock[df_stock['disponible'] > 0]
    if disponibles.empty:
        st.error("No hay herramientas disponibles")
    else:
        herramienta = st.selectbox("Herramienta", disponibles['nombre'])
        usuario = st.text_input("Nombre de quien retira *")
        fecha = st.date_input("Fecha de préstamo", date.today())
        if st.button("🔓 Registrar préstamo", use_container_width=True):
            if not usuario:
                st.error("Debes escribir el nombre")
            else:
                id_her = disponibles[disponibles['nombre'] == herramienta]['id'].values[0]
                fecha_str = fecha.isoformat()
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                try:
                    c.execute("INSERT INTO movimientos (herramienta_id, usuario, fecha_prestamo) VALUES (?, ?, ?)",
                              (id_her, usuario, fecha_str))
                    conn.commit()
                    nuevo_id = c.lastrowid
                    st.success(f"✅ Préstamo registrado con ID {nuevo_id} - {herramienta} a {usuario} el {fecha_str}")
                    st.balloons()
                    # Mostrar el registro recién insertado para depuración
                    st.info("🔍 Revisa 'Diagnóstico' para confirmar que el registro existe.")
                except Exception as e:
                    st.error(f"❌ Error al guardar: {e}")
                finally:
                    conn.close()

# ---------- REGISTRAR DEVOLUCIÓN ----------
elif menu == "📥 Registrar devolución":
    st.subheader("Devolución de herramienta")
    prestamos = get_prestamos_activos()
    if prestamos.empty:
        st.info("No hay préstamos activos en este momento.")
        # Mostrar todos los movimientos para ayudar a depurar
        with st.expander("Ver todos los movimientos (incluyendo devueltos)"):
            df_all = get_movimientos()
            st.dataframe(df_all, width='stretch')
    else:
        st.dataframe(prestamos[['id', 'nombre', 'usuario', 'fecha_prestamo']], width='stretch')
        opciones = prestamos.apply(lambda x: f"ID:{x['id']} - {x['nombre']} - {x['usuario']} ({x['fecha_prestamo']})", axis=1)
        seleccion = st.selectbox("Selecciona el préstamo a devolver", opciones)
        fecha_dev = st.date_input("Fecha de devolución", date.today())
        if st.button("🔒 Marcar devuelto", use_container_width=True):
            idx = opciones[opciones == seleccion].index[0]
            id_mov = prestamos.loc[idx, 'id']
            fecha_dev_str = fecha_dev.isoformat()
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            try:
                c.execute("UPDATE movimientos SET fecha_devolucion = ? WHERE id = ?", (fecha_dev_str, id_mov))
                conn.commit()
                st.success(f"✅ Devolución registrada para el préstamo ID {id_mov}")
                st.balloons()
            except Exception as e:
                st.error(f"❌ Error al devolver: {e}")
            finally:
                conn.close()

# ---------- HISTORIAL ----------
elif menu == "📜 Historial":
    st.subheader("Historial completo de movimientos")
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT h.nombre AS Herramienta, m.usuario AS Usuario, 
               m.fecha_prestamo AS "Fecha préstamo", 
               m.fecha_devolucion AS "Fecha devolución"
        FROM movimientos m
        JOIN herramientas h ON m.herramienta_id = h.id
        ORDER BY m.fecha_prestamo DESC
    """, conn)
    conn.close()
    if df.empty:
        st.info("No hay movimientos registrados.")
    else:
        st.dataframe(df, width='stretch')
        st.caption(f"Total de registros: {len(df)}")

# ---------- EXPORTAR INVENTARIO ----------
elif menu == "📤 Exportar inventario":
    st.subheader("Exportar inventario a Excel")
    df_export = exportar_excel()
    if df_export.empty:
        st.warning("No hay herramientas para exportar.")
    else:
        st.dataframe(df_export, width='stretch')
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Inventario')
        st.download_button(
            label="📥 Descargar inventario_actual.xlsx",
            data=output.getvalue(),
            file_name="inventario_actual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------- SUBIR ARCHIVO ----------
elif menu == "🔄 Subir archivo":
    st.subheader("Importar herramientas desde Excel/CSV")
    st.info("El archivo debe tener columnas: 'nombre' y 'cantidad_total'. Las herramientas existentes se omiten.")
    archivo = st.file_uploader("Selecciona archivo", type=['xlsx', 'xls', 'csv'])
    if archivo is not None:
        try:
            if archivo.name.endswith('.csv'):
                df = pd.read_csv(archivo)
            else:
                df = pd.read_excel(archivo, engine='openpyxl')
            df.columns = df.columns.str.strip().str.lower()
            if 'nombre' in df.columns and 'cantidad_total' in df.columns:
                df['nombre'] = df['nombre'].astype(str)
                df['cantidad_total'] = pd.to_numeric(df['cantidad_total'], errors='coerce').fillna(1).astype(int)
                if 'descripcion' not in df.columns:
                    df['descripcion'] = ''
                importar_dataframe(df, archivo.name)
                st.rerun()
            else:
                st.error("Columnas requeridas: 'nombre' y 'cantidad_total'")
        except Exception as e:
            st.error(f"Error: {e}")

# ---------- DIAGNÓSTICO ----------
elif menu == "🔍 Diagnóstico":
    st.subheader("Diagnóstico de la base de datos")
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Tabla herramientas**")
        conn = sqlite3.connect(DB_NAME)
        df_her = pd.read_sql("SELECT * FROM herramientas", conn)
        conn.close()
        st.dataframe(df_her, width='stretch')
    with col2:
        st.write("**Tabla movimientos**")
        conn = sqlite3.connect(DB_NAME)
        df_mov = pd.read_sql("SELECT * FROM movimientos", conn)
        conn.close()
        st.dataframe(df_mov, width='stretch')
    st.markdown("---")
    st.write("**Préstamos activos (fecha_devolucion IS NULL)**")
    df_activos = get_prestamos_activos()
    st.dataframe(df_activos, width='stretch')
    st.write("**Cantidad de herramientas:**", len(df_her))
    st.write("**Cantidad de movimientos:**", len(df_mov))
