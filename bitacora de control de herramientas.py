import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
import os

# Configuración de página
st.set_page_config(page_title="Bitácora de Herramientas", page_icon="🔧", layout="wide")

# CSS responsivo
st.markdown("""
<style>
    .stButton button { width: 100%; border-radius: 10px; padding: 0.5rem; font-size: 1.2rem; }
    @media (max-width: 768px) { .stDataFrame { font-size: 12px; } }
</style>
""", unsafe_allow_html=True)

# Conexión a base de datos
DB_NAME = "herramientas.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS herramientas
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  nombre TEXT UNIQUE,
                  cantidad_total INTEGER,
                  descripcion TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS movimientos
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  herramienta_id INTEGER,
                  usuario TEXT,
                  fecha_prestamo DATE,
                  fecha_devolucion DATE)''')
    conn.commit()
    conn.close()

init_db()

# Funciones de consulta
def get_herramientas():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT id, nombre, cantidad_total FROM herramientas", conn)
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
    prestamos = pd.read_sql("SELECT herramienta_id, COUNT(*) as prestados FROM movimientos WHERE fecha_devolucion IS NULL GROUP BY herramienta_id", conn)
    conn.close()
    df_herramientas['prestados'] = 0
    for _, row in prestamos.iterrows():
        df_herramientas.loc[df_herramientas['id'] == row['herramienta_id'], 'prestados'] = row['prestados']
    df_herramientas['disponible'] = df_herramientas['cantidad_total'] - df_herramientas['prestados']
    return df_herramientas

def importar_dataframe(df, nombre_archivo):
    """Inserta un DataFrame en la tabla herramientas (omite duplicados)"""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    insertados = 0
    for _, row in df.iterrows():
        try:
            c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                      (row['nombre'], row['cantidad_total'], row.get('descripcion', '')))
            insertados += 1
        except sqlite3.IntegrityError:
            st.warning(f"⚠️ Herramienta '{row['nombre']}' ya existe, se omite")
    conn.commit()
    conn.close()
    if insertados > 0:
        st.success(f"✅ Importadas {insertados} herramientas desde {nombre_archivo}")
    return insertados

def importar_desde_archivo_subido(archivo_subido):
    """Lee un archivo subido por el usuario y lo importa"""
    try:
        nombre = archivo_subido.name
        if nombre.endswith('.csv'):
            df = pd.read_csv(archivo_subido)
        else:
            df = pd.read_excel(archivo_subido, engine='openpyxl')
        
        st.write("📋 Columnas encontradas:", list(df.columns))
        
        # Mapeo flexible de columnas
        col_nombre = None
        col_cantidad = None
        col_descripcion = None
        
        for col in df.columns:
            col_lower = col.lower().strip()
            if 'nombre' in col_lower or 'herramienta' in col_lower:
                if col_nombre is None:
                    col_nombre = col
            if 'cantidad' in col_lower or 'stock' in col_lower or 'total' in col_lower:
                if col_cantidad is None:
                    col_cantidad = col
            if 'descripcion' in col_lower or 'detalle' in col_lower:
                col_descripcion = col
        
        if col_nombre is None or col_cantidad is None:
            st.error("No se pudieron identificar las columnas 'nombre' y 'cantidad_total'.")
            st.info("Usa la opción 'Subir archivo manual' y asigna las columnas correctas (próxima mejora). Por ahora, renombra las columnas en tu Excel a: 'nombre', 'cantidad_total'")
            return False
        
        df = df.rename(columns={col_nombre: 'nombre', col_cantidad: 'cantidad_total'})
        if col_descripcion:
            df = df.rename(columns={col_descripcion: 'descripcion'})
        else:
            df['descripcion'] = ''
        
        df['nombre'] = df['nombre'].astype(str)
        df['cantidad_total'] = pd.to_numeric(df['cantidad_total'], errors='coerce').fillna(1).astype(int)
        
        return importar_dataframe(df, nombre)
    except Exception as e:
        st.error(f"Error al leer el archivo: {str(e)}")
        return False

def importar_desde_archivo_local():
    """Busca automáticamente herramientas.xlsx o .csv en la carpeta actual"""
    # Posibles nombres
    posibles = ["herramientas.xlsx", "herramientas.csv", "inventario.xlsx", "inventario.csv"]
    archivo_encontrado = None
    for nombre in posibles:
        if os.path.exists(nombre):
            archivo_encontrado = nombre
            break
    
    if archivo_encontrado is None:
        st.info("📂 No se encontró un archivo 'herramientas.xlsx' o 'herramientas.csv' en la carpeta del proyecto. Puedes subir uno manualmente desde el menú.")
        return False
    
    try:
        if archivo_encontrado.endswith('.csv'):
            df = pd.read_csv(archivo_encontrado, encoding='utf-8')
        else:
            df = pd.read_excel(archivo_encontrado, engine='openpyxl')
        
        # Verificar columnas necesarias (exactas o flexibles)
        columnas_originales = list(df.columns)
        # Estandarizar nombres de columna: eliminar espacios y pasar a minúsculas
        df.columns = df.columns.str.strip().str.lower()
        
        if 'nombre' not in df.columns or 'cantidad_total' not in df.columns:
            st.error(f"El archivo debe tener columnas 'nombre' y 'cantidad_total'. Encontradas: {columnas_originales}")
            st.info("Renombra las columnas en tu Excel a: 'nombre', 'cantidad_total' (sin espacios, minúsculas)")
            return False
        
        # Asegurar tipo
        df['nombre'] = df['nombre'].astype(str)
        df['cantidad_total'] = pd.to_numeric(df['cantidad_total'], errors='coerce').fillna(1).astype(int)
        if 'descripcion' not in df.columns:
            df['descripcion'] = ''
        else:
            df['descripcion'] = df['descripcion'].fillna('').astype(str)
        
        return importar_dataframe(df, archivo_encontrado)
    except Exception as e:
        st.error(f"Error al leer el archivo local {archivo_encontrado}: {str(e)}")
        st.info("Asegúrate de tener instalado 'openpyxl': pip install openpyxl")
        return False

# --- Interfaz principal ---
st.title("🔧 Bitácora de Herramientas")

# Inicializar estado de sesión
if 'datos_cargados' not in st.session_state:
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM herramientas")
    count = c.fetchone()[0]
    conn.close()
    st.session_state.datos_cargados = count > 0

# Si no hay datos, intentar importar automáticamente desde archivo local
if not st.session_state.datos_cargados:
    with st.spinner("Buscando archivo de inventario en la carpeta..."):
        if importar_desde_archivo_local():
            st.session_state.datos_cargados = True
            st.rerun()
        else:
            st.warning("📂 No hay herramientas en la base de datos. Sube tu archivo Excel/CSV para comenzar.")
            archivo = st.file_uploader("Selecciona tu archivo de inventario (Excel o CSV)", type=['xlsx', 'xls', 'csv'])
            if archivo is not None:
                if importar_desde_archivo_subido(archivo):
                    st.session_state.datos_cargados = True
                    st.rerun()
            st.stop()  # No mostrar el resto hasta que haya datos

# Menú principal
menu = st.sidebar.radio(
    "📋 Menú",
    ["📊 Inventario", "➕ Nueva herramienta", "📤 Registrar préstamo", "📥 Registrar devolución", "📜 Historial", "🔄 Subir archivo manual"]
)

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
            st.dataframe(df_stock[['nombre', 'cantidad_total', 'prestados', 'disponible']], use_container_width=True)

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
                st.success(f"✅ {nombre} agregada")
            except sqlite3.IntegrityError:
                st.error("Ya existe una herramienta con ese nombre")
            conn.close()

elif menu == "📤 Registrar préstamo":
    st.subheader("Salida de herramienta")
    df_stock = get_stock_disponible()
    disponibles = df_stock[df_stock['disponible'] > 0]
    if disponibles.empty:
        st.error("No hay herramientas disponibles")
    else:
        herramienta = st.selectbox("Herramienta", disponibles['nombre'])
        usuario = st.text_input("Nombre de quien retira")
        fecha = st.date_input("Fecha", date.today())
        if st.button("Registrar préstamo"):
            if usuario:
                id_her = disponibles[disponibles['nombre']==herramienta]['id'].values[0]
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                c.execute("INSERT INTO movimientos (herramienta_id, usuario, fecha_prestamo) VALUES (?, ?, ?)",
                          (id_her, usuario, fecha))
                conn.commit()
                conn.close()
                st.success("Préstamo registrado")
                st.balloons()
            else:
                st.error("Escribe el nombre")

elif menu == "📥 Registrar devolución":
    st.subheader("Devolución")
    prestamos = get_prestamos_activos()
    if prestamos.empty:
        st.info("No hay préstamos activos")
    else:
        opciones = prestamos.apply(lambda x: f"{x['nombre']} - {x['usuario']} ({x['fecha_prestamo']})", axis=1)
        seleccion = st.selectbox("Préstamo a devolver", opciones)
        fecha_dev = st.date_input("Fecha devolución", date.today())
        if st.button("Marcar devuelto"):
            idx = opciones[opciones == seleccion].index[0]
            id_mov = prestamos.loc[idx, 'id']
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("UPDATE movimientos SET fecha_devolucion = ? WHERE id = ?", (fecha_dev, id_mov))
            conn.commit()
            conn.close()
            st.success("Devolución registrada")
            st.balloons()

elif menu == "📜 Historial":
    st.subheader("Historial completo")
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT h.nombre AS Herramienta, m.usuario AS Usuario, m.fecha_prestamo AS "Fecha préstamo", m.fecha_devolucion AS "Fecha devolución"
        FROM movimientos m JOIN herramientas h ON m.herramienta_id = h.id
        ORDER BY m.fecha_prestamo DESC
    """, conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif menu == "🔄 Subir archivo manual":
    st.subheader("Importar desde archivo (agrega nuevas herramientas)")
    st.info("El archivo debe tener columnas 'nombre' y 'cantidad_total' (sin espacios, minúsculas). Las herramientas existentes se omitirán.")
    archivo = st.file_uploader("Selecciona Excel o CSV", type=['xlsx', 'xls', 'csv'])
    if archivo is not None:
        if importar_desde_archivo_subido(archivo):
            st.rerun()