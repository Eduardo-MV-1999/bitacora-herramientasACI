import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
import os
from io import BytesIO
import hashlib

# ========== CONFIGURACIÓN ==========
st.set_page_config(page_title="Bitácora de Herramientas", page_icon="🔧", layout="wide")

# ========== AUTENTICACIÓN ==========
# Definir usuario y contraseña (en producción, usa st.secrets)
USUARIO_VALIDO = "AUTOMATIK"
CONTRASENA_VALIDA = "ACI940930L91"  # Cambia esto por una contraseña segura

def verificar_credenciales(usuario, contrasena):
    # En producción, compara con hash almacenado en secrets
    return usuario == USUARIO_VALIDO and contrasena == CONTRASENA_VALIDA

def login():
    st.sidebar.title("🔐 Acceso")
    usuario = st.sidebar.text_input("Usuario")
    contrasena = st.sidebar.text_input("Contraseña", type="password")
    if st.sidebar.button("Iniciar sesión"):
        if verificar_credenciales(usuario, contrasena):
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.sidebar.error("Usuario o contraseña incorrectos")
    st.sidebar.info("Contacta al administrador si no tienes acceso.")

# ========== INICIALIZACIÓN DE SESIÓN ==========
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# Mostrar login si no está autenticado
if not st.session_state.autenticado:
    login()
    st.stop()  # No mostrar nada más

# ---------- A partir de aquí, la aplicación completa ----------

# ========== BASE DE DATOS ==========
DB_NAME = "herramientas.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # Tabla herramientas
    c.execute('''CREATE TABLE IF NOT EXISTS herramientas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE,
                    cantidad_total INTEGER,
                    descripcion TEXT
                )''')
    # Tabla movimientos
    c.execute('''CREATE TABLE IF NOT EXISTS movimientos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    herramienta_id INTEGER,
                    usuario TEXT,
                    cantidad INTEGER DEFAULT 1,
                    fecha_prestamo TEXT,
                    fecha_devolucion TEXT,
                    FOREIGN KEY(herramienta_id) REFERENCES herramientas(id)
                )''')
    # Tabla usuarios
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE
                )''')
    # Insertar usuarios por defecto
    c.execute("SELECT COUNT(*) FROM usuarios")
    if c.fetchone()[0] == 0:
        usuarios_default = ["emv", "lalo", "ryy", "juan", "maria"]
        for u in usuarios_default:
            c.execute("INSERT INTO usuarios (nombre) VALUES (?)", (u,))
    conn.commit()
    conn.close()

init_db()

# ========== FUNCIONES DE BASE DE DATOS ==========
def get_herramientas():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT id, nombre, cantidad_total, descripcion FROM herramientas", conn)
    conn.close()
    return df

def get_movimientos():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT * FROM movimientos ORDER BY id DESC", conn)
    conn.close()
    return df

def get_usuarios():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT id, nombre FROM usuarios ORDER BY nombre", conn)
    conn.close()
    return df

def agregar_usuario(nombre):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("INSERT INTO usuarios (nombre) VALUES (?)", (nombre,))
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        st.error(f"❌ El usuario '{nombre}' ya existe")
        return False
    except Exception as e:
        st.error(f"Error al agregar usuario: {e}")
        return False

def eliminar_usuario(id_usuario):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM usuarios WHERE id = ?", (id_usuario,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error al eliminar usuario: {e}")
        return False

def get_stock_disponible():
    conn = sqlite3.connect(DB_NAME)
    df_herramientas = pd.read_sql("SELECT id, nombre, cantidad_total FROM herramientas", conn)
    prestamos = pd.read_sql("""
        SELECT herramienta_id, SUM(cantidad) as prestados
        FROM movimientos
        WHERE fecha_devolucion IS NULL
        GROUP BY herramienta_id
    """, conn)
    conn.close()
    
    df_herramientas['prestados'] = 0
    if not prestamos.empty:
        for _, row in prestamos.iterrows():
            df_herramientas.loc[df_herramientas['id'] == row['herramienta_id'], 'prestados'] = row['prestados']
    df_herramientas['disponible'] = df_herramientas['cantidad_total'] - df_herramientas['prestados']
    return df_herramientas

def eliminar_herramienta(id_herramienta):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM movimientos WHERE herramienta_id = ?", (id_herramienta,))
        c.execute("DELETE FROM herramientas WHERE id = ?", (id_herramienta,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error en la base de datos: {e}")
        return False

def actualizar_herramienta(id_herramienta, nombre, cantidad, descripcion):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("UPDATE herramientas SET nombre = ?, cantidad_total = ?, descripcion = ? WHERE id = ?",
                  (nombre, cantidad, descripcion, id_herramienta))
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        st.error("❌ Ya existe otra herramienta con ese nombre")
        return False
    except Exception as e:
        st.error(f"Error en la base de datos: {e}")
        return False

def agregar_herramienta(nombre, cantidad, descripcion):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                  (nombre, cantidad, descripcion))
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        st.error("❌ Ya existe una herramienta con ese nombre")
        return False
    except Exception as e:
        st.error(f"Error en la base de datos: {e}")
        return False

# ========== IMPORTAR/EXPORTAR EXCEL ==========
def exportar_excel():
    df = get_herramientas()
    if df.empty:
        return None
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df[['nombre', 'cantidad_total', 'descripcion']].to_excel(writer, index=False, sheet_name='Inventario')
    return output.getvalue()

def importar_excel(archivo):
    try:
        if archivo.name.endswith('.csv'):
            df = pd.read_csv(archivo)
        else:
            df = pd.read_excel(archivo, engine='openpyxl')
        df.columns = df.columns.str.strip().str.lower()
        if 'nombre' not in df.columns or 'cantidad_total' not in df.columns:
            st.error("El archivo debe tener columnas 'nombre' y 'cantidad_total'")
            return False
        df['nombre'] = df['nombre'].astype(str)
        df['cantidad_total'] = pd.to_numeric(df['cantidad_total'], errors='coerce').fillna(1).astype(int)
        if 'descripcion' not in df.columns:
            df['descripcion'] = ''
        else:
            df['descripcion'] = df['descripcion'].fillna('').astype(str)
        
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        insertados = 0
        for _, row in df.iterrows():
            try:
                c.execute("INSERT INTO herramientas (nombre, cantidad_total, descripcion) VALUES (?, ?, ?)",
                          (row['nombre'], row['cantidad_total'], row['descripcion']))
                insertados += 1
            except sqlite3.IntegrityError:
                st.warning(f"⚠️ '{row['nombre']}' ya existe, omitida")
        conn.commit()
        conn.close()
        st.success(f"✅ Importadas {insertados} herramientas")
        return True
    except Exception as e:
        st.error(f"Error al leer archivo: {e}")
        return False

# ========== INICIALIZACIÓN ==========
if 'datos_cargados' not in st.session_state:
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM herramientas")
    count = c.fetchone()[0]
    conn.close()
    st.session_state.datos_cargados = count > 0

if 'editando_id' not in st.session_state:
    st.session_state.editando_id = None

# ========== MENÚ PRINCIPAL ==========
st.title("🔧 Bitácora de Herramientas")

menu = st.sidebar.radio(
    "📋 Menú",
    ["📊 Inventario", "📤 Registrar préstamo", "📥 Registrar devolución", "📜 Historial", "🔍 Diagnóstico"]
)

# ---------- INVENTARIO ----------
if menu == "📊 Inventario":
    st.subheader("Inventario actual")
    
    # EXPANDER: Importar / Exportar
    with st.expander("📂 Importar / Exportar inventario", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Exportar**")
            df_export = get_herramientas()
            if not df_export.empty:
                excel_data = exportar_excel()
                if excel_data:
                    st.download_button(
                        label="📥 Descargar inventario.xlsx",
                        data=excel_data,
                        file_name="inventario_actual.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("No hay herramientas para exportar")
        with col2:
            st.write("**Importar**")
            archivo = st.file_uploader("Selecciona Excel o CSV", type=['xlsx', 'xls', 'csv'])
            if archivo is not None:
                if importar_excel(archivo):
                    st.rerun()
    
    # EXPANDER: Gestión de usuarios
    with st.expander("👥 Gestionar usuarios", expanded=False):
        st.write("Usuarios registrados:")
        df_usuarios = get_usuarios()
        if df_usuarios.empty:
            st.info("No hay usuarios registrados.")
        else:
            for _, row in df_usuarios.iterrows():
                col1, col2 = st.columns([4, 1])
                col1.write(f"👤 {row['nombre']}")
                if col2.button("🗑️", key=f"del_user_{row['id']}"):
                    if eliminar_usuario(row['id']):
                        st.success(f"Usuario '{row['nombre']}' eliminado")
                        st.rerun()
        st.divider()
        st.write("**Agregar nuevo usuario:**")
        nuevo_usuario = st.text_input("Nombre del usuario")
        if st.button("Agregar usuario"):
            if nuevo_usuario:
                if agregar_usuario(nuevo_usuario):
                    st.success(f"✅ Usuario '{nuevo_usuario}' agregado")
                    st.rerun()
            else:
                st.error("Escribe un nombre")
    
    # FORMULARIO: Agregar nueva herramienta
    with st.expander("➕ Agregar nueva herramienta", expanded=False):
        with st.form(key="agregar_form"):
            col1, col2 = st.columns(2)
            with col1:
                nuevo_nombre = st.text_input("Nombre *")
            with col2:
                nueva_cantidad = st.number_input("Cantidad total", min_value=1, step=1, value=1)
            nueva_descripcion = st.text_area("Descripción")
            agregar_btn = st.form_submit_button("Agregar herramienta")
            if agregar_btn and nuevo_nombre:
                if agregar_herramienta(nuevo_nombre, nueva_cantidad, nueva_descripcion):
                    st.success(f"✅ '{nuevo_nombre}' agregada correctamente")
                    st.rerun()
    
    # TABLA DE HERRAMIENTAS
    df_herramientas = get_herramientas()
    if df_herramientas.empty:
        st.info("No hay herramientas registradas.")
    else:
        # Edición
        if st.session_state.editando_id is not None:
            editar_id = st.session_state.editando_id
            df_edit = df_herramientas[df_herramientas['id'] == editar_id]
            if not df_edit.empty:
                st.info(f"✏️ Editando: **{df_edit.iloc[0]['nombre']}**")
                with st.form(key="editar_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        nuevo_nombre = st.text_input("Nombre", value=df_edit.iloc[0]['nombre'])
                    with col2:
                        nueva_cantidad = st.number_input("Cantidad total", min_value=1, step=1, 
                                                         value=int(df_edit.iloc[0]['cantidad_total']))
                    nueva_descripcion = st.text_area("Descripción", value=df_edit.iloc[0]['descripcion'] if df_edit.iloc[0]['descripcion'] else "")
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        guardar_btn = st.form_submit_button("💾 Guardar cambios")
                    with col_btn2:
                        cancelar_btn = st.form_submit_button("❌ Cancelar")
                    
                    if guardar_btn:
                        if actualizar_herramienta(editar_id, nuevo_nombre, nueva_cantidad, nueva_descripcion):
                            st.success(f"✅ Herramienta actualizada correctamente")
                            st.session_state.editando_id = None
                            st.rerun()
                    if cancelar_btn:
                        st.session_state.editando_id = None
                        st.rerun()
        
        # Mostrar tabla con botones
        df_stock = get_stock_disponible()
        for idx, row in df_herramientas.iterrows():
            with st.container():
                col1, col2, col3, col4, col5, col6 = st.columns([3, 1.5, 1.5, 1.2, 1.2, 1.2])
                col1.markdown(f"**{row['nombre']}**")
                disponible = df_stock[df_stock['id'] == row['id']]['disponible'].values[0] if not df_stock.empty else row['cantidad_total']
                col2.write(f"Total: {row['cantidad_total']}")
                col3.write(f"📦 Disp: {disponible}")
                if row['descripcion']:
                    col4.write(f"📝 {row['descripcion'][:20]}")
                else:
                    col4.write("")
                
                if st.session_state.editando_id is None or st.session_state.editando_id == row['id']:
                    if col5.button("✏️", key=f"edit_{row['id']}"):
                        st.session_state.editando_id = row['id']
                        st.rerun()
                else:
                    col5.write("")
                
                if col6.button("🗑️", key=f"del_{row['id']}"):
                    conn = sqlite3.connect(DB_NAME)
                    c = conn.cursor()
                    c.execute("SELECT SUM(cantidad) FROM movimientos WHERE herramienta_id = ? AND fecha_devolucion IS NULL", (row['id'],))
                    prestado = c.fetchone()[0] or 0
                    conn.close()
                    
                    if prestado > 0:
                        st.warning(f"⚠️ Herramienta tiene {prestado} unidad(es) prestadas. ¿Eliminar también los movimientos?")
                        if st.button("Sí, eliminar todo", key=f"confirm_del_{row['id']}"):
                            if eliminar_herramienta(row['id']):
                                st.success(f"✅ Herramienta y sus movimientos eliminados")
                                st.rerun()
                    else:
                        if eliminar_herramienta(row['id']):
                            st.success(f"✅ Herramienta eliminada")
                            st.rerun()
                st.divider()

# ---------- REGISTRAR PRÉSTAMO ----------
elif menu == "📤 Registrar préstamo":
    st.subheader("Salida de herramienta")
    
    try:
        df_stock = get_stock_disponible()
        disponibles = df_stock[df_stock['disponible'] > 0]
        
        if disponibles.empty:
            st.error("No hay herramientas disponibles")
        else:
            herramienta = st.selectbox("Selecciona herramienta", disponibles['nombre'])
            fila = disponibles[disponibles['nombre'] == herramienta].iloc[0]
            id_her = int(fila['id'])
            disponible = int(fila['disponible'])
            
            st.info(f"Disponible: **{disponible}** unidades")
            
            # Usuarios
            df_usuarios = get_usuarios()
            lista_usuarios = df_usuarios['nombre'].tolist() if not df_usuarios.empty else []
            opciones_usuario = lista_usuarios + ["➕ Otro (nuevo)"]
            
            usuario_seleccionado = st.selectbox("Nombre de quien retira", opciones_usuario)
            
            usuario = ""
            if usuario_seleccionado == "➕ Otro (nuevo)":
                usuario = st.text_input("Escribe el nombre del nuevo usuario *")
            else:
                usuario = usuario_seleccionado
            
            cantidad = st.number_input("Cantidad a retirar", min_value=1, max_value=disponible, value=1, step=1)
            fecha = st.date_input("Fecha de préstamo", date.today())
            
            if st.button("🔓 Registrar préstamo", use_container_width=True):
                if not usuario:
                    st.error("Debes escribir el nombre del usuario")
                elif cantidad < 1:
                    st.error("La cantidad debe ser al menos 1")
                else:
                    if usuario_seleccionado == "➕ Otro (nuevo)":
                        if not agregar_usuario(usuario):
                            st.stop()
                    
                    fecha_str = fecha.isoformat()
                    conn = sqlite3.connect(DB_NAME)
                    c = conn.cursor()
                    try:
                        c.execute("INSERT INTO movimientos (herramienta_id, usuario, cantidad, fecha_prestamo) VALUES (?, ?, ?, ?)",
                                  (id_her, usuario, cantidad, fecha_str))
                        conn.commit()
                        st.success(f"✅ Préstamo registrado: {cantidad} unidad(es) de {herramienta} a {usuario}")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Error al guardar: {e}")
                    finally:
                        conn.close()
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")

# ---------- REGISTRAR DEVOLUCIÓN ----------
elif menu == "📥 Registrar devolución":
    st.subheader("Devolución de herramienta")
    
    try:
        conn = sqlite3.connect(DB_NAME)
        prestamos = pd.read_sql("""
            SELECT m.id, h.nombre, m.usuario, m.cantidad, m.fecha_prestamo
            FROM movimientos m
            JOIN herramientas h ON m.herramienta_id = h.id
            WHERE m.fecha_devolucion IS NULL
            ORDER BY m.fecha_prestamo DESC
        """, conn)
        conn.close()
        
        if prestamos.empty:
            st.info("No hay préstamos activos en este momento.")
        else:
            st.dataframe(prestamos[['id', 'nombre', 'usuario', 'cantidad', 'fecha_prestamo']], use_container_width=True)
            
            opciones = prestamos.apply(lambda x: f"ID:{x['id']} - {x['nombre']} ({x['cantidad']} unid) - {x['usuario']}", axis=1)
            seleccion = st.selectbox("Selecciona el préstamo a devolver", opciones)
            fecha_dev = st.date_input("Fecha de devolución", date.today())
            
            if st.button("🔒 Marcar devuelto", use_container_width=True):
                idx = opciones[opciones == seleccion].index[0]
                id_mov = prestamos.loc[idx, 'id']
                fecha_dev_str = fecha_dev.isoformat()
                conn = sqlite3.connect(DB_NAME)
                c = conn.cursor()
                try:
                    c.execute("UPDATE movimientos SET fecha_devolucion = ? WHERE id = ?", (fecha_dev_str, int(id_mov)))
                    conn.commit()
                    st.success(f"✅ Devolución registrada")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ Error al devolver: {e}")
                finally:
                    conn.close()
    except Exception as e:
        st.error(f"Error al cargar devoluciones: {e}")

# ---------- HISTORIAL ----------
elif menu == "📜 Historial":
    st.subheader("Historial completo de movimientos")
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("""
        SELECT h.nombre AS Herramienta, m.usuario AS Usuario, 
               m.cantidad AS Cantidad,
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
        st.dataframe(df, use_container_width=True)
        st.caption(f"Total de registros: {len(df)}")

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
        st.dataframe(df_her, use_container_width=True)
        st.write(f"Total: {len(df_her)}")
    with col2:
        st.write("**Tabla movimientos**")
        conn = sqlite3.connect(DB_NAME)
        df_mov = pd.read_sql("SELECT * FROM movimientos", conn)
        conn.close()
        st.dataframe(df_mov, use_container_width=True)
        st.write(f"Total: {len(df_mov)}")
    st.markdown("---")
    st.write("**Tabla usuarios**")
    conn = sqlite3.connect(DB_NAME)
    df_usu = pd.read_sql("SELECT * FROM usuarios", conn)
    conn.close()
    st.dataframe(df_usu, use_container_width=True)
    st.write(f"Total: {len(df_usu)}")
    st.markdown("---")
    st.write("**Ubicación de la base de datos**")
    st.code(f"{os.path.abspath(DB_NAME)}")