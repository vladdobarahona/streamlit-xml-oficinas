import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import tempfile
import re
from io import BytesIO

# Fondo personalizado y fuente
st.markdown("""
<style>
    body {
        background-color: #abbe4c;
        font-family: 'Handel Gothic', 'Frutiger light - Roman';
    }
    .stApp {
        background-color: rgb(255, 255, 255);
        font-family: 'Frutiger Bold', sans-serif;
    }
</style>
    """, unsafe_allow_html=True)

# Cargar el archivo Excel desde el archivo local
df = pd.read_excel("oficinas.xlsx",  engine="openpyxl", dtype=str)

# Convertir el DataFrame a un archivo Excel en memoria
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Oficinas')
    output.seek(0)
    return output

excel_file = to_excel(df)

# Botón de descarga directo
st.download_button(
    label="Descargar plantilla Excel",
    data=excel_file,
    file_name="plantilla_creación_oficinas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    icon=":material/download:"
)
st.divider()
# Logo a la izquierda y título a la derecha
col1, col2 = st.columns([1, 2])
with col1:
    st.image('https://www.finagro.com.co/sites/default/files/logo-front-finagro.png', width=200)
with col2:
    st.markdown(
        '<h1 style="color: rgb(120,154,61); font-size: 2.25rem; font-weight: bold;">Generador de XML de Oficinas a partir de un archivo Excel</h1>',
        unsafe_allow_html=True
    )

# Cargar archivo DIVIPOLA fijo desde el repositorio
divipola = pd.read_excel("divipola.xlsx", sheet_name='Sheet1', engine="openpyxl", dtype=str)

# Subida del archivo de oficinas
##oficinas_file = st.file_uploader("Sube el archivo de oficinas (Excel)", type=["xlsx"])
st.markdown(
    '<span style="color: rgb(120, 154, 61); font-size: 22px;">Sube el archivo de oficinas (Excel)</span>',
    unsafe_allow_html=True
)
oficinas_file = st.file_uploader("", type=["xlsx"])

if oficinas_file:
    if st.button("Generar XML"):
        #oficinas = pd.read_excel(oficinas_file, sheet_name='Hoja1', engine="openpyxl", dtype=str)
        oficinas = pd.read_excel(oficinas_file, engine="openpyxl", dtype=str)
        oficinas['CODIGO_DEPARTAMENTO'] = oficinas['CODIGO DEL DEPARTAMENTO '].astype('int')
        oficinas['CODIGO_MUNICIPIO'] = oficinas['CODIGO DEL MUNICIPIO '].astype('int')
        divipola['CODIGO_DEPARTAMENTO_ORIGINAL'] = divipola['CODIGO_DEPARTAMENTO']
        divipola['CODIGO_MUNICIPIO_ORIGINAL'] = divipola['CODIGO_MUNICIPIO']
        divipola['CODIGO_DEPARTAMENTO'] = divipola['CODIGO_DEPARTAMENTO'].astype('int')
        divipola['CODIGO_MUNICIPIO'] = divipola['CODIGO_MUNICIPIO'].astype('int')

        oficinas = pd.merge(oficinas, divipola, how='left', on=['CODIGO_DEPARTAMENTO', 'CODIGO_MUNICIPIO'])
        Cantidad_oficinas = str(len(oficinas))

        #ET.register_namespace('', "http://www.finagro.com.co/sit")
        ET.register_namespace('sit', "http://www.finagro.com.co/sit")
        ET.register_namespace('xsi', "http://www.w3.org/2001/XMLSchema-instance")
        root_attributes = {
                            "cifraDeControl": str(Cantidad_oficinas),
                            # El atributo xsi:schemaLocation debe incluir su URI de espacio de nombres completo
                            # para que ElementTree lo maneje correctamente como un atributo calificado.
                            "{http://www.w3.org/2001/XMLSchema-instance}schemaLocation": "http://www.finagro.com.co/sit sucursales.xsd "
                        }
        #sucursales = ET.Element("{http://www.finagro.com.co/sit}sit:sucursales", cifraDeControl=cantidad_oficinas)
        sucursales = ET.Element("{http://www.finagro.com.co/sit}sucursales", **root_attributes)


        for _, row in oficinas.iterrows():
            sucursal = ET.SubElement(sucursales, "{http://www.finagro.com.co/sit}sit:sucursal")
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:codigoIntermediario").text = str(row.get('CODIGO DEL INTERMEDIARIO FINANCIERO', '')).strip()
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}codigoIdentificacionSucursal").text = str(row.get('CODIGO DE LA OFICINA', '')).strip()
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}nombre").text = str(row.get('NOMBRE DE LA OFICINA', '')).strip()
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}codigoDepartamento").text = str(row.get('CODIGO_DEPARTAMENTO_ORIGINAL', '')).strip()
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}codigoMunicipio").text = str(row.get('CODIGO_DPTO_MPIO', '')).strip()
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}direccion").text = str(row.get('DIRECCION DE LA OFICINA ', '')).strip()
            
            # Crear numeroTelefonoFijo con sus atributos
            numeroTelefonoFijo = ET.SubElement(
                sucursal,
                "{http://www.finagro.com.co/sit}numeroTelefonoFijo",
                extension='',
                prefijoCiudad=str(row.get('PREFIJO TELEFONICO DEL MUNICIPIO ', '')).strip()
            )
            numeroTelefonoFijo.text = str(row.get('NUMERO TELEFONICO DE LA OFICINA 1 ', '')).strip()
        
            # Crear numeroTelefonoFax con sus atributos.
            # Establecer el texto como una cadena vacía para que coincida con la salida deseada (sin espacio).
            numeroTelefonoFax = ET.SubElement(
                sucursal,
                "{http://www.finagro.com.co/sit}numeroTelefonoFax",
                prefijoCiudad=""
            )
            numeroTelefonoFax.text = "" # Esto asegura una etiqueta vacía, no una etiqueta con un espacio
        
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}correoElectronico").text = str("info@coomeva.com").strip()
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}nombreGerente").text = re.sub(r"\(\s*[Ee]?\s*\)", "", str(row.get('NOMBRE DEL GERENTE', ''))).strip()
                                    		
        # Crear el árbol XML
        def sanitize_element_debug(element, log=[]):
            if element.text is not None and not isinstance(element.text, str):
                log.append(f"[Texto no válido] Elemento: <{element.tag}> - Valor: {element.text} - Tipo: {type(element.text)}")
                element.text = str(element.text)
            for key, value in element.attrib.items():
                if not isinstance(value, str):
                    log.append(f"[Atributo no válido] Elemento: <{element.tag}> - Atributo: {key} - Valor: {value} - Tipo: {type(value)}")
                    element.attrib[key] = str(value)
            for child in element:
                sanitize_element_debug(child, log)
            return log
        
        
        log = sanitize_element_debug(sucursales)
        if log:
            print("🔍 Valores corregidos en el XML")
            for entry in log:
                print(entry)
        else:
            print("✅ Todos los valores del XML ya eran válidos.")


        tree = ET.ElementTree(sucursales)
        ET.indent(tree, space="  ", level=0)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
            tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
            st.success("✅ XML generado exitosamente.")
            with open(tmp.name, "rb") as f:
                st.download_button("📥 Descargar XML", f, file_name="oficinas.xml", mime="application/xml")
