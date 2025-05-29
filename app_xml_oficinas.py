import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import tempfile
import locale

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

st.title("Generador de XML de Oficinas")

# Subida de archivos
oficinas_file = st.file_uploader("Sube el archivo de oficinas (Excel)", type=["xlsx"])
divipola_file = st.file_uploader("Sube el archivo DIVIPOLA (Excel)", type=["xlsx"])

if oficinas_file and divipola_file:
    if st.button("Generar XML"):
        oficinas = pd.read_excel(oficinas_file, sheet_name='Hoja1', engine="openpyxl", dtype=str)
        divipola = pd.read_excel(divipola_file, sheet_name='Sheet1', engine="openpyxl", dtype=str)

        oficinas['CODIGO_DEPARTAMENTO'] = oficinas['CODIGO DEL DEPARTAMENTO '].astype('int')
        oficinas['CODIGO_MUNICIPIO'] = oficinas['CODIGO DEL MUNICIPIO '].astype('int')
        divipola['CODIGO_DEPARTAMENTO_ORIGINAL'] = divipola['CODIGO_DEPARTAMENTO']
        divipola['CODIGO_MUNICIPIO_ORIGINAL'] = divipola['CODIGO_MUNICIPIO']
        divipola['CODIGO_DEPARTAMENTO'] = divipola['CODIGO_DEPARTAMENTO'].astype('int')
        divipola['CODIGO_MUNICIPIO'] = divipola['CODIGO_MUNICIPIO'].astype('int')

        oficinas = pd.merge(oficinas, divipola, how='left', on=['CODIGO_DEPARTAMENTO', 'CODIGO_MUNICIPIO'])
        cantidad_oficinas = str(len(oficinas))

        ET.register_namespace('', "http://www.finagro.com.co/sit")
        sucursales = ET.Element("{http://www.finagro.com.co/sit}sit:sucursales", cifraDeControl=cantidad_oficinas)

        for _, row in oficinas.iterrows():
            sucursal = ET.SubElement(sucursales, "{http://www.finagro.com.co/sit}sit:sucursal")
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:codigoIntermediario").text = "101054"
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:codigoIdentificacionSucursal").text = row['CODIGO DE LA OFICINA']
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:nombre").text = row['NOMBRE DE LA OFICINA']
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:codigoDepartamento").text = row['CODIGO_DEPARTAMENTO_ORIGINAL']
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:codigoMunicipio").text = row['CODIGO_DPTO_MPIO']
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:direccion").text = "R\n" + row['DIRECCION DE LA OFICINA ']
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:numeroTelefonoFijo",
                          extension='', prefijoCiudad=row['PREFIJO TELEFONICO DEL MUNICIPIO ']).text = row['NUMERO TELEFONICO DE LA OFICINA 1 ']
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:numeroTelefonoFax", prefijoCiudad="").text = " "
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:correoElectronico").text = "info@coomeva.com"
            ET.SubElement(sucursal, "{http://www.finagro.com.co/sit}sit:nombreGerente").text = row['NOMBRE DEL GERENTE']

        tree = ET.ElementTree(sucursales)
        ET.indent(tree, space=" ", level=0)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
            tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
            st.success("✅ XML generado exitosamente.")
            with open(tmp.name, "rb") as f:
                st.download_button("📥 Descargar XML", f, file_name="oficinas.xml", mime="application/xml")
