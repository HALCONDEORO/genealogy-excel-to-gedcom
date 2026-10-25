"""Genera fixtures/fixture_sintetico.xlsx con datos ficticios."""
import os, openpyxl
os.makedirs('fixtures', exist_ok=True)
wb = openpyxl.Workbook(); wb.remove(wb.active)
ws = wb.create_sheet('30_Arbol_Confirmado')
ws.append(['id_persona','nombre','apellido1','apellido2','sexo','fecha_nacimiento','lugar_nacimiento','fecha_defuncion','lugar_defuncion','id_padre','id_madre','id_conyuge','fecha_matrimonio','lugar_matrimonio','fuente','url_fuente','notas'])
ws.append(['P000001','Ana','Ficticia','Ejemplo','F','1980-05-02','Villa Falsa','','','P000002','P000003','','','','S000001 Registro ficticio','',''])
ws.append(['P000002','Luis','Ejemplo','Modelo','M','c. 1950-1951','Villa Falsa','','','','','P000003','1975-01-01','Villa Falsa','S000001 Registro ficticio','',''])
ws.append(['P000003','Eva','Modelo','Prueba','F','1952','Villa Falsa','','','','','P000002','1975-01-01','Villa Falsa','S000001 Registro ficticio','',''])
def sheet(n, hdr, rows):
    w = wb.create_sheet(n)
    w.append(['titulo ficticio']); w.append([]); w.append([]); w.append(hdr)
    for r in rows: w.append(r)
sheet('02_Relaciones', ['ID_Relacion','ID_Persona_1','Tipo_relacion','ID_Persona_2','Fecha_desde','Fecha_hasta','ID_Lugar','Familia_Gramps_ID','Estado_hechos','Confianza','ID_Cita','Notas'],
      [['R000001','P000001','Padre biológico','P000002'], ['R000002','P000001','Madre biológica','P000003'], ['R000003','P000002','Cónyuge','P000003']])
sheet('04_Fuentes', ['ID_Fuente','Tipo_fuente','Titulo','Autor_o_entidad','Repositorio','Coleccion_o_fondo','Signatura','Fecha_fuente','Original_o_derivada','Calidad_fuente','URL','Notas'],
      [['S000001','Acta','Registro ficticio','Autor ficticio','Archivo ficticio','','','','Original','Media','','']])
sheet('05_Citas', ['ID_Cita','ID_Fuente','ID_Persona','Tipo_hecho','Valor_afirmado','Localizador_pagina_folio','Extracto_o_resumen','Evaluacion_evidencia','Corroborada','Confianza','ID_Documento','URL','Notas'],
      [['C000001','S000001','P000001','Nacimiento','1980-05-02','Folio 1','extracto','ok','Sí','Alta','','',''], ['C000002','S000001','P000002','Nacimiento','c.1950','Folio 2','extracto','ok','No','D000009','','','']])
sheet('06_Documentos', ['ID_Documento','Nombre_archivo','Tipo_documento','ID_Persona_principal','Personas_relacionadas','Fecha_documento','ID_Lugar','ID_Fuente','ID_Cita','Paginas','Transcripcion','OCR','Derechos_privacidad','SHA256','Ruta_local','Backup_verificado','URL_origen','Notas'],
      [['D000001','doc_ficticio.pdf','Acta','P000001','','','','S000001','C000001','1','No','No','Privado','','','','','']])
wb.save('fixtures/fixture_sintetico.xlsx')
