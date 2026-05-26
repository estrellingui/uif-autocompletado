/**
 * Excel read/write module.
 * Uses SheetJS (xlsx) to read both .xls and .xlsx templates.
 * Uses ExcelJS to write the filled output as .xlsx with color coding.
 */

import ExcelJS from "exceljs";
import type { ExtractionData } from "./ai-mapper";

// ── Field mapping: AI JSON key → exact Excel column header ──────────────────
const FIELD_TO_COLUMN: Record<string, string> = {
  ID_Operacion: "ID_Operacion",
  Fecha_de_la_operacion: "Fecha_de_la_operación",
  Tipo_de_moneda_de_origen: "Tipo_de_moneda_de_origen",
  Tipo_de_moneda_extranjera: "Tipo_de_moneda_extranjera",
  Monto_total_operacion_en_Pesos: "Monto_total_de_la_operación_equivalente_en_Pesos",
  Monto_total_operacion_en_moneda_origen: "Monto_total_de_la_operación_en_moneda_de_origen",
  Nomenclatura_catastral_o_matricula: "Nomenclatura_catastral_o_matrícula_del_inmueble_transferido",
  Provincia_inmueble: "Provincia_del_inmueble",
  Localidad_inmueble: "Localidad_del_inmueble",
  Calle_inmueble: "Calle_del_inmueble",
  Numero_inmueble: "Número_del_inmueble",
  Piso_inmueble: "Piso_del_inmueble",
  Departamento_inmueble: "Departamento_del_inmueble",
  Codigo_postal_inmueble: "Código_postal_del_inmueble",
  Forma_de_pago: "Forma_de_pago",
  Tipo_de_activo_virtual: "Tipo_de_activo_virtual",
  Otra_forma_pago: "Otra",
  Tipo_moneda_origen_pago: "Tipo_de_moneda_de_origen_del_pago",
  Tipo_moneda_extranjera_pago: "Tipo_de_moneda_extranjera_de_origen_del_pago",
  Monto_pagado_en_Pesos: "Monto_Pagado_de_la_operación_equivalente_en_Pesos",
  Monto_pagado_en_moneda_origen: "Monto_Pagado_de_la_operación_en_moneda_de_origen",
  Cambio_del_dia: "Cambio del dia ",
  Rol: "Rol_en_la_Operación_CompradorVendedor",
  Tipo_de_Persona: "Tipo_de_Persona_CompradorVendedor",
  Denominacion_PJ: "Denominación_Persona_Jurídica",
  Denominacion_PJ_extranjera: "Denominación_Persona_Jurídica_extranjera",
  CUIT_CUIL_PH: "Número_de_CUIT_CUIL_Persona_Humana",
  CUIT_CUIL_PJ: "Número_de_CUIT_CUIL_Persona_Jurídica",
  CDI_PH: "Número_de_CDI_Persona_Humana",
  CDI_PJ: "Número_de_CDI_Persona_Jurídica",
  Tipo_ID_PJ_Extranjera: "Tipo_Identificador_Tributario_Persona_Jurídica_Extranjera",
  Nro_ID_PJ_Extranjera: "Nro_Identificador_Tributario_Persona_Jurídica_Extranjera",
  Apellidos_PH: "Apellidos_PersonaHumana",
  Apellidos_PH_Extranjera: "Apellidos_PersonaHumanaExtranjera",
  Nombres_PH: "Nombres_PersonaHumana",
  Nombres_PH_Extranjera: "Nombres_PersonaHumanaExtranjera",
  Tipo_Documento_PH: "Tipo_de_Documento_PersonaHumana",
  Tipo_Documento_PH_Extranjera: "Tipo_de_Documento_PersonaHumanaExtranjera",
  Numero_Documento_PH: "Número_de_Documento_PersonaHumana",
  Numero_Documento_PH_Extranjera: "Número_de_Documento_PersonaHumanaExtranjera",
  Nacionalidad_PH: "Nacionalidad_PersonaHumana",
  Nacionalidad_PH_Extranjera: "Nacionalidad_PersonaHumanaExtranjera",
  Fecha_nacimiento_PH: "Fecha_de_nacimiento_PersonaHumana",
  Fecha_nacimiento_PH_Extranjera: "Fecha_de_nacimiento_PersonaHumanaExtranjera",
  Es_PEP_PH: "Es_PEP_PersonaHumana",
  Es_PEP_PH_Extranjera: "Es_PEP_PersonaHumanaExtranjera",
  Porcentaje: "Porcentaje_CompradorVendedor",
  Pais: "País_CompradorVendedor",
  Provincia: "Provincia_CompradorVendedor",
  Otro_pais: "Otro_país_CompradorVendedor",
  Provincia_Estado_extranjero: "Provincia_Estado_CompradorVendedor",
  Localidad: "Localidad_CompradorVendedor",
  Localidad_ciudad_extranjera: "Localidad_Ciudad_CompradorVendedor",
  Calle: "Calle_CompradorVendedor",
  Numero: "Número_CompradorVendedor",
  Piso: "Piso_CompradorVendedor",
  Departamento: "Departamento_CompradorVendedor",
  Codigo_Postal: "Codigo Postal",
  Codigo_area_tel: "Código_de_área_telefónico_CompradorVendedor",
  Telefono: "Teléfono_CompradorVendedor",
  Email: "Dirección_de_correo_electrónico_CompradorVendedor",
};

const PARTY_FIELDS = new Set([
  "Rol","Tipo_de_Persona","Denominacion_PJ","Denominacion_PJ_extranjera",
  "CUIT_CUIL_PH","CUIT_CUIL_PJ","CDI_PH","CDI_PJ",
  "Tipo_ID_PJ_Extranjera","Nro_ID_PJ_Extranjera",
  "Apellidos_PH","Apellidos_PH_Extranjera","Nombres_PH","Nombres_PH_Extranjera",
  "Tipo_Documento_PH","Tipo_Documento_PH_Extranjera",
  "Numero_Documento_PH","Numero_Documento_PH_Extranjera",
  "Nacionalidad_PH","Nacionalidad_PH_Extranjera",
  "Fecha_nacimiento_PH","Fecha_nacimiento_PH_Extranjera",
  "Es_PEP_PH","Es_PEP_PH_Extranjera","Porcentaje",
  "Pais","Provincia","Otro_pais","Provincia_Estado_extranjero",
  "Localidad","Localidad_ciudad_extranjera",
  "Calle","Numero","Piso","Departamento","Codigo_Postal",
  "Codigo_area_tel","Telefono","Email",
]);

export interface RowData {
  [header: string]: string | number | null;
}

const columnToField = Object.fromEntries(
  Object.entries(FIELD_TO_COLUMN).map(([k, v]) => [v, k])
);

/**
 * Read column headers from .xls or .xlsx using SheetJS.
 * SheetJS handles both formats correctly.
 */
export async function readExcelHeaders(buffer: Buffer): Promise<string[]> {
  // Dynamic import — SheetJS is a CommonJS module
  const XLSX = await import("xlsx");
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("No se encontró ninguna hoja en el archivo Excel.");

  const rows = XLSX.utils.sheet_to_json<(string | undefined)[]>(ws, {
    header: 1,
    defval: undefined,
  });

  if (!rows[0] || rows[0].length === 0) {
    throw new Error("La primera fila del Excel está vacía. Verificá que sea la planilla UIF correcta.");
  }

  return (rows[0] as (string | undefined)[]).map((h, i) =>
    h ? String(h) : `__col_${i + 1}__`
  );
}

/**
 * Build data rows from AI extraction and write a new .xlsx workbook.
 * Always outputs .xlsx regardless of input format.
 */
export async function buildAndWriteExcel(
  _templateBuffer: Buffer,
  extracted: ExtractionData,
  headers: string[]
): Promise<{ buffer: Buffer; rows: RowData[]; report: string }> {
  const operacion = extracted.operacion || {};
  const partes = extracted.partes?.length ? extracted.partes : [{}];

  // Build one row per party
  const rows: RowData[] = partes.map((parte) => {
    const row: RowData = {};
    headers.forEach((header) => {
      if (header.startsWith("__col_")) { row[header] = null; return; }
      const aiField = columnToField[header];
      if (!aiField) { row[header] = null; return; }
      const source = PARTY_FIELDS.has(aiField) ? parte : operacion;
      const value = (source as Record<string, unknown>)[aiField];
      if (value === null || value === undefined) {
        row[header] = "No encontrado";
      } else if (typeof value === "string" && value.toUpperCase().startsWith("REVISAR")) {
        row[header] = value;
      } else {
        row[header] = value as string | number;
      }
    });
    return row;
  });

  // Create a fresh .xlsx workbook with headers + data
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Hoja1");

  // Header row
  ws.addRow(headers);
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3E4F7" } };
  headerRow.commit();

  // Data rows with color coding
  rows.forEach((rowData) => {
    const values = headers.map((h) => rowData[h] ?? "");
    ws.addRow(values);
    const dataRow = ws.lastRow!;
    headers.forEach((header, colIdx) => {
      const value = rowData[header];
      const cell = dataRow.getCell(colIdx + 1);
      if (value === "No encontrado") {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
        cell.font = { color: { argb: "FF9C0006" } };
      } else if (typeof value === "string" && value.startsWith("REVISAR")) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFEB9C" } };
        cell.font = { color: { argb: "FF9C5700" } };
      }
    });
    dataRow.commit();
  });

  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer, rows, report: generateReport(extracted, rows) };
}

function generateReport(extracted: ExtractionData, rows: RowData[]): string {
  const { confianza, partes } = extracted;
  const now = new Date().toLocaleString("es-AR");
  const lines = [
    "=".repeat(60),
    "REPORTE DE EXTRACCIÓN UIF",
    `Fecha: ${now}`,
    "=".repeat(60),
    "",
    `PARTES IDENTIFICADAS: ${partes?.length ?? 0}`,
  ];
  partes?.forEach((p, i) => {
    const rol = String(p.Rol ?? "Desconocido");
    const tipo = String(p.Tipo_de_Persona ?? "?");
    const nombre =
      tipo === "Humana"
        ? `${p.Apellidos_PH ?? ""} ${p.Nombres_PH ?? ""}`.trim()
        : String(p.Denominacion_PJ ?? "sin nombre");
    const doc =
      tipo === "Humana"
        ? String(p.Numero_Documento_PH ?? "sin doc")
        : String(p.CUIT_CUIL_PJ ?? "sin CUIT");
    lines.push(`  ${i + 1}. [${rol}] ${nombre} — ${doc}`);
  });
  lines.push("", `NIVEL DE CONFIANZA: ${confianza?.nivel ?? "N/A"}`, "");
  lines.push("CAMPOS ENCONTRADOS:");
  confianza?.campos_encontrados?.forEach((f) => lines.push(`  ✓ ${f}`));
  lines.push("", "CAMPOS NO ENCONTRADOS (celdas rojas en el Excel):");
  confianza?.campos_no_encontrados?.forEach((f) => lines.push(`  ✗ ${f}`));
  if (confianza?.campos_revisar?.length) {
    lines.push("", "CAMPOS A REVISAR (celdas amarillas en el Excel):");
    confianza.campos_revisar.forEach((f) => lines.push(`  ⚠ ${f}`));
  }
  lines.push("", "=".repeat(60));
  return lines.join("\n");
}
