/**
 * Claude AI data extraction from legal documents.
 * Returns structured JSON matching the UIF Excel columns.
 */

import Anthropic from "@anthropic-ai/sdk";

const SYSTEM_PROMPT = `Eres un experto en cumplimiento normativo UIF (Unidad de Información Financiera) de Argentina.
Tu tarea es extraer información de documentos legales de compraventa de inmuebles (boletos, escrituras, cesiones)
y mapearla a los campos exactos de la planilla UIF.

REGLAS CRÍTICAS:
1. NUNCA inventar datos. Si un dato no aparece explícitamente, usar null.
2. Si el dato es ambiguo, marcarlo como "REVISAR: [razón breve]".
3. Retornar EXACTAMENTE el JSON solicitado, sin texto ni explicaciones extra.
4. Una operación tiene múltiples partes (comprador, vendedor, etc.). Crear un objeto por cada parte.
5. Para personas jurídicas argentinas: el CUIT va en CUIT_CUIL_PJ.
6. Para personas humanas argentinas: el CUIT/CUIL va en CUIT_CUIL_PH.
7. El DNI va en Numero_Documento_PH SIN puntos ni espacios (solo dígitos).
8. Las fechas deben estar en formato DD/MM/AAAA.
9. Los montos deben ser números (sin símbolos, sin puntos de miles, sin comas de miles).
10. PEP = Persona Expuesta Políticamente. Si no se menciona explícitamente, dejar null.
11. Si el rol no está claro, usar "REVISAR: rol no claro".`;

const EXTRACTION_TEMPLATE = `Analizá este documento legal y extraé todos los datos disponibles.

DOCUMENTO:
{TEXT}

Retorná ÚNICAMENTE este JSON (sin texto antes ni después):
{
  "operacion": {
    "Fecha_de_la_operacion": "DD/MM/AAAA o null",
    "Tipo_de_moneda_de_origen": "Pesos o Dolares Estadounidenses u otra moneda o null",
    "Tipo_de_moneda_extranjera": "nombre completo de la moneda extranjera si aplica, o null",
    "Monto_total_operacion_en_moneda_origen": <número o null>,
    "Monto_total_operacion_en_Pesos": <número o null>,
    "Nomenclatura_catastral_o_matricula": "descripción del lote/unidad o null",
    "Provincia_inmueble": "nombre de provincia o null",
    "Localidad_inmueble": "localidad o null",
    "Calle_inmueble": "calle o null",
    "Numero_inmueble": "número o null",
    "Piso_inmueble": null,
    "Departamento_inmueble": null,
    "Codigo_postal_inmueble": "CP o null",
    "Forma_de_pago": "Transferencia bancaria / Efectivo / Cheque / Depósito bancario / otro o null",
    "Tipo_de_activo_virtual": null,
    "Otra_forma_pago": null,
    "Tipo_moneda_origen_pago": "Pesos o Dolares Estadounidenses o null",
    "Tipo_moneda_extranjera_pago": null,
    "Monto_pagado_en_moneda_origen": <número o null>,
    "Monto_pagado_en_Pesos": null,
    "Cambio_del_dia": null
  },
  "partes": [
    {
      "Rol": "Comprador / Vendedor / Cedente / Cesionario / Apoderado",
      "Tipo_de_Persona": "Humana o Juridica",
      "Denominacion_PJ": "razón social si es jurídica, null si humana",
      "Denominacion_PJ_extranjera": null,
      "CUIT_CUIL_PH": "XX-XXXXXXXX-X o null",
      "CUIT_CUIL_PJ": "XX-XXXXXXXX-X o null",
      "CDI_PH": null,
      "CDI_PJ": null,
      "Tipo_ID_PJ_Extranjera": null,
      "Nro_ID_PJ_Extranjera": null,
      "Apellidos_PH": "apellido/s en mayúsculas o null",
      "Apellidos_PH_Extranjera": null,
      "Nombres_PH": "nombre/s o null",
      "Nombres_PH_Extranjera": null,
      "Tipo_Documento_PH": "DNI o Pasaporte o null",
      "Tipo_Documento_PH_Extranjera": null,
      "Numero_Documento_PH": "solo dígitos sin puntos o null",
      "Numero_Documento_PH_Extranjera": null,
      "Nacionalidad_PH": "gentilicio (ej: Argentina, Uruguaya) o null",
      "Nacionalidad_PH_Extranjera": null,
      "Fecha_nacimiento_PH": "DD/MM/AAAA o null",
      "Fecha_nacimiento_PH_Extranjera": null,
      "Es_PEP_PH": null,
      "Es_PEP_PH_Extranjera": null,
      "Porcentaje": null,
      "Pais": "Argentina u otro o null",
      "Provincia": "nombre de provincia argentina o null",
      "Otro_pais": null,
      "Provincia_Estado_extranjero": null,
      "Localidad": "localidad o null",
      "Localidad_ciudad_extranjera": null,
      "Calle": "nombre de la calle o null",
      "Numero": "número de la puerta o null",
      "Piso": null,
      "Departamento": null,
      "Codigo_Postal": "CP o null",
      "Codigo_area_tel": null,
      "Telefono": null,
      "Email": "correo electrónico o null"
    }
  ],
  "confianza": {
    "nivel": "ALTO / MEDIO / BAJO",
    "campos_encontrados": ["lista de los campos que tienen datos reales"],
    "campos_no_encontrados": ["lista de los campos que quedaron en null"],
    "campos_revisar": ["lista de campos con motivo de revisión"]
  }
}`;

export interface ExtractionData {
  operacion: Record<string, unknown>;
  partes: Record<string, unknown>[];
  confianza: {
    nivel: string;
    campos_encontrados: string[];
    campos_no_encontrados: string[];
    campos_revisar: string[];
  };
}

export async function extractDataWithAI(
  documentText: string,
  apiKey: string,
  model: string = "claude-sonnet-4-5"
): Promise<ExtractionData> {
  if (!apiKey?.trim()) {
    throw new Error("API Key de Claude no configurada.");
  }
  if (!documentText || documentText.trim().length < 50) {
    throw new Error("El texto del documento está vacío o es demasiado corto.");
  }

  const client = new Anthropic({ apiKey: apiKey.trim() });

  const prompt = EXTRACTION_TEMPLATE.replace(
    "{TEXT}",
    documentText.slice(0, 15000)
  );

  const message = await client.messages.create({
    model,
    max_tokens: 4096,
    system: SYSTEM_PROMPT,
    messages: [{ role: "user", content: prompt }],
  });

  const raw = (message.content[0] as { text: string }).text;
  return parseJsonResponse(raw);
}

function parseJsonResponse(raw: string): ExtractionData {
  let text = raw.trim();

  // Strip markdown code fences
  if (text.includes("```json")) {
    text = text.split("```json")[1].split("```")[0].trim();
  } else if (text.includes("```")) {
    text = text.split("```")[1].split("```")[0].trim();
  }

  try {
    return JSON.parse(text) as ExtractionData;
  } catch (e) {
    throw new Error(
      `La IA devolvió una respuesta que no se pudo interpretar. Intentá de nuevo.`
    );
  }
}
