/**
 * POST /api/process
 * Receives JSON: { pdfText, excelBase64, apiKey, model }
 * PDF text is extracted client-side (avoids Vercel 4.5MB body limit).
 */

import { NextRequest, NextResponse } from "next/server";
import { extractDataWithAI } from "../../../../lib/ai-mapper";
import { readExcelHeaders, buildAndWriteExcel } from "../../../../lib/excel-handler";

export const runtime = "nodejs";
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { pdfText, excelBase64, apiKey: bodyApiKey, model } = body;

    const apiKey = bodyApiKey || process.env.ANTHROPIC_API_KEY || "";

    // Validations
    if (!excelBase64) {
      return NextResponse.json(
        { success: false, error: "No se recibió el archivo Excel." },
        { status: 400 }
      );
    }
    if (!pdfText || String(pdfText).trim().length < 50) {
      return NextResponse.json(
        { success: false, error: "No se pudo extraer texto del PDF. Asegurate de que sea un PDF digital (no escaneado)." },
        { status: 400 }
      );
    }
    if (!apiKey) {
      return NextResponse.json(
        { success: false, error: "API Key de Claude no configurada. Ingresala en el campo de Configuración." },
        { status: 400 }
      );
    }

    // Read Excel headers
    const excelBuffer = Buffer.from(excelBase64, "base64");
    const headers = await readExcelHeaders(excelBuffer);

    // AI extraction
    const aiResult = await extractDataWithAI(
      String(pdfText),
      apiKey,
      model || "claude-sonnet-4-5"
    );

    // Build output Excel
    const { buffer, rows, report } = await buildAndWriteExcel(
      excelBuffer,
      aiResult,
      headers
    );

    const timestamp = new Date()
      .toISOString()
      .slice(0, 19)
      .replace(/[:T]/g, "-");

    return NextResponse.json({
      success: true,
      filename: `UIF_completado_${timestamp}.xlsx`,
      excelBase64: buffer.toString("base64"),
      rows,
      report,
      confianza: aiResult.confianza,
      partes: aiResult.partes,
      warnings: [],
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("Error in /api/process:", err);
    return NextResponse.json(
      { success: false, error: `Error inesperado: ${msg}` },
      { status: 500 }
    );
  }
}
