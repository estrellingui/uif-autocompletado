/**
 * POST /api/process
 * Receives: FormData { excel: File, pdfs: File[], apiKey: string, model: string }
 * Returns:  JSON { success, filename, excelBase64, rows, report, warnings }
 */

import { NextRequest, NextResponse } from "next/server";
import { extractDataWithAI } from "../../../../lib/ai-mapper";
import {
  readExcelHeaders,
  buildAndWriteExcel,
} from "../../../../lib/excel-handler";

export const runtime = "nodejs";
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  const warnings: string[] = [];

  try {
    // ── Parse multipart form data ─────────────────────────────────────────
    let formData: FormData;
    try {
      formData = await req.formData();
    } catch {
      return NextResponse.json(
        { success: false, error: "No se pudo leer el formulario enviado." },
        { status: 400 }
      );
    }

    const excelFile = formData.get("excel") as File | null;
    const pdfFiles = formData.getAll("pdfs") as File[];
    const apiKey =
      (formData.get("apiKey") as string | null) ||
      process.env.ANTHROPIC_API_KEY ||
      "";
    const model = (formData.get("model") as string | null) || "claude-sonnet-4-5";

    // ── Validation ────────────────────────────────────────────────────────
    if (!excelFile) {
      return NextResponse.json(
        { success: false, error: "No se recibió el archivo Excel." },
        { status: 400 }
      );
    }
    if (!pdfFiles.length) {
      return NextResponse.json(
        { success: false, error: "No se recibió ningún PDF." },
        { status: 400 }
      );
    }
    if (!apiKey) {
      return NextResponse.json(
        { success: false, error: "API Key de Claude no configurada." },
        { status: 400 }
      );
    }

    // ── Read Excel template ────────────────────────────────────────────────
    const excelBuffer = Buffer.from(await excelFile.arrayBuffer());
    const headers = await readExcelHeaders(excelBuffer);

    // ── Process each PDF ──────────────────────────────────────────────────
    let combinedText = "";
    const filenames: string[] = [];

    for (const pdfFile of pdfFiles) {
      const pdfBuffer = Buffer.from(await pdfFile.arrayBuffer());
      const extracted = await extractTextFromPdf(pdfBuffer, pdfFile.name);

      filenames.push(pdfFile.name);
      warnings.push(...extracted.warnings);

      if (extracted.text.trim().length > 50) {
        combinedText +=
          `\n\n--- DOCUMENTO: ${pdfFile.name} ---\n` + extracted.text;
      }
    }

    if (!combinedText.trim()) {
      return NextResponse.json(
        {
          success: false,
          error:
            "No se pudo extraer texto de ningún PDF. " +
            "Asegurate de subir PDFs con texto digital (no escaneados).",
          warnings,
        },
        { status: 422 }
      );
    }

    // ── AI extraction ─────────────────────────────────────────────────────
    let aiResult;
    try {
      aiResult = await extractDataWithAI(combinedText, apiKey, model);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return NextResponse.json(
        {
          success: false,
          error: `Error al consultar la IA: ${msg}`,
          warnings,
        },
        { status: 502 }
      );
    }

    // ── Build Excel with extracted data ───────────────────────────────────
    const { buffer, rows, report } = await buildAndWriteExcel(
      excelBuffer,
      aiResult,
      headers
    );

    // ── Response ──────────────────────────────────────────────────────────
    const timestamp = new Date()
      .toISOString()
      .slice(0, 19)
      .replace(/[:T]/g, "-");
    const outputFilename = `UIF_completado_${timestamp}.xlsx`;

    return NextResponse.json({
      success: true,
      filename: outputFilename,
      excelBase64: buffer.toString("base64"),
      rows,
      report,
      confianza: aiResult.confianza,
      partes: aiResult.partes,
      warnings,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("Unexpected error in /api/process:", err);
    return NextResponse.json(
      {
        success: false,
        error: `Error inesperado: ${msg}`,
        warnings,
      },
      { status: 500 }
    );
  }
}
