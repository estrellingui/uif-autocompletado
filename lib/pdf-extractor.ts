/**
 * PDF text extraction for Node.js / Next.js serverless.
 * Uses pdf-parse for digital PDFs.
 * OCR not available in serverless — returns a warning for scanned PDFs.
 */

export interface ExtractionResult {
  text: string;
  pageCount: number;
  method: string;
  warnings: string[];
}

export async function extractTextFromPdf(
  buffer: Buffer,
  filename: string
): Promise<ExtractionResult> {
  const result: ExtractionResult = {
    text: "",
    pageCount: 0,
    method: "digital",
    warnings: [],
  };

  try {
    // Dynamic import to avoid edge-runtime issues
    const pdfParse = (await import("pdf-parse")).default;
    const data = await pdfParse(buffer);

    result.text = data.text || "";
    result.pageCount = data.numpages || 0;
    result.method = "digital (pdf-parse)";

    if (result.text.trim().length < 100) {
      result.warnings.push(
        `El PDF "${filename}" parece estar escaneado o tener muy poco texto. ` +
          "Los documentos escaneados no se pueden procesar automáticamente en esta versión. " +
          "Usá un PDF con texto digital (generado por computadora)."
      );
    }

    return result;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    result.warnings.push(
      `No se pudo extraer texto de "${filename}": ${msg}`
    );
    return result;
  }
}
