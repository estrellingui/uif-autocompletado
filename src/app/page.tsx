"use client";

import { useState, useRef, useCallback } from "react";

// ─── PDF text extraction via CDN — no npm package needed ─────────────────────
async function extractPdfText(
  file: File
): Promise<{ text: string; warning?: string }> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if (typeof window !== "undefined" && !(window as any).pdfjsLib) {
      await new Promise<void>((resolve, reject) => {
        const script = document.createElement("script");
        script.src =
          "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js";
        script.onload = () => resolve();
        script.onerror = () =>
          reject(new Error("No se pudo cargar el lector de PDF"));
        document.head.appendChild(script);
      });
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const pdfjsLib = (window as any).pdfjsLib;
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js";

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib
      .getDocument({ data: new Uint8Array(arrayBuffer) })
      .promise;

    const pages: string[] = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pageText = content.items
        .filter((item: any) => "str" in item)
        .map((item: any) => item.str)
        .join(" ");
      pages.push(pageText);
    }
    const text = pages.join("\n\n");
    if (text.trim().length < 100) {
      return {
        text,
        warning: `El PDF "${file.name}" tiene muy poco texto. Si está escaneado no se puede procesar automáticamente.`,
      };
    }
    return { text };
  } catch (err) {
    return {
      text: "",
      warning: `No se pudo leer "${file.name}": ${
        err instanceof Error ? err.message : String(err)
      }`,
    };
  }
}

// ─── File to base64 ───────────────────────────────────────────────────────────
function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.split(",")[1]);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ─── Types ────────────────────────────────────────────────────────────────────
interface ApiResponse {
  success: boolean;
  error?: string;
  filename?: string;
  excelBase64?: string;
  rows?: Record<string, string | number | null>[];
  report?: string;
  confianza?: {
    nivel: string;
    campos_encontrados: string[];
    campos_no_encontrados: string[];
    campos_revisar: string[];
  };
  partes?: Record<string, unknown>[];
  warnings?: string[];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function downloadBase64(base64: string, filename: string) {
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function ConfidenceBadge({ nivel }: { nivel: string }) {
  if (nivel === "ALTO")
    return <span className="badge-green">Confianza ALTA</span>;
  if (nivel === "MEDIO")
    return (
      <span className="badge-yellow">
