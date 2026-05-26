"use client";

import { useState, useRef, useCallback } from "react";

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────
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

function confidenceBadge(nivel: string) {
  if (nivel === "ALTO")
    return (
      <span className="badge-green">Confianza ALTA</span>
    );
  if (nivel === "MEDIO")
    return (
      <span className="badge-yellow">Confianza MEDIA — revisá los campos amarillos</span>
    );
  return (
    <span className="badge-red">Confianza BAJA — revisá el Excel cuidadosamente</span>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Component: File drop zone
// ─────────────────────────────────────────────────────────────────────────────
function DropZone({
  label,
  accept,
  multiple,
  files,
  onChange,
}: {
  label: string;
  accept: string;
  multiple?: boolean;
  files: File[];
  onChange: (files: File[]) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const dropped = Array.from(e.dataTransfer.files);
      onChange(multiple ? dropped : [dropped[0]]);
    },
    [multiple, onChange]
  );

  return (
    <div
      className={`upload-zone ${dragging ? "active" : ""}`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        multiple={multiple}
        className="hidden"
        onChange={(e) => {
          const selected = Array.from(e.target.files || []);
          onChange(multiple ? selected : [selected[0]]);
        }}
      />
      {files.length === 0 ? (
        <>
          <div className="text-4xl mb-3">📂</div>
          <p className="text-gray-600 font-medium">{label}</p>
          <p className="text-gray-400 text-sm mt-1">
            Hacé clic o arrastrá el archivo aquí
          </p>
        </>
      ) : (
        <>
          <div className="text-4xl mb-3">✅</div>
          {files.map((f, i) => (
            <p key={i} className="text-blue-700 font-medium text-sm">
              {f.name}{" "}
              <span className="text-gray-400">
                ({(f.size / 1024).toFixed(0)} KB)
              </span>
            </p>
          ))}
          <p className="text-gray-400 text-xs mt-2">
            Hacé clic para cambiar el archivo
          </p>
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Main Page
// ─────────────────────────────────────────────────────────────────────────────
export default function Home() {
  const [excelFiles, setExcelFiles] = useState<File[]>([]);
  const [pdfFiles, setPdfFiles] = useState<File[]>([]);
  const [apiKey, setApiKey] = useState("");
  const [showKey, setShowKey] = useState(false);
  const [model, setModel] = useState("claude-sonnet-4-5");
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState("");
  const [result, setResult] = useState<ApiResponse | null>(null);
  const [error, setError] = useState<string | null>(null);

  const canProcess =
    excelFiles.length > 0 && pdfFiles.length > 0 && apiKey.trim().length > 0;

  const handleProcess = async () => {
    if (!canProcess) return;
    setIsProcessing(true);
    setError(null);
    setResult(null);
    setStatusMsg("Subiendo archivos...");

    try {
      const formData = new FormData();
      formData.append("excel", excelFiles[0]);
      pdfFiles.forEach((f) => formData.append("pdfs", f));
      formData.append("apiKey", apiKey.trim());
      formData.append("model", model);

      setStatusMsg("Extrayendo texto del PDF...");
      const res = await fetch("/api/process", {
        method: "POST",
        body: formData,
      });

      setStatusMsg("Analizando con inteligencia artificial...");
      const data: ApiResponse = await res.json();

      if (!data.success) {
        setError(data.error || "Error desconocido.");
      } else {
        setResult(data);
      }
    } catch (e) {
      setError(
        "No se pudo conectar con el servidor. Verificá tu conexión a internet."
      );
    } finally {
      setIsProcessing(false);
      setStatusMsg("");
    }
  };

  return (
    <div className="min-h-screen">
      {/* ── Header ── */}
      <header className="bg-gradient-to-r from-blue-900 to-blue-700 text-white py-8 px-6 shadow-lg">
        <div className="max-w-4xl mx-auto">
          <h1 className="text-3xl font-bold tracking-tight">
            📋 UIF — Autocompletado de Planilla
          </h1>
          <p className="text-blue-200 mt-2 text-lg">
            Completá automáticamente la planilla UIF desde boletos de
            compraventa o escrituras en PDF
          </p>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 py-8 space-y-6">
        {/* ── API Key config ── */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
            <span>🔑</span> Configuración
          </h2>
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                API Key de Claude (Anthropic)
              </label>
              <div className="flex gap-2">
                <input
                  type={showKey ? "text" : "password"}
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="sk-ant-..."
                  className="flex-1 border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <button
                  type="button"
                  onClick={() => setShowKey(!showKey)}
                  className="px-3 py-2 text-gray-500 border border-gray-300 rounded-lg hover:bg-gray-50 text-sm"
                >
                  {showKey ? "Ocultar" : "Mostrar"}
                </button>
              </div>
              <p className="text-xs text-gray-400 mt-1">
                Conseguila en{" "}
                <a
                  href="https://console.anthropic.com"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-blue-500 underline"
                >
                  console.anthropic.com
                </a>{" "}
                → API Keys → Create Key. No se guarda en ningún servidor.
              </p>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Modelo de IA
              </label>
              <select
                value={model}
                onChange={(e) => setModel(e.target.value)}
                className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <option value="claude-opus-4-5">
                  Opus 4.5 — Máxima precisión (recomendado para documentos complejos)
                </option>
                <option value="claude-sonnet-4-5">
                  Sonnet 4.5 — Equilibrio velocidad y precisión
                </option>
                <option value="claude-haiku-4-5">
                  Haiku 4.5 — Más rápido (documentos simples)
                </option>
              </select>
            </div>
          </div>
        </div>

        {/* ── Step 1: Excel ── */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-800 mb-1 flex items-center gap-2">
            <span className="bg-blue-600 text-white rounded-full w-7 h-7 flex items-center justify-center text-sm font-bold">1</span>
            Subir planilla UIF vacía (.xls o .xlsx)
          </h2>
          <p className="text-sm text-gray-500 mb-4">
            Es el Excel base de la UIF con las columnas ya definidas
          </p>
          <DropZone
            label="Seleccioná la planilla UIF"
            accept=".xls,.xlsx"
            files={excelFiles}
            onChange={setExcelFiles}
          />
        </div>

        {/* ── Step 2: PDFs ── */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-800 mb-1 flex items-center gap-2">
            <span className="bg-blue-600 text-white rounded-full w-7 h-7 flex items-center justify-center text-sm font-bold">2</span>
            Subir PDF(s) del documento legal
          </h2>
          <p className="text-sm text-gray-500 mb-4">
            Boleto de compraventa, escritura, cesión — podés subir varios a la vez
          </p>
          <DropZone
            label="Seleccioná uno o más PDF"
            accept=".pdf"
            multiple
            files={pdfFiles}
            onChange={setPdfFiles}
          />
        </div>

        {/* ── Step 3: Process ── */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
            <span className="bg-blue-600 text-white rounded-full w-7 h-7 flex items-center justify-center text-sm font-bold">3</span>
            Procesar
          </h2>

          {!apiKey.trim() && (
            <p className="text-sm text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3">
              ⚠️ Ingresá la API Key de Claude en el panel de arriba para continuar.
            </p>
          )}

          <button
            onClick={handleProcess}
            disabled={!canProcess || isProcessing}
            className="btn-primary w-full text-base py-4"
          >
            {isProcessing ? (
              <>
                <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24" fill="none">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
                </svg>
                {statusMsg || "Procesando..."}
              </>
            ) : (
              "🚀 Procesar documentos"
            )}
          </button>

          {isProcessing && (
            <p className="text-sm text-gray-500 text-center mt-3">
              Esto puede tardar hasta 60 segundos. No cerrés la página.
            </p>
          )}
        </div>

        {/* ── Error ── */}
        {error && (
          <div className="bg-red-50 border border-red-200 rounded-xl p-5">
            <h3 className="font-semibold text-red-700 mb-1">❌ Error</h3>
            <p className="text-red-600 text-sm">{error}</p>
          </div>
        )}

        {/* ── Results ── */}
        {result?.success && (
          <div className="space-y-4">
            {/* Summary */}
            <div className="card">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-lg font-semibold text-gray-800">
                  ✅ Procesamiento exitoso
                </h2>
                {result.confianza && confidenceBadge(result.confianza.nivel)}
              </div>

              <div className="grid grid-cols-3 gap-4 mb-4">
                <div className="bg-green-50 rounded-lg p-3 text-center">
                  <div className="text-2xl font-bold text-green-700">
                    {result.confianza?.campos_encontrados?.length ?? 0}
                  </div>
                  <div className="text-xs text-green-600 mt-1">Campos encontrados</div>
                </div>
                <div className="bg-red-50 rounded-lg p-3 text-center">
                  <div className="text-2xl font-bold text-red-700">
                    {result.confianza?.campos_no_encontrados?.length ?? 0}
                  </div>
                  <div className="text-xs text-red-600 mt-1">No encontrados 🔴</div>
                </div>
                <div className="bg-yellow-50 rounded-lg p-3 text-center">
                  <div className="text-2xl font-bold text-yellow-700">
                    {result.confianza?.campos_revisar?.length ?? 0}
                  </div>
                  <div className="text-xs text-yellow-600 mt-1">A revisar 🟡</div>
                </div>
              </div>

              {/* Parties */}
              {result.partes && result.partes.length > 0 && (
                <div>
                  <h3 className="text-sm font-semibold text-gray-700 mb-2">
                    Partes identificadas ({result.partes.length})
                  </h3>
                  <div className="space-y-2">
                    {result.partes.map((p, i) => {
                      const rol = String(p.Rol ?? "?");
                      const tipo = String(p.Tipo_de_Persona ?? "Humana");
                      const nombre =
                        tipo === "Humana"
                          ? `${p.Apellidos_PH ?? ""} ${p.Nombres_PH ?? ""}`.trim()
                          : String(p.Denominacion_PJ ?? "Sin denominación");
                      const doc =
                        tipo === "Humana"
                          ? `DNI ${p.Numero_Documento_PH ?? "—"}`
                          : `CUIT ${p.CUIT_CUIL_PJ ?? "—"}`;
                      return (
                        <div
                          key={i}
                          className="flex items-center gap-3 bg-gray-50 rounded-lg px-4 py-2 text-sm"
                        >
                          <span className="font-medium text-blue-700 w-24 shrink-0">
                            {rol}
                          </span>
                          <span className="font-semibold text-gray-800">{nombre}</span>
                          <span className="text-gray-400 text-xs">{doc}</span>
                          <span className="text-gray-400 text-xs">
                            {tipo === "Humana" ? "👤 Persona Humana" : "🏢 Persona Jurídica"}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            {/* Download */}
            <div className="card bg-blue-50 border-blue-200">
              <h3 className="font-semibold text-blue-800 mb-3">
                📥 Descargar Excel completado
              </h3>
              <p className="text-sm text-blue-600 mb-4">
                Las celdas en <strong>rojo</strong> son campos que no se
                encontraron en el documento. Las celdas en{" "}
                <strong>amarillo</strong> requieren verificación manual.
              </p>
              <button
                onClick={() =>
                  downloadBase64(result.excelBase64!, result.filename!)
                }
                className="btn-primary"
              >
                ⬇️ Descargar {result.filename}
              </button>
            </div>

            {/* Report */}
            <div className="card">
              <h3 className="font-semibold text-gray-800 mb-3">
                📋 Reporte detallado
              </h3>
              <pre className="bg-gray-50 border border-gray-200 rounded-lg p-4 text-xs text-gray-700 overflow-auto max-h-64 whitespace-pre-wrap">
                {result.report}
              </pre>
              <button
                onClick={() => {
                  const blob = new Blob([result.report!], { type: "text/plain" });
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement("a");
                  a.href = url;
                  a.download = `reporte_${result.filename?.replace(".xlsx", "")}.txt`;
                  a.click();
                }}
                className="mt-3 text-sm text-blue-600 underline hover:text-blue-800"
              >
                Descargar reporte como .txt
              </button>
            </div>

            {/* Warnings */}
            {result.warnings && result.warnings.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 rounded-xl p-4">
                <h3 className="font-semibold text-amber-700 mb-2">
                  ⚠️ Avisos
                </h3>
                <ul className="space-y-1">
                  {result.warnings.map((w, i) => (
                    <li key={i} className="text-sm text-amber-700">
                      • {w}
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        )}
      </main>

      {/* ── Footer ── */}
      <footer className="border-t border-gray-200 mt-12 py-6 px-4 text-center text-xs text-gray-400">
        <p>
          Los archivos se procesan en la nube (Vercel). Solo el texto del PDF se
          envía a la API de Claude (Anthropic) para su interpretación. No se
          almacena ningún archivo ni dato personal en los servidores.
        </p>
      </footer>
    </div>
  );
}
