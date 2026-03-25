"use client";

import { useCallback, useState } from "react";
import * as XLSX from "xlsx";
import { generateTxt, type GaliciaResult } from "./lib/galicia";

type State =
  | { status: "idle" }
  | { status: "processing" }
  | { status: "done"; result: GaliciaResult }
  | { status: "error"; message: string };

export default function Home() {
  const [state, setState] = useState<State>({ status: "idle" });

  const processFile = useCallback((file: File) => {
    setState({ status: "processing" });

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("No se pudo leer el archivo");

        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
          bookVBA: true,
        });

        const result = generateTxt(workbook);
        setState({ status: "done", result });
      } catch (err) {
        setState({
          status: "error",
          message: err instanceof Error ? err.message : String(err),
        });
      }
    };
    reader.onerror = () => {
      setState({ status: "error", message: "Error al leer el archivo" });
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      const file = e.dataTransfer.files?.[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const handleDownload = useCallback(() => {
    if (state.status !== "done") return;
    const { filename, content } = state.result;
    const blob = new Blob([content.buffer as ArrayBuffer], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }, [state]);

  const handleReset = () => setState({ status: "idle" });

  return (
    <main className="min-h-screen bg-gray-50 flex items-start justify-center pt-16 px-4">
      <div className="w-full max-w-lg">
        {/* Header */}
        <div className="mb-8 text-center">
          <div className="inline-flex items-center justify-center w-12 h-12 bg-red-600 rounded-xl mb-4">
            <svg
              className="w-6 h-6 text-white"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
          </div>
          <h1 className="text-2xl font-semibold text-gray-900">
            Generador TXT
          </h1>
          <p className="text-gray-500 mt-1 text-sm">
            Pago de Haberes — Banco Galicia
          </p>
        </div>

        {/* Card */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-200 p-8">
          {state.status === "idle" || state.status === "processing" ? (
            <div
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              className="border-2 border-dashed border-gray-200 rounded-xl p-10 text-center hover:border-red-400 hover:bg-red-50 transition-colors cursor-pointer"
              onClick={() => document.getElementById("file-input")?.click()}
            >
              {state.status === "processing" ? (
                <div className="flex flex-col items-center gap-3">
                  <div className="w-8 h-8 border-2 border-red-600 border-t-transparent rounded-full animate-spin" />
                  <p className="text-gray-500 text-sm">Procesando...</p>
                </div>
              ) : (
                <>
                  <svg
                    className="w-10 h-10 text-gray-300 mx-auto mb-3"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={1.5}
                      d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                    />
                  </svg>
                  <p className="text-gray-600 font-medium">
                    Arrastrá el archivo acá
                  </p>
                  <p className="text-gray-400 text-sm mt-1">
                    o hacé click para seleccionar
                  </p>
                  <p className="text-gray-300 text-xs mt-3">
                    Archivo .xlsm de Pago de Haberes Galicia
                  </p>
                </>
              )}
              <input
                id="file-input"
                type="file"
                accept=".xlsm,.xlsx"
                className="hidden"
                onChange={handleFileInput}
              />
            </div>
          ) : state.status === "error" ? (
            <div className="text-center">
              <div className="inline-flex items-center justify-center w-12 h-12 bg-red-100 rounded-full mb-4">
                <svg
                  className="w-6 h-6 text-red-600"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M6 18L18 6M6 6l12 12"
                  />
                </svg>
              </div>
              <p className="text-gray-800 font-medium mb-2">
                Error al procesar
              </p>
              <p className="text-sm text-red-600 bg-red-50 rounded-lg px-4 py-3 text-left font-mono">
                {state.message}
              </p>
              <button
                onClick={handleReset}
                className="mt-6 text-sm text-gray-500 hover:text-gray-700 underline"
              >
                Intentar de nuevo
              </button>
            </div>
          ) : (
            // done
            <div className="text-center">
              <div className="inline-flex items-center justify-center w-12 h-12 bg-green-100 rounded-full mb-4">
                <svg
                  className="w-6 h-6 text-green-600"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M5 13l4 4L19 7"
                  />
                </svg>
              </div>
              <p className="text-gray-800 font-medium mb-1">Archivo listo</p>
              <p className="text-xs text-gray-400 font-mono mb-6 break-all">
                {state.result.filename}
              </p>

              {/* Stats */}
              <div className="grid grid-cols-2 gap-3 mb-6">
                <div className="bg-gray-50 rounded-xl p-4">
                  <p className="text-2xl font-semibold text-gray-900">
                    {state.result.employeeCount}
                  </p>
                  <p className="text-xs text-gray-400 mt-0.5">empleados</p>
                </div>
                <div className="bg-gray-50 rounded-xl p-4">
                  <p className="text-2xl font-semibold text-gray-900 text-sm">
                    {new Intl.NumberFormat("es-AR", {
                      style: "currency",
                      currency: "ARS",
                      maximumFractionDigits: 0,
                    }).format(state.result.totalAmount)}
                  </p>
                  <p className="text-xs text-gray-400 mt-0.5">importe total</p>
                </div>
              </div>

              <button
                onClick={handleDownload}
                className="w-full bg-red-600 hover:bg-red-700 text-white font-medium py-3 px-4 rounded-xl transition-colors flex items-center justify-center gap-2"
              >
                <svg
                  className="w-4 h-4"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
                  />
                </svg>
                Descargar TXT
              </button>

              <button
                onClick={handleReset}
                className="mt-3 text-sm text-gray-400 hover:text-gray-600"
              >
                Procesar otro archivo
              </button>
            </div>
          )}
        </div>

        <p className="text-center text-xs text-gray-300 mt-6">
          El archivo se procesa localmente — nunca sale de tu navegador
        </p>
      </div>
    </main>
  );
}
