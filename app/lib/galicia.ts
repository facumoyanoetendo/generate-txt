/**
 * Generador de TXT para Pago de Haberes - Banco Galicia
 * Port de generar_txt_galicia.py
 */

import * as XLSX from "xlsx";

// --- Tablas de mapeo ---

const CONVENIO_CODE: Record<string, string> = {
  "Pago de Haberes": "*H3",
  "Pago a Proveedores": "*H3",
};

const TIPO_CUENTA: Record<string, string> = {
  "Caja de Ahorro": "A",
  "Cuenta Corriente": "C",
  "Cuenta Cese Laboral": "A",
};

const MONEDA: Record<string, string> = {
  Pesos: "1",
  Dólares: "2",
};

const CONSOLIDADO: Record<string, string> = {
  Consolidado: "S",
  "No Consolidado": "N",
};

const CONCEPTO: Record<string, string> = {
  "Acreditamiento Haberes": "01",
  "Horas Extra": "02",
  "Reintegro por Viáticos": "03",
  "Sueldo Anual Complementario": "04",
  "Subsidio Vacacional": "05",
  "Gastos de Representacion": "06",
  "Honorarios de Profesionales": "07",
  "Asignacion Personal Contratado": "08",
  "Asignacion Becas/Pasantias": "09",
  "Premio por Productividad/Calidad": "10",
  "Reembolso Gastos": "11",
  "Indemnizacion/Liquidacion Final": "12",
};

// --- Funciones de formato ---

function valStr(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    // Check if it's a date serial (XLSX date)
    return String(Math.trunc(value));
  }
  return String(value).trim();
}

function dateStr(value: unknown): string {
  if (value === null || value === undefined) return "";
  // SheetJS can give us a Date object if cellDates: true
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}${m}${d}`;
  }
  // Fallback: treat as string
  return valStr(value);
}

function dateLong(value: unknown): string {
  // DDMMYYYY for filename
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${d}${m}${y}`;
  }
  return valStr(value);
}

function padRight(value: unknown, width: number): string {
  const s = valStr(value);
  return s.slice(0, width).padEnd(width, " ");
}

function padLeftZeros(value: unknown, width: number): string {
  const s = valStr(value);
  return s.padStart(width, "0").slice(-width);
}

function amountStr(value: unknown, width = 14): string {
  if (!value && value !== 0) return "0".repeat(width);
  const cents = Math.round(Number(value) * 100);
  return String(cents).padStart(width, "0");
}

function cbuStr(value: unknown, width = 26): string {
  if (value === null || value === undefined || value === "") {
    return "0".repeat(width);
  }
  const digits = String(value)
    .trim()
    .replace(/\D/g, "");
  return digits.padStart(width, "0").slice(-width);
}

// --- Generación de registros ---

function makeHeader(ws: XLSX.WorkSheet): string {
  const cell = (addr: string) => ws[addr]?.v;

  const tipoConv = valStr(cell("B2"));
  const convenio = padLeftZeros(cell("B3"), 6);
  const cuit = padLeftZeros(cell("B4"), 11);
  const tipoCta = TIPO_CUENTA[valStr(cell("B5"))] ?? "C";
  const moneda = MONEDA[valStr(cell("B6"))] ?? "1";
  const ctaDeb = padLeftZeros(cell("B7"), 12);
  const leyenda = padRight(valStr(cell("B8")).toUpperCase(), 15);
  const consol = CONSOLIDADO[valStr(cell("B9"))] ?? "N";
  const importe = amountStr(cell("B10"), 14);
  const fecha = dateStr(cell("B11"));
  const convCode = CONVENIO_CODE[tipoConv] ?? "*H3";

  const line =
    convCode +       // pos   1-3  (3)
    convenio +       // pos   4-9  (6)
    cuit +           // pos  10-20 (11)
    tipoCta +        // pos  21    (1)
    moneda +         // pos  22    (1)
    ctaDeb +         // pos  23-34 (12)
    "0".repeat(26) + // pos  35-60 (26)  CBU Débito → ceros
    importe +        // pos  61-74 (14)
    fecha +          // pos  75-82 (8)
    leyenda +        // pos  83-97 (15)
    consol +         // pos  98    (1)
    " ".repeat(379); // pos  99-477 (379)

  if (line.length !== 477) {
    throw new Error(`Header length error: ${line.length} (expected 477)`);
  }
  return line;
}

function makeDetail(cols: unknown[]): string {
  const nombre = cols[0];
  const cuitEmp = cols[1];
  const fechaAcred = cols[2];
  const tipoCta = cols[3];
  const monedaC = cols[4];
  const cuentaCred = cols[5];
  const cbuCred = cols[6];
  const importe = cols[7];
  const leyendaCred = cols[8];
  const idInterna = cols[9];
  const fechaProc = cols[10];
  const concepto = cols[11];
  const pagoCom = cols[12];
  const nroVep = cols[13];
  const leyendaDeb = cols[14];
  const periodo = cols[15];

  const nombreS = padRight(nombre, 16);
  const cuitS = padLeftZeros(cuitEmp, 11);
  const fecAcredS = dateStr(fechaAcred);
  const tipoCts = TIPO_CUENTA[valStr(tipoCta)] ?? "A";
  const monedaS = MONEDA[valStr(monedaC)] ?? "1";
  const cuentaS = cbuStr(cuentaCred, 12);
  const cbuS = cbuStr(cbuCred, 26);
  const importeS = amountStr(importe, 14);
  const leyCred = padRight(leyendaCred, 15);
  const idInt = padRight(idInterna, 22);
  const fecProc = dateStr(fechaProc);
  const conceptoS = (CONCEPTO[valStr(concepto)] ?? "01").padStart(2, "0");
  const pagoComS = padRight(pagoCom, 2);
  const nroVepS = padRight(nroVep, 14);
  const emailS = " ".repeat(60);
  const leyDebS = padRight(valStr(leyendaDeb).toUpperCase(), 35);
  const periodoS = padRight(periodo, 15);

  const line =
    nombreS +       // pos   1-16  (16)
    cuitS +         // pos  17-27  (11)
    fecAcredS +     // pos  28-35  (8)
    tipoCts +       // pos  36     (1)
    monedaS +       // pos  37     (1)
    cuentaS +       // pos  38-49  (12)
    cbuS +          // pos  50-75  (26)
    "32" +          // pos  76-77  (2)   Código Transacción (fijo)
    "1" +           // pos  78     (1)   Tipo de Pago (fijo)
    importeS +      // pos  79-92  (14)
    leyCred +       // pos  93-107 (15)
    idInt +         // pos 108-129 (22)
    fecProc +       // pos 130-137 (8)
    conceptoS +     // pos 138-139 (2)
    pagoComS +      // pos 140-141 (2)
    nroVepS +       // pos 142-155 (14)
    emailS +        // pos 156-215 (60)
    leyDebS +       // pos 216-250 (35)
    periodoS +      // pos 251-265 (15)
    " ".repeat(212);// pos 266-477 (212)

  if (line.length !== 477) {
    throw new Error(`Detail length error: ${line.length} (expected 477)`);
  }
  return line;
}

function makeTrailer(convenio: string, count: number): string {
  const line =
    "*F" +                              // pos  1-2  (2)
    String(convenio).padStart(6, "0") + // pos  3-8  (6)
    String(count).padStart(7, "0") +    // pos  9-15 (7)
    " ".repeat(462);                    // pos 16-477 (462)

  if (line.length !== 477) {
    throw new Error(`Trailer length error: ${line.length} (expected 477)`);
  }
  return line;
}

// --- Tipos de resultado ---

export interface GaliciaResult {
  filename: string;
  content: Uint8Array;
  employeeCount: number;
  totalAmount: number;
  tipoConv: string;
}

// --- Función principal ---

export function generateTxt(workbook: XLSX.WorkBook): GaliciaResult {
  const ws = workbook.Sheets["Transferencias y Pagos"];
  if (!ws) {
    throw new Error(
      'No se encontró la hoja "Transferencias y Pagos" en el archivo'
    );
  }

  const cell = (addr: string) => ws[addr]?.v;

  const tipoConv = valStr(cell("B2"));
  const convenio = valStr(cell("B3"));
  const fechaB11 = cell("B11");
  const fechaLarga = dateLong(fechaB11);
  const totalAmount = Number(cell("B10")) || 0;

  // Header
  const header = makeHeader(ws);

  // Empleados (filas 15 en adelante, col A con nombre)
  const employees: unknown[][] = [];
  let row = 15;
  while (true) {
    const nombre = ws[`A${row}`]?.v;
    if (nombre === null || nombre === undefined) break;
    const cols: unknown[] = [];
    for (let c = 1; c <= 16; c++) {
      const addr = XLSX.utils.encode_cell({ r: row - 1, c: c - 1 });
      cols.push(ws[addr]?.v ?? null);
    }
    employees.push(cols);
    row++;
  }

  if (employees.length === 0) {
    throw new Error("No se encontraron empleados en el archivo (fila 15+)");
  }

  // Líneas de detalle
  const details = employees.map(makeDetail);

  // Trailer
  const trailer = makeTrailer(convenio, details.length);

  // Nombre del archivo de salida
  const filename = `${tipoConv}_${convenio}_${fechaLarga}.txt`;

  // Construir contenido: UTF-8 BOM + líneas con CRLF
  const lines = [header, ...details, trailer];
  const crlf = "\r\n";
  const body = lines.join(crlf) + crlf;
  const bom = new Uint8Array([0xef, 0xbb, 0xbf]);
  const bodyBytes = new TextEncoder().encode(body);
  const content = new Uint8Array(bom.length + bodyBytes.length);
  content.set(bom, 0);
  content.set(bodyBytes, bom.length);

  return {
    filename,
    content,
    employeeCount: employees.length,
    totalAmount,
    tipoConv,
  };
}
