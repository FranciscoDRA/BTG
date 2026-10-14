// Parser de extractos BTG Pactual (iBanca) -> movimientos. Corre en el navegador y en Node.
const FECHA = /^\d{2}\/\d{2}\/\d{4}$/;
const MONTO = /^-?\d{1,3}(?:\.\d{3})*,\d{2}$/;
const ENTERO = /^-?\d{1,3}(?:\.\d{3})*$/;

const aCentavos = (t) => Math.round(parseFloat(t.replace(/\./g, "").replace(",", ".")) * 100);

/** Agrupa los fragmentos de texto de una página en líneas (por coordenada Y). */
function lineas(items) {
  const filas = [];
  for (const it of items) {
    const s = it.str.trim();
    if (!s) continue;
    const x = it.transform[4], y = it.transform[5];
    let f = filas.find((r) => Math.abs(r.y - y) < 3);
    if (!f) filas.push((f = { y, items: [] }));
    f.items.push({ x, s });
  }
  filas.sort((a, b) => b.y - a.y);
  return filas.map((f) => f.items.sort((a, b) => a.x - b.x).map((i) => i.s));
}

/** pdfjsLib: el módulo pdf.js ya cargado. data: Uint8Array del PDF. */
export async function parsearPdf(pdfjsLib, data) {
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const movs = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const { items } = await page.getTextContent();
    for (const partes of lineas(items)) {
      // separar tokens: algunos fragmentos pueden traer fecha + texto juntos
      const tokens = partes.flatMap((s) => s.split(/\s{2,}/));
      const first = tokens[0].split(/\s+/)[0];
      if (!FECHA.test(first)) continue;
      const resto = [tokens[0].slice(first.length).trim(), ...tokens.slice(1)].filter(Boolean);
      // Montos completos (con centavos). El último es el importe salvo que haya saldo completo después.
      const montosIdx = resto.map((t, i) => (MONTO.test(t) ? i : -1)).filter((i) => i >= 0);
      if (!montosIdx.length) continue;
      let iImp, saldo = null, saldoEntero = null;
      const ultimo = resto.length - 1;
      if (montosIdx.length >= 2 && montosIdx[montosIdx.length - 1] === ultimo) {
        iImp = montosIdx[montosIdx.length - 2];
        saldo = aCentavos(resto[ultimo]);
      } else {
        // Saldo recortado por el borde de la hoja (ej. "-2.391" en vez de "-2.391,07") o ausente
        iImp = montosIdx[montosIdx.length - 1];
        const sig = resto[iImp + 1];
        if (sig && ENTERO.test(sig)) saldoEntero = parseInt(sig.replace(/\./g, ""), 10);
      }
      const imp = aCentavos(resto[iImp]);
      const [d, m, a] = first.split("/").map(Number);
      movs.push({
        fecha: new Date(Date.UTC(a, m - 1, d)),
        concepto: resto.slice(0, iImp).join(" ").replace(/\s+/g, " ").trim(),
        debito: imp < 0 ? -imp : 0,     // en centavos
        credito: imp > 0 ? imp : 0,
        saldo,                          // en centavos, o null si está recortado
        saldoEntero,                    // pesos sin centavos cuando el saldo está recortado
      });
    }
  }
  return movs;
}

/**
 * Devuelve { errores: [nº de fila...], aproximado: bool }.
 * Con saldo completo: saldo anterior - débito + crédito = saldo (exacto).
 * Con saldo recortado: compara solo la parte entera (tolerancia de 2 pesos).
 */
export function validar(movs) {
  const errores = [];
  let aproximado = false;
  const entero = (m) => (m.saldo !== null ? Math.trunc(m.saldo / 100) : m.saldoEntero);
  for (let i = 1; i < movs.length; i++) {
    const a = movs[i - 1], b = movs[i], neto = b.credito - b.debito;
    if (a.saldo !== null && b.saldo !== null) {
      if (a.saldo + neto !== b.saldo) errores.push(i + 1);
    } else if (entero(a) !== null && entero(b) !== null) {
      aproximado = true;
      if (Math.abs((entero(b) - entero(a)) * 100 - neto) >= 200) errores.push(i + 1);
    } else {
      aproximado = true;
    }
  }
  return { errores, aproximado };
}

/** Genera el .xlsx (ArrayBuffer) con ExcelJS. */
export async function aExcel(ExcelJS, movs) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.columns = [
    { header: "Fecha", key: "fecha", width: 21, style: { numFmt: "dd/mm/yy" } },
    { header: "Concepto", key: "concepto", width: 37 },
    { header: "Débito", key: "debito", width: 12, style: { numFmt: "0.00" } },
    { header: "Crédito", key: "credito", width: 12, style: { numFmt: "0.00" } },
  ];
  for (const m of movs) ws.addRow({ ...m, debito: m.debito / 100, credito: m.credito / 100 });
  ws.getRow(1).eachCell((c) => { c.font = { bold: true }; c.alignment = { horizontal: "center" }; c.numFmt = "General"; });
  return wb.xlsx.writeBuffer();
}
