// Genera el informe Word del caso Cafés Monte Bianco a partir de datos.json (extraído del Excel recalculado).
// Uso: node construir_informe.js <carpeta_con_datos_y_graficos> <salida.docx>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Header, Footer,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber,
  PageBreak,
} = require("docx");

const dir = process.argv[2];
const salida = process.argv[3];
const d = JSON.parse(fs.readFileSync(path.join(dir, "datos.json"), "utf8"));

// ---------- formato numérico (es) ----------
function agrupa(entero) {
  return String(entero).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}
function n0(x) {
  if (x === null || x === undefined) return "";
  const r = Math.round(x);
  if (r === 0) return "–";
  return r < 0 ? `(${agrupa(-r)})` : agrupa(r);
}
function nn(x) { // con signo menos, para texto corrido
  const r = Math.round(x);
  return (r < 0 ? "–" : "") + agrupa(Math.abs(r));
}
function dec(x, k = 2) {
  const s = Math.abs(x).toFixed(k).split(".");
  return (x < 0 ? "–" : "") + agrupa(s[0]) + (k ? "," + s[1] : "");
}
function pct(x, k = 1) { return dec(x * 100, k) + " %"; }
function veces(x) { return dec(x, 2) + "x"; }
function kg(x) { return agrupa(Math.round(x)); }

// ---------- estilos ----------
const FUENTE = "Arial";
const AZUL_OSC = "1F3864";
const GRIS = "595959";
const ANCHO = 9360; // carta con márgenes de 1"
const borde = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };

function p(texto, opts = {}) {
  const runs = Array.isArray(texto) ? texto : [texto];
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    ...opts.par,
    children: runs.map((t) => (typeof t === "string" ? new TextRun({ text: t }) : t)),
  });
}
const b = (t) => new TextRun({ text: t, bold: true });
const it = (t) => new TextRun({ text: t, italics: true });
function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] }); }
function bullet(texto, nivel = 0) {
  const runs = Array.isArray(texto) ? texto : [texto];
  return new Paragraph({
    numbering: { reference: "viñetas", level: nivel },
    spacing: { after: 80, line: 264 },
    children: runs.map((t) => (typeof t === "string" ? new TextRun({ text: t }) : t)),
  });
}
function numerada(texto) {
  const runs = Array.isArray(texto) ? texto : [texto];
  return new Paragraph({
    numbering: { reference: "numeros", level: 0 },
    spacing: { after: 100, line: 264 },
    children: runs.map((t) => (typeof t === "string" ? new TextRun({ text: t }) : t)),
  });
}
function nota(t) {
  return new Paragraph({
    spacing: { before: 60, after: 160 },
    children: [new TextRun({ text: t, italics: true, size: 16, color: GRIS })],
  });
}
function figura(archivo, ancho, alto, pie) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 60 },
      children: [new ImageRun({
        type: "png", data: fs.readFileSync(path.join(dir, archivo)),
        transformation: { width: ancho, height: alto },
        altText: { title: pie, description: pie, name: archivo },
      })],
    }),
    nota(pie),
  ];
}

// tabla: filas = [{c: [..], bold, shade}] ; la primera columna es texto, el resto numérico
function tabla(encab, filas, anchos, colsTexto = [0]) {
  const total = anchos.reduce((a, x) => a + x, 0);
  const celda = (texto, w, { bold = false, head = false, izq = false, shade = null } = {}) => new TableCell({
    borders: bordes,
    width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: AZUL_OSC, type: ShadingType.CLEAR, color: "auto" }
      : shade ? { fill: shade, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({
      alignment: izq ? AlignmentType.LEFT : (head ? AlignmentType.CENTER : AlignmentType.RIGHT),
      children: [new TextRun({ text: String(texto), bold: bold || head, size: 17, color: head ? "FFFFFF" : undefined })],
    })],
  });
  const rows = [new TableRow({
    tableHeader: true,
    children: encab.map((t, i) => celda(t, anchos[i], { head: true })),
  })];
  for (const f of filas) {
    if (f.seccion) {
      rows.push(new TableRow({
        children: [new TableCell({
          borders: bordes, columnSpan: anchos.length, width: { size: total, type: WidthType.DXA },
          shading: { fill: "D9E1F2", type: ShadingType.CLEAR, color: "auto" },
          margins: { top: 40, bottom: 40, left: 90, right: 90 },
          children: [new Paragraph({ children: [new TextRun({ text: f.seccion, bold: true, size: 17 })] })],
        })],
      }));
      continue;
    }
    rows.push(new TableRow({
      children: f.c.map((t, i) => celda(t, anchos[i], { bold: f.bold, izq: colsTexto.includes(i) || f.izq, shade: f.shade })),
    }));
  }
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: anchos, rows });
}
const espacio = () => new Paragraph({ spacing: { after: 60 }, children: [] });

// ---------- cifras ----------
const [U19, UMP, UMX] = d.un;
const roe = d.r_roe;
const i19 = 0, iMP = 1, iMX = 2;
const ventas = d.ventas;
const levRoePts = roe[0] - d.d_rona;
const mini = d.mini;
const res = mini["Resultado"];
const unMini = mini["Utilidad neta (40 % de impuestos)"];
const iA = d.mix_grados.indexOf(d.opt_cal);
const iAAA = d.mix_grados.indexOf("AAA");
const mej = d.mejoras; // filas: [palanca, dUN, dCaja, dDeuda, UN, excedente, ROE]
const mejTodas = mej[5];

// ---------- contenido ----------
const c = [];

// Portada
c.push(new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.LEFT,
  children: [new TextRun({ text: "Cafés Monte Bianco", bold: true, size: 56, color: AZUL_OSC })] }));
c.push(new Paragraph({ spacing: { after: 400 },
  children: [new TextRun({ text: "La elaboración de un plan de ganancias para 2020", size: 32, color: GRIS })] }));
c.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: AZUL_OSC, space: 4 } }, children: [] }));
c.push(new Paragraph({ spacing: { before: 300, after: 120 },
  children: [new TextRun({ text: "Informe de análisis financiero y estratégico", size: 26 })] }));
c.push(p([new TextRun({ text: "Caso HBS 204-S18 (Simons y Dávila). Cifras en miles de euros salvo indicación. "
  + "El modelo completo, con todas las fórmulas, está en el archivo Monte_Bianco_Analisis.xlsx.", color: GRIS })],
  { align: AlignmentType.LEFT }));
c.push(new Paragraph({ children: [new PageBreak()] }));

// 1. Resumen ejecutivo
c.push(h1("1. Resumen ejecutivo"));
c.push(p(["La pregunta de Giacomo Salvetti es si conviene llenar la planta ampliada (6.000.000 kg/año) solo con marcas "
  + "propias de supermercados. La respuesta del análisis es ", b("no"), ": esa estrategia destruye utilidad y "
  + "agota la caja. La mejor alternativa es una ", b("mezcla"), " de café de calidad ", b(d.opt_cal), " con "
  + pct(d.opt_m, 0) + " de gasto de marketing sobre ventas, que ocupa la capacidad restante con marca propia."]));
c.push(bullet([b("2019 no fue “excelente”. "), `La utilidad neta fue de ${nn(U19)} (margen neto de ${pct(d.r_mn[0])}). `
  + `Los intereses se llevaron el ${pct(d.d_int_ebit, 0)} de la utilidad operativa, la línea de crédito de 20.000 está `
  + `usada al 100 % y la razón corriente es ${veces(d.d_rc)}.`]));
c.push(bullet([b(`El ROE de ${pct(roe[0])} es engañoso. `), `Se explica por el apalancamiento (activos/patrimonio de `
  + `${veces(d.r_mult[0])}). La rentabilidad operativa después de impuestos es de solo ${pct(d.d_rona)}; unos `
  + `${dec(levRoePts * 100, 0)} puntos del ROE provienen de la deuda.`]));
c.push(bullet([b("Solo marcas propias: "), `ventas de ${nn(ventas[iMP])}, utilidad neta de ${nn(UMP)} (–${pct(1 - UMP / U19, 0)} `
  + `frente a 2019), ROE de ${pct(roe[iMP])} y cobertura de intereses de ${veces(d.r_cob[iMP])}. Las cuentas a cobrar se `
  + `duplican (${nn(d.cxc[iMP])}) y faltan ${nn(-d.nec[iMP])} de financiamiento, con la línea de crédito ya agotada.`]));
c.push(bullet([b("Solo AAA: "), `el mercado absorbe como máximo ${kg(d.solo_q[iAAA])} kg; la planta quedaría ociosa en un `
  + `${pct(d.solo_ocio[iAAA], 0)} y, después de gastos de estructura e intereses, la empresa perdería ${nn(-unMini[1])}.`]));
c.push(bullet([b(`Mix óptimo (calidad ${d.opt_cal} al ${pct(d.opt_m, 0)} + marca propia): `), `${kg(d.mix_qpr[iA])} kg de calidad `
  + `${d.opt_cal} y ${kg(d.mix_qmp[iA])} kg de marca propia. Resultado del mini estado de ${nn(d.opt_res)} frente a `
  + `${nn(d.mp_solo_mini)} con solo marca propia; utilidad neta proyectada de ${nn(UMX)} y ROE de ${pct(roe[iMX])} con `
  + `mucho menos apalancamiento (${veces(d.r_mult[iMX])}).`]));
c.push(bullet([b("Recomendación: "), "adoptar el mix, firmar contratos de marca propia en bloques de 500.000 kg para "
  + "la capacidad restante, negociar plazos de cobro, vender el lote desocupado para reducir deuda y gestionar la "
  + "liquidez con un presupuesto de caja mensual (sección 7)."]));

// 2. Diagnóstico 2019
c.push(h1("2. Diagnóstico de la situación actual (2019)"));
c.push(h2("2.1 ¿Fue 2019 un año “excelente”?"));
c.push(p(`No. El estado de resultados muestra un negocio con buen margen bruto pero con poca utilidad final. De cada euro `
  + `vendido quedan ${dec(d.r_ebit[0] * 100, 1)} céntimos de utilidad operativa y solo ${dec(d.r_mn[0] * 100, 1)} céntimos de `
  + `utilidad neta. La diferencia se va en intereses: ${nn(d.int[0])}, es decir, el ${pct(d.d_int_ebit, 0)} del EBIT. Además, `
  + `los estados son estimaciones al 15 de diciembre, no cifras auditadas.`));
c.push(tabla(["Indicador 2019", "Valor", "Lectura"], [
  { c: ["Ventas", n0(ventas[0]), "Marca propia 17,7 %, calidad superior 82,3 %"], },
  { c: ["Margen bruto", pct(d.r_mb[0]), "Sano gracias a la calidad superior"] },
  { c: ["Margen operativo (EBIT/ventas)", pct(d.r_ebit[0]), "Razonable"] },
  { c: ["Intereses / EBIT", pct(d.d_int_ebit), "Más de la mitad del EBIT se va en intereses"] },
  { c: ["Cobertura de intereses", veces(d.r_cob[0]), "Menos de 2x: riesgo financiero alto"] },
  { c: ["Margen neto", pct(d.r_mn[0]), "Muy bajo"] },
  { c: ["Deuda financiera / EBITDA", veces(d.d_dbt), "Endeudamiento elevado"] },
  { c: ["Razón corriente", veces(d.d_rc), "Pasivo corriente mayor que activo corriente"] },
  { c: ["Línea de crédito disponible", "0", "20.000 de 20.000 utilizados"] },
  { c: ["Lote desocupado / activos", pct(d.d_lote, 0), "Activo improductivo de 10.000 financiado con deuda"] },
  { c: ["Uso de la capacidad (350.000 kg/mes)", pct(d.d_uso19), `Con la ampliación bajaría a ${pct(d.d_uso20)}`] },
], [3400, 1500, 4460], [0, 2]));
c.push(nota("Fuente: Anexo 1 y hoja Diagnostico_2019 del Excel. El total del activo es 35.515 (el PDF muestra 33.515 por error de tipeo; el patrimonio de 7.127 solo cuadra con 35.515)."));

c.push(h2("2.2 ¿Por qué el ROE es tan alto?"));
c.push(p(`El ROE de ${pct(roe[0])} no refleja un negocio muy rentable sino uno muy endeudado. La descomposición DuPont lo muestra:`));
c.push(tabla(["Componente", "2019"], [
  { c: ["Margen neto (utilidad neta / ventas)", pct(d.r_mn[0])] },
  { c: ["× Rotación de activos (ventas / activos)", veces(d.r_rot[0])] },
  { c: ["× Multiplicador de capital (activos / patrimonio)", veces(d.r_mult[0])] },
  { c: ["= ROE", pct(roe[0])], bold: true },
], [6360, 3000]));
c.push(espacio());
c.push(p(["Con 7.127 de patrimonio la empresa sostiene 35.515 de activos: por cada euro de los accionistas hay casi "
  + "cuatro euros de deuda (D/E = " + veces(d.d_de) + "). La fórmula del apalancamiento financiero separa los dos efectos:"]));
c.push(p([it("ROE = RONA + (RONA – kd·(1–t)) × D/E"), ` = ${pct(d.d_rona)} + (${pct(d.d_rona)} – ${pct(d.d_kd)}) × `
  + `${dec(d.d_de, 2)} = ${pct(roe[0])}`], { align: AlignmentType.CENTER }));
c.push(p(`Sin deuda, el ROE sería igual al RONA: ${pct(d.d_rona)}. Unos ${dec(levRoePts * 100, 0)} puntos del ROE se deben a `
  + `que el negocio rinde un poco más que el costo de la deuda después de impuestos (${pct(d.d_kd)}) y a que ese pequeño `
  + `diferencial se multiplica por una deuda cuatro veces mayor que el patrimonio. Es un ROE frágil: si el rendimiento `
  + `operativo cae por debajo del ${pct(d.d_kd)}, el apalancamiento juega en contra y el ROE se desploma. Además, el `
  + `denominador es pequeño porque el patrimonio fue erosionado por años de bajas utilidades.`));

c.push(h2("2.3 Necesidades de capital de trabajo"));
c.push(p(`El ciclo de conversión de efectivo es de ${dec(d.d_ccc, 0)} días: la empresa cobra en ${dec(d.d_dcxc, 0)} días en `
  + `promedio, mantiene ${dec(d.d_dinv, 0)} días de materia prima y ${dec(d.d_dpt, 0)} días de producto terminado, pero paga `
  + `a sus proveedores en apenas ${dec(d.d_dpag, 0)} días. El capital de trabajo operativo es de ${nn(d.d_ctno)}, el `
  + `${pct(d.d_ctnov, 0)} de las ventas: cada euro adicional de ventas exige unos ${dec(d.d_ctnov * 100, 0)} céntimos de `
  + `financiamiento.`));
c.push(p(`El problema es cómo se financia. La línea de crédito de corto plazo (20.000) financia tanto el capital de trabajo `
  + `como parte de la planta y un lote desocupado de 10.000. La razón corriente es de ${veces(d.d_rc)} y la prueba ácida de `
  + `${veces(d.d_acida)}. Hay un descalce de plazos: activos de largo plazo con deuda de corto plazo y sin margen para `
  + `crecer.`));
c.push(p(`Las ventas son estacionales (Anexo 4): octubre a diciembre concentran el 37,2 % del año y el verano es muy flojo. `
  + `Como los minoristas pagan a 90 días, la cuenta por cobrar de marca propia al cierre equivale a las ventas de octubre, `
  + `noviembre y diciembre, no a 90/365 de las ventas anuales. La cuenta por cobrar estimada con esta regla (${nn(d.d_cxc_est)}) `
  + `es muy parecida a la del balance (7.449), lo que valida el supuesto para las proyecciones.`));

// 3. Solo marcas propias
c.push(h1("3. Escenario 2020: concentrarse solo en marcas propias"));
c.push(h2("3.1 Supuestos"));
for (const s of [
  "Ventas de 6.000.000 kg a 7,00 €/kg (capacidad plena de 500.000 kg/mes, Anexo 5); el inventario inicial de 174.000 kg se vende y el inventario final de producto terminado es cero.",
  "Costo variable de 5,25 €/kg y costos fijos de manufactura de 2.641 (Anexo 3). Se supone que esos costos fijos ya incluyen la depreciación base de 2.081, igual que el costo de ventas de 2019. Se suma la depreciación incremental de la ampliación: 4.500 / 15 = 300.",
  "Depreciación total de 2020: 2.081 + 300 = 2.381, que se suma a la depreciación acumulada (9.758 → 12.139). La ampliación de 4.500 ya está en los bienes de uso de 2019, así que no hay nuevas inversiones.",
  "Ahorros de Carla: marketing –100 %, ventas –65 %, I+D –75 % y administración –50 %.",
  "Cuentas a cobrar: ventas de octubre, noviembre y diciembre (pago a 90 días y estacionalidad del Anexo 4).",
  "Inventario de materia prima y cuentas a pagar: misma rotación de 2019 sobre el costo de ventas sin depreciación.",
  "Endeudamiento constante (20.000 de línea de crédito + 8.000 de largo plazo), con la misma tasa implícita de 2019 (10,9 %). La caja es la variable de cierre del balance y se compara con una caja mínima operativa (2 % de las ventas, como en 2019). Impuestos del 40 % y sin dividendos.",
]) c.push(bullet(s));

const filaER = (et, k, bold = false) => ({ c: [et, n0(d[k][0]), n0(d[k][1])], bold });
c.push(h2("3.2 Estado de resultados"));
c.push(tabla(["Miles de €", "2019 real", "2020 solo marca propia"], [
  filaER("Ventas", "ventas", true),
  filaER("Costo de mercadería vendida", "cmv"),
  filaER("Margen bruto", "mb", true),
  filaER("Marketing", "mkt"), filaER("I+D", "id"), filaER("Gastos de venta", "vta"), filaER("Administración", "adm"),
  filaER("Utilidad operativa (EBIT)", "ebit", true),
  filaER("Intereses", "int"), filaER("Utilidad antes de impuestos", "uai"), filaER("Impuestos (40 %)", "imp"),
  filaER("Utilidad neta", "un", true),
  { c: ["Margen bruto %", pct(d.r_mb[0]), pct(d.r_mb[1])] },
  { c: ["Margen neto %", pct(d.r_mn[0]), pct(d.r_mn[1])] },
], [4960, 2200, 2200]));
c.push(nota("Costo de ventas 2020 = 6.000.000 × 5,25 + 2.641 + 300 = 34.441."));

c.push(h2("3.3 Balance general al 31 de diciembre"));
c.push(tabla(["Miles de €", "2019 real", "2020 solo marca propia"], [
  filaER("Caja", "caja"), filaER("Inventario de materia prima", "invmp"), filaER("Bienes terminados", "invpt"),
  filaER("Cuentas a cobrar", "cxc"), filaER("Activo corriente", "ac", true),
  filaER("Bienes de uso (bruto)", "ppe"), filaER("Depreciación acumulada", "depac"), filaER("Bienes de uso netos", "ppen"),
  filaER("Total activos", "ta", true),
  filaER("Cuentas a pagar", "cxp"), filaER("Línea de crédito", "lc"), filaER("Deuda a largo plazo", "dlp"),
  filaER("Total pasivos", "tp", true), filaER("Patrimonio neto", "pat", true),
  filaER("Total pasivos y patrimonio", "tpp", true),
  { c: ["Caja mínima operativa", n0(d.cajamin[0]), n0(d.cajamin[1])] },
  { c: ["Excedente / (necesidad) de financiamiento", n0(d.nec[0]), n0(d.nec[1])], bold: true, shade: "FBE5D6" },
], [4960, 2200, 2200]));
c.push(nota("Con la deuda congelada, la caja resulta negativa: es el monto que habría que conseguir para cerrar el balance."));

c.push(h2("3.4 Estado de flujo de efectivo 2020"));
c.push(tabla(["Miles de €", "2020 solo marca propia"], [
  { c: ["Utilidad neta", n0(d.un[1])] },
  { c: ["(+) Depreciación", n0(d.f_dep[1])] },
  { c: ["(–) Aumento de cuentas a cobrar", n0(d.f_cxc[1])] },
  { c: ["(–) Aumento de inventario de materia prima", n0(d.f_imp[1])] },
  { c: ["(+) Disminución de bienes terminados", n0(d.f_ipt[1])] },
  { c: ["(+) Aumento de cuentas a pagar", n0(d.f_cxp[1])] },
  { c: ["Flujo de caja operativo", n0(d.fco[1])], bold: true },
  { c: ["Flujo de caja de inversión", n0(d.fci[1])] },
  { c: ["Flujo de caja de financiamiento (deuda constante, sin dividendos)", n0(d.fcf[1])] },
  { c: ["Variación neta de caja", n0(d.dcaja[1])], bold: true },
  { c: ["Caja inicial", n0(d.cini[1])] },
  { c: ["Caja final", n0(d.cfin[1])], bold: true, shade: "FBE5D6" },
], [6360, 3000]));

c.push(h2("3.5 ¿Qué nos dicen estos estados financieros?"));
c.push(p([b("Rueda de ganancias. "), `Las ventas caen a ${nn(ventas[1])} aunque se venden 2,6 veces más kilos, porque el `
  + `precio promedio baja de unos 19 €/kg a 7 €/kg. El margen bruto cae de ${pct(d.r_mb[0], 0)} a ${pct(d.r_mb[1], 0)} y el EBIT `
  + `por kilo de ${dec(d.r_ebkg[0])} € a ${dec(d.r_ebkg[1])} €. Los ahorros de Carla (${nn(d.gop[0] - d.gop[1])}) no compensan `
  + `la pérdida de margen bruto (${nn(d.mb[0] - d.mb[1])}). La utilidad neta cae a ${nn(UMP)}.`]));
c.push(p([b("Rueda de caja. "), `Es la parte más grave. Las cuentas a cobrar pasan de ${nn(d.cxc[0])} a ${nn(d.cxc[1])} `
  + `(${dec(d.r_dcxc[1], 0)} días de ventas) porque todo se cobra a 90 días y el último trimestre es el más fuerte del año. `
  + `El flujo operativo es negativo (${nn(d.fco[1])}) y la caja termina en ${nn(d.cfin[1])}. Faltan ${nn(-d.nec[1])} para `
  + `tener una caja mínima operativa, con la línea de crédito ya agotada. Si ese faltante se financiara con deuda al `
  + `10,9 %, los intereses adicionales (${nn(d.intad[1])}) dejarían la utilidad neta en ${nn(d.unaj[1])}.`]));
c.push(p([b("Rueda del ROE. "), `El ROE cae a ${pct(roe[1])} (${pct(d.r_mn[1])} × ${veces(d.r_rot[1])} × ${veces(d.r_mult[1])}). `
  + `La cobertura de intereses baja a ${veces(d.r_cob[1])} y la deuda/EBITDA sube a ${veces(d.r_dbt[1])}. Las ventas de equilibrio `
  + `(incluyendo intereses) son de ${nn(d.r_bei[1])}: el margen de seguridad es de solo ${pct(d.r_ms[1])}. Bastaría una baja de `
  + `${dec(d.uai[1] / 6000)} €/kg en el precio que negocian los supermercados para eliminar toda la utilidad antes de impuestos.`]));

c.push(h2("3.6 ¿Se pueden mejorar los indicadores?"));
c.push(p("Sí, pero no lo suficiente. La hoja Mejoras_MP mide cada palanca por separado:"));
c.push(tabla(["Palanca", "Δ utilidad neta", "Δ caja", "Utilidad neta", "Excedente / (necesidad)", "ROE"],
  mej.map((f, i) => ({ c: [f[0], n0(f[1]), n0(f[2]), n0(f[4]), n0(f[5]), pct(f[6])], bold: i === 5, shade: i === 5 ? "E2EFDA" : null })),
  [3160, 1150, 1150, 1250, 1500, 1150]));
c.push(nota("Palancas: cobro a 60 días; precio +0,25 €/kg; venta del lote de 10.000 a valor libro para reducir la línea de crédito; pago a proveedores a 30 días; inventario de materia prima –30 %."));
c.push(p(`Con todas las palancas juntas la empresa cerraría el faltante de caja y el ROE subiría a ${pct(mejTodas[6])}, pero la `
  + `utilidad neta (${nn(mejTodas[4])}) seguiría por debajo de la del mix óptimo (${nn(UMX)}). Además, las palancas dependen `
  + `de negociar con clientes que, según el caso, “negocian mucho los precios” y cambian de proveedor ante cualquier falla, y `
  + `la mayoría (vender el lote, alargar pagos, reducir inventarios) también se puede aplicar al mix.`));

c.push(h2("3.7 ¿Es una buena decisión estratégica?"));
c.push(p(`No. Concentrarse solo en marcas propias reduce la utilidad un ${pct(1 - UMP / U19, 0)}, convierte un problema de `
  + `liquidez en una crisis de caja, deja a la empresa con un margen de seguridad de ${pct(d.r_ms[1], 0)} frente a pocos `
  + `clientes con mucho poder de negociación y exige un compromiso de largo plazo difícil de revertir. También abandona la `
  + `marca, el conocimiento técnico y las relaciones con productores que dan a Monte Bianco su ventaja. La estabilidad del `
  + `volumen no compensa el bajo margen unitario: 1,75 €/kg de contribución frente a 13,85 €/kg de la calidad A.`));

// 4. Solo AAA
c.push(h1("4. ¿Se puede vender solo calidad Premium AAA?"));
c.push(p(`Técnicamente sí, pero sería un error. Según el Anexo 2, aun con el gasto máximo de marketing (10 % de las ventas) `
  + `el mercado solo compra ${kg(d.solo_q[iAAA])} kg de AAA al año. Con una planta de 6.000.000 kg, la capacidad ociosa `
  + `sería del ${pct(d.solo_ocio[iAAA], 1)}.`));
c.push(tabla(["Miles de €", "Solo marca propia", "Solo AAA (10 % mkt)", `Solo calidad ${d.opt_cal}`, "Mix óptimo"], [
  { c: ["Kilos vendidos", kg(mini["Calidad superior: kilos"][0] + mini["Marca propia: kilos"][0]),
    kg(mini["Calidad superior: kilos"][1]), kg(mini["Calidad superior: kilos"][2]),
    kg(mini["Calidad superior: kilos"][3] + mini["Marca propia: kilos"][3])] },
  { c: ["Capacidad utilizada", ...mini["Capacidad utilizada"].map((x) => pct(x, 0))] },
  { c: ["Ventas", ...mini["Ventas totales"].map(n0)] },
  { c: ["Margen de contribución (p – cv)·q", ...mini["Margen bruto de contribución (p – cv)·q"].map(n0)] },
  { c: ["Costos fijos de manufactura", ...mini["Costos fijos de manufactura"].map(n0)] },
  { c: ["Publicidad y ventas (%M·p·q)", ...mini["Publicidad y ventas (%M·p·q)"].map(n0)] },
  { c: ["Resultado (p – cv)·q – CF – %M·p·q", ...res.map(n0)], bold: true },
  { c: ["Depreciación incremental", ...mini["Depreciación incremental de la ampliación"].map(n0)] },
  { c: ["I+D, ventas y administración", ...mini["I+D, ventas y administración"].map(n0)] },
  { c: ["Intereses", ...mini["Intereses (deuda constante)"].map(n0)] },
  { c: ["Utilidad neta", ...unMini.map(n0)], bold: true, shade: "E2EFDA" },
], [3060, 1575, 1575, 1575, 1575]));
c.push(nota("Solo marca propia aplica los ahorros de Carla; las demás columnas mantienen la estructura de gastos de 2019 porque siguen vendiendo calidad superior."));
c.push(p(`El café AAA tiene el mejor margen por kilo, pero el mercado es muy pequeño y muy sensible al marketing. Los costos `
  + `fijos, la estructura de I+D, ventas y administración y los intereses no se diluyen: la empresa perdería `
  + `${nn(-unMini[1])}. La capacidad ociosa es justamente la razón por la que Monte Bianco empezó a fabricar marcas propias.`));
c.push(...figura("grafico_utilidad.png", 600, 272,
  "Gráfico 1. Utilidad neta 2020 por alternativa (hoja Mix_Produccion y EEFF_2020)."));

// 5. Mix óptimo
c.push(h1("5. El mix óptimo entre calidad superior y marcas propias"));
c.push(h2("5.1 Método"));
c.push(p(["Para cada calidad (C, B, BB, A, AA, AAA) y cada nivel de gasto de marketing (0 % a 10 % de las ventas), el Anexo 2 "
  + "da el volumen que absorbe el mercado. La capacidad restante (6.000.000 kg – q) se llena con marca propia a 7 €/kg. "
  + "Para cada una de las 66 combinaciones se calcula el mini estado de resultados:"]));
c.push(p([it("Resultado = (p – cv)·q  +  (7,00 – 5,25)·(6.000.000 – q)  –  3.262  –  %M·p·q")], { align: AlignmentType.CENTER }));
c.push(p("Se usan los costos unitarios variables y los costos fijos del Anexo 3 (3.262, que ya incluyen los 2.641 de marca "
  + "propia). Siguiendo la simplificación del enunciado, el mix es la marca propia más una sola calidad de alta gama. La "
  + "tabla completa está en la hoja Mix_Produccion (Tablas A, B y C)."));
c.push(...figura("grafico_mix.png", 600, 300,
  "Gráfico 2. Mejor resultado por calidad, con el % de marketing óptimo de cada una. La línea punteada es la alternativa de solo marca propia."));
c.push(h2("5.2 Resultados"));
c.push(tabla(["Calidad", "Precio €/kg", "%M óptimo", "Kg calidad sup.", "Kg marca propia", "Resultado mix", "Solo calidad sup."],
  d.mix_grados.map((g, i) => ({
    c: [g, dec([15.5, 21, 24, 28, 31, 34][i]), pct(d.mix_m[i], 0), kg(d.mix_qpr[i]), kg(d.mix_qmp[i]), n0(d.mix_best[i]), n0(d.solo_best[i])],
    bold: g === d.opt_cal, shade: g === d.opt_cal ? "E2EFDA" : null,
  })), [1000, 1160, 1100, 1500, 1500, 1550, 1550]));
c.push(espacio());
c.push(p([`La mejor combinación es `, b(`calidad ${d.opt_cal} a 28 €/kg con ${pct(d.opt_m, 0)} de marketing`),
  `: ${kg(d.mix_qpr[iA])} kg de calidad superior (${pct(d.opt_peso, 0)} del volumen) y ${kg(d.mix_qmp[iA])} kg de marca propia. `
  + `Su resultado es de ${nn(d.opt_res)}, ${nn(d.opt_res - d.mp_solo_mini)} más que la alternativa de solo marca propia.`]));
c.push(p("¿Por qué gana la calidad A y no la AAA? Porque el negocio de alta gama aporta margen, pero la marca propia aporta "
  + "volumen para cubrir costos fijos, y el costo de oportunidad de cada kilo de alta gama es de solo 1,75 € (lo que dejaría "
  + "la marca propia). La calidad A combina un margen alto (28 – 14,15 = 13,85 €/kg) con un volumen grande (1,8 millones de "
  + "kg). Las calidades AA y AAA tienen mejor margen unitario pero su mercado es demasiado pequeño, y la C y la B venden mucho "
  + "volumen con márgenes bajos. Sobre el marketing: pasar de 7 % a 8 % en la calidad A sube el volumen solo 31.000 kg y no "
  + "paga el gasto extra; el óptimo está donde el último punto de marketing todavía genera más contribución de la que cuesta."));
c.push(p("Comprobación de capacidad mensual (hoja Plan_Mensual): la calidad superior no se puede almacenar, pero su venta "
  + "máxima (marzo, unos 211.000 kg) cabe en los 500.000 kg mensuales. La marca propia sí se almacena y se produce con la "
  + "capacidad libre siguiendo la lógica del Anexo 5, sin faltantes en ningún mes."));

// 6. Comparación de estados
c.push(h1("6. Estados financieros comparados: 2019, solo marca propia y mix óptimo"));
const f3 = (et, k, bold = false, shade = null) => ({ c: [et, n0(d[k][0]), n0(d[k][1]), n0(d[k][2])], bold, shade });
const r3 = (et, k, fmt, bold = false) => ({ c: [et, fmt(d[k][0]), fmt(d[k][1]), fmt(d[k][2])], bold });
c.push(tabla(["Miles de €", "2019 real", "2020 solo marca propia", `2020 mix óptimo (${d.opt_cal} + MP)`], [
  { seccion: "Estado de resultados" },
  f3("Ventas", "ventas", true), f3("  Marca propia", "v_mp"), f3("  Calidad superior", "v_pr"),
  f3("Costo de mercadería vendida", "cmv"), f3("Margen bruto", "mb", true),
  f3("Marketing", "mkt"), f3("I+D", "id"), f3("Gastos de venta", "vta"), f3("Administración", "adm"),
  f3("EBIT", "ebit", true), f3("Intereses", "int"), f3("Impuestos", "imp"), f3("Utilidad neta", "un", true, "E2EFDA"),
  { seccion: "Balance general" },
  f3("Caja", "caja"), f3("Inventario de materia prima", "invmp"), f3("Bienes terminados", "invpt"), f3("Cuentas a cobrar", "cxc"),
  f3("Bienes de uso netos", "ppen"), f3("Total activos", "ta", true), f3("Cuentas a pagar", "cxp"),
  { c: ["Deuda financiera (línea de crédito + largo plazo)", ...[0, 1, 2].map((i) => n0(d.lc[i] + d.dlp[i]))] },
  f3("Patrimonio neto", "pat", true),
  { seccion: "Flujo de efectivo 2020" },
  f3("Flujo de caja operativo", "fco"), f3("Variación neta de caja", "dcaja"), f3("Caja final", "cfin", true),
  f3("Excedente / (necesidad) de financiamiento", "nec", true, "FBE5D6"),
], [3960, 1800, 1800, 1800]));
c.push(nota("Deuda constante en 2020 por supuesto; la caja es la variable de cierre del balance."));
c.push(tabla(["Indicador", "2019 real", "Solo marca propia", "Mix óptimo"], [
  r3("Margen bruto", "r_mb", pct), r3("Margen operativo", "r_ebit", pct), r3("Margen neto", "r_mn", pct),
  r3("Rotación de activos", "r_rot", veces), r3("Multiplicador de capital", "r_mult", veces),
  r3("ROE (patrimonio final)", "r_roe", pct, true), r3("ROA", "r_roa", pct),
  r3("Cobertura de intereses", "r_cob", veces), r3("Deuda financiera / patrimonio", "r_de", veces),
  r3("Deuda financiera / EBITDA", "r_dbt", veces), r3("Días de cobro", "r_dcxc", (x) => dec(x, 0)),
  r3("Capital de trabajo operativo / ventas", "r_ctnov", pct), r3("Razón corriente", "r_rc", veces),
  { c: ["Margen de seguridad (con intereses)", "–", pct(d.r_ms[1]), pct(d.r_ms[2])] },
], [3960, 1800, 1800, 1800]));
c.push(espacio());
c.push(p(`El mix óptimo mejora todas las ruedas a la vez. La utilidad neta sube a ${nn(UMX)}, cinco veces la de 2019. El ROE `
  + `llega a ${pct(roe[2])}, y esta vez se apoya en el margen (${pct(d.r_mn[2])}) y en la rotación (${veces(d.r_rot[2])}), no en `
  + `el apalancamiento, que baja de ${veces(d.r_mult[0])} a ${veces(d.r_mult[2])}. La cobertura de intereses pasa a `
  + `${veces(d.r_cob[2])} y el margen de seguridad a ${pct(d.r_ms[2], 0)}.`));
c.push(p(`Un punto de atención: crecer de ${nn(ventas[0])} a ${nn(ventas[2])} en ventas consume capital de trabajo. Las cuentas `
  + `a cobrar suben a ${nn(d.cxc[2])} y el flujo operativo del primer año es levemente negativo (${nn(d.fco[2])}). La caja `
  + `cierra en ${nn(d.cfin[2])}, positiva pero ${nn(-d.nec[2])} por debajo de la caja mínima operativa. Es una necesidad `
  + `pequeña y transitoria que se cubre con cualquiera de las palancas de la sección 3.6 (por ejemplo, la venta del lote).`));

// 7. Recomendación
c.push(h1("7. Recomendaciones para Giacomo Salvetti"));
for (const r of [
  [b("No pasar al 100 % de marcas propias. "), `Reduce la utilidad a ${nn(UMP)}, deja el ROE en ${pct(roe[1])} y genera un `
    + `faltante de caja de ${nn(-d.nec[1])} que la empresa no puede financiar con la línea de crédito agotada.`],
  [b("No concentrarse solo en AAA. "), `Dejaría el ${pct(d.solo_ocio[iAAA], 0)} de la planta ociosa y generaría pérdidas.`],
  [b(`Adoptar el mix: calidad ${d.opt_cal} con ${pct(d.opt_m, 0)} de marketing y marca propia en la capacidad restante. `),
    `Es la combinación que más dinero genera (${nn(d.opt_res)} en el mini estado y ${nn(UMX)} de utilidad neta). La marca `
    + `propia sigue cumpliendo su papel: cubrir costos fijos y estabilizar la carga de la planta.`],
  [b("Contratar la marca propia en bloques. "), `Los ${kg(d.mix_qmp[iA])} kg de marca propia equivalen a unos ocho `
    + `bloques de 500.000 kg. Conviene firmar contratos con los minoristas que pagan antes y no comprometer toda la capacidad, `
    + `para poder reasignarla si el mercado de alta gama cambia.`],
  [b("Gestionar el capital de trabajo. "), "Negociar cobros a 60 días, pagar a proveedores en plazos de mercado (hoy se paga en "
    + "6 días) y producir marca propia para inventario en verano según el plan mensual."],
  [b("Desapalancar y ordenar el financiamiento. "), "Vender el lote desocupado (10.000) para reducir la línea de crédito y "
    + "refinanciar a largo plazo la parte de la deuda que financia la planta. El ROE debe venir del negocio y no de la deuda."],
  [b("Controlar con presupuesto de caja mensual. "), "La estacionalidad (octubre a diciembre = 37 % de las ventas) obliga a "
    + "planificar la caja mes a mes, no solo al cierre del año."],
]) c.push(numerada(r));

// 8. Supuestos y limitaciones
c.push(h1("8. Supuestos y limitaciones"));
for (const s of [
  "Los volúmenes del Anexo 2 son estimaciones de marketing para un mercado que el propio caso califica de muy volátil. Una recesión como la de 2017 afectaría sobre todo al componente de alta gama del mix.",
  "Se supone que los costos fijos del Anexo 3 ya incluyen la depreciación base de 2.081 (como el costo de ventas de 2019). Si no fuera así, habría que restar 2.081 adicionales en cada escenario; el Excel tiene un interruptor para ese caso (hoja Supuestos). El orden de las alternativas no cambia.",
  "En el mix se mantienen los gastos de I+D, ventas y administración de 2019, y el marketing es el %M de las ventas de alta gama. Es un supuesto conservador.",
  "Se usan costos variables constantes por kilo; no se usa la matriz de costo por volumen del Anexo 3, como pide el enunciado.",
  "Deuda constante y tasa implícita de 2019; impuestos del 40 % y sin dividendos. Las palancas de mejora se miden por separado y su suma es aproximada.",
  "El mix con una sola calidad de alta gama es una simplificación del enunciado. Combinar varias calidades podría mejorar algo el resultado, pero exige más complejidad comercial.",
]) c.push(bullet(s));

// Anexo
c.push(h1("Anexo. Resultado del mix por calidad y nivel de marketing"));
c.push(p("Resultado = (p – cv)·q + 1,75·(6.000.000 – q) – 3.262 – %M·p·q, en miles de euros. En verde, el máximo."));
let maxv = -Infinity;
for (const f of d.tabla_c) for (const v of f.slice(2)) if (v > maxv) maxv = v;
c.push(tabla(["%M", ...d.mix_grados], d.tabla_c.map((f) => ({
  c: [pct(f[0], 0), ...f.slice(2).map((v) => n0(v))],
})), [1560, 1300, 1300, 1300, 1300, 1300, 1300]));
// resaltar el máximo: se rehace la tabla con sombreado en la celda óptima
c.pop();
{
  const anchos = [1560, 1300, 1300, 1300, 1300, 1300, 1300];
  const cel = (t, w, opt = {}) => new TableCell({
    borders: bordes, width: { size: w, type: WidthType.DXA },
    shading: opt.head ? { fill: AZUL_OSC, type: ShadingType.CLEAR, color: "auto" }
      : opt.max ? { fill: "C6E0B4", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ alignment: opt.head ? AlignmentType.CENTER : AlignmentType.RIGHT,
      children: [new TextRun({ text: t, size: 17, bold: !!(opt.head || opt.max), color: opt.head ? "FFFFFF" : undefined })] })],
  });
  const rows = [new TableRow({ tableHeader: true, children: ["%M", ...d.mix_grados].map((t, i) => cel(t, anchos[i], { head: true })) })];
  for (const f of d.tabla_c) {
    rows.push(new TableRow({ children: [cel(pct(f[0], 0), anchos[0]), ...f.slice(2).map((v, i) => cel(n0(v), anchos[i + 1], { max: v === maxv }))] }));
  }
  c.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: anchos, rows }));
}

// ---------- documento ----------
const doc = new Document({
  creator: "Análisis de caso",
  title: "Cafés Monte Bianco – Plan de ganancias 2020",
  styles: {
    default: { document: { run: { font: FUENTE, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: FUENTE, color: AZUL_OSC },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FUENTE, color: "2F5496" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "viñetas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "numeros", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "Cafés Monte Bianco – Plan de ganancias 2020", size: 16, color: GRIS })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ["Página ", PageNumber.CURRENT], size: 16, color: GRIS })] })] }) },
    children: c,
  }],
});

Packer.toBuffer(doc).then((buf) => fs.writeFileSync(salida, buf));
