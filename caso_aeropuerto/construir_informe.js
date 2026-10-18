// Genera el informe de 2 páginas para la Junta Directiva (cifras tomadas del modelo Excel).
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, LevelFormat,
} = require("docx");

const AZUL = "1F3864", VERDE = "E2EFDA", GRIS = "F2F2F2";
const F = "Arial", T = 21; // 10,5 pt

const run = (text, o = {}) => new TextRun({ text, font: F, size: T, ...o });
const p = (children, o = {}) => new Paragraph({
  children: typeof children === "string" ? [run(children)] : children,
  spacing: { after: 60, line: 264 }, alignment: AlignmentType.JUSTIFIED, ...o,
});
const h = (text) => new Paragraph({
  children: [run(text, { bold: true, size: 22, color: AZUL })],
  spacing: { before: 120, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 1 } },
});
const bullet = (children) => new Paragraph({
  children, numbering: { reference: "vinetas", level: 0 },
  spacing: { after: 40, line: 264 }, alignment: AlignmentType.JUSTIFIED,
});

const ANCHO = 10206; // 18 cm útiles
const cols = [2766, 2480, 2480, 2480];
const bordeFino = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const celda = (text, i, o = {}) => new TableCell({
  width: { size: cols[i], type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
  margins: { top: 30, bottom: 30, left: 80, right: 80 },
  borders: { top: bordeFino, bottom: bordeFino, left: bordeFino, right: bordeFino },
  children: [new Paragraph({
    alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
    children: [run(text, { size: 18, bold: !!o.bold, color: o.color })],
  })],
});
const filaTabla = (vals, o = {}) => new TableRow({
  children: vals.map((v, i) => celda(v, i, {
    bold: o.bold || i === 0 && o.boldFirst,
    fill: o.head ? AZUL : (i === 3 ? VERDE : (o.zebra ? GRIS : undefined)),
    color: o.head ? "FFFFFF" : undefined,
  })),
});

const tabla = new Table({
  width: { size: ANCHO, type: WidthType.DXA },
  columnWidths: cols,
  rows: [
    filaTabla(["Componente (CAPM)", "A. Postura del regulador", "B. Máximo de la empresa", "C. Propuesta recomendada"], { head: true, bold: true }),
    filaTabla(["Tasa libre de riesgo", "4,22% (bono 10 años, spot)", "4,84% (bono 30 años)", "4,77% (bono 10 años, prom. 2000–04)"], { boldFirst: true }),
    filaTabla(["Prima de mercado (ERP)", "4,8% (geométrica)", "7,2% (aritmética)", "6,0% (reguladores price cap)"], { boldFirst: true, zebra: true }),
    filaTabla(["Beta", "0,65 (comparables)", "0,94 (+20% calificación, +20% tamaño)", "0,78 (+20% calificación)"], { boldFirst: true }),
    filaTabla(["Prima regulatoria", "0%", "2%", "0% (ya está en la beta)"], { boldFirst: true, zebra: true }),
    filaTabla(["Riesgo país", "0 pb", "2.857 pb (Arg. 1996–2004)", "619 pb (normalizado)"], { boldFirst: true }),
    filaTabla(["WACC = Ke (US$)", "7,4%", "42,2%", "15,7%"], { bold: true, zebra: true }),
    filaTabla(["Costo del capital por pasajero", "$1,32", "$7,55", "$2,80"], { boldFirst: true }),
    filaTabla(["Costo del capital / ingresos 2004", "21%", "121%", "45%"], { boldFirst: true, zebra: true }),
  ],
});

const doc = new Document({
  styles: { default: { document: { run: { font: F, size: T } } } },
  numbering: { config: [{ reference: "vinetas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 284, hanging: 200 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 900, left: 1017, right: 1017 } } },
    children: [
      new Paragraph({ children: [run("Concesión del Aeropuerto de Neuquén: tasa de descuento que debemos proponer", { bold: true, size: 30, color: AZUL })], spacing: { after: 40 } }),
      new Paragraph({
        children: [run("Para: Junta Directiva – Aeropuertos del Neuquén S.A. / AA2000     De: Gerencia de Proyectos     Fecha: 23 de diciembre de 2004", { size: 16, color: "595959" })],
        spacing: { after: 100 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 4 } },
      }),
      new Paragraph({
        shading: { type: ShadingType.CLEAR, fill: VERDE, color: "auto" },
        spacing: { before: 60, after: 100 },
        border: { left: { style: BorderStyle.SINGLE, size: 24, color: "548235", space: 6 } },
        alignment: AlignmentType.JUSTIFIED,
        children: [
          run("Recomendación: ", { bold: true }),
          run("proponer al regulador un WACC de 15,7% en dólares (costo del capital propio, porque AdeN no tiene deuda financiera). Abrir la negociación en 17,7% y no aceptar menos de 12,1%. Con ese piso, el capital todavía cubre el riesgo de invertir en Argentina."),
        ],
      }),

      h("1. Situación"),
      p([run("La renegociación del contrato está cerrada en todo menos un punto: "), run("la tasa a la que se descuentan los flujos de fondos", { bold: true }), run(". De ella sale la tarifa que AdeN podrá cobrar durante la concesión. Los hechos que importan:")]),
      bullet([run("AdeN pierde dinero: ", { bold: true }), run("tuvo una pérdida operativa de $344 mil en 2004 y sobrevive con $4,4 millones de deudas con accionistas. Una tasa baja vuelve permanente esa situación.")]),
      bullet([run("Cada punto de WACC vale $57 mil por año ", { bold: true }), run("(1% × base de activos de $5,7 millones), es decir $0,18 por pasajero. El ingreso actual es de $6,23 por pasajero.")]),
      bullet([run("Ya hay acuerdo en el método: ", { bold: true }), run("WACC con CAPM y sin deuda (WACC = Ke). La discusión está en los insumos. La prima de mercado va de 2,9% a 8,1% según el período y el promedio que se use. La beta se toma de operadores comparables (0,55–0,75), porque AdeN no cotiza. "), run("La variable decisiva es el riesgo país", { bold: true }), run(": el EMBI+ de Argentina está en 4.668 pb por el default, y algunos reguladores (EPRE San Juan) directamente no lo incluyen.")]),

      h("2. Alternativas evaluadas"),
      p("Construimos tres combinaciones con los datos del caso (detalle y sensibilidades en el modelo Excel adjunto):", { spacing: { after: 80 } }),
      tabla,
      new Paragraph({ children: [run("Costo del capital = WACC × base de activos regulados ($5,73 millones). Pasajeros 2004: 320 mil. Fuente: modelo Excel, hojas \"Alternativas\" y \"Sensibilidad\".", { size: 14, italics: true, color: "595959" })], spacing: { before: 40, after: 80 } }),
      bullet([run("A. Postura del regulador (7,4%). ", { bold: true }), run("Es la que la Provincia aceptaría sin discutir, pero exige el retorno de una empresa de servicios en EE.UU. a una inversión en un país en default. Ningún accionista pondría capital nuevo a esa tasa, se mantendrían las pérdidas y "), run("quedaría un precedente peligroso para los 33 aeropuertos de AA2000", { bold: true }), run(".")]),
      bullet([run("B. Máximo de la empresa (42,2%). ", { bold: true }), run("Se puede sostener con números del caso, pero el costo del capital sería el 121% de los ingresos actuales, así que las tasas tendrían que más que duplicarse. Con el riesgo país del día la tasa llegaría a 56%. El regulador la rechazaría y pondría en riesgo un acuerdo que ya está casi cerrado.")]),
      bullet([run("C. Propuesta recomendada (15,7%). ", { bold: true }), run("Cada supuesto tiene un precedente regulatorio. Usa el riesgo país normalizado de 619 pb (CREG Colombia, 2003), casi igual al de Argentina antes de la crisis (659 pb), y no la cifra del default. Sube la beta 20%, como se hizo con la DAA de Dublín por una baja de calificación menor que la distancia entre A y la BBB+ de AA2000. "), run("No suma la prima regulatoria", { bold: true }), run(": los cuatro comparables ya operan con price cap, así que sumarla sería contar el mismo riesgo dos veces, y la contraparte lo notaría.")]),

      h("3. Recomendación"),
      p([run("Adoptar la alternativa C y negociar con un rango definido de antemano:", { bold: true })]),
      bullet([run("Abrir en 17,7% ", { bold: true }), run("(C más la prima regulatoria de 2% que es consenso de mercado). Es una concesión que podemos ceder sin tocar el fondo de la propuesta.")]),
      bullet([run("Cerrar en 15,7%. ", { bold: true }), run("Es un resultado robusto: con cualquier prima de mercado entre 4,1% y 7,2%, el WACC solo se mueve entre 14,2% y 16,6%. Deja un costo del capital de $2,80 por pasajero (45% de los ingresos de 2004), que se puede trasladar a la tarifa sin el salto de la alternativa B.")]),
      bullet([run("Piso de 12,1% ", { bold: true }), run("(prima de mercado de 4,8% y riesgo país igual al promedio de emergentes, 354 pb). Si el regulador insiste en un riesgo país de cero, conviene seguir negociando antes que firmar: 7,4% no compensa el riesgo de invertir en Argentina.")]),
      bullet([run("Dos condiciones para firmar. ", { bold: true }), run("(i) Los flujos deben estar en dólares, porque la tasa se calculó en esa moneda; si están en pesos, hay que sumarle el diferencial de inflación. (ii) Una cláusula que revise el riesgo país cada 5 años según el EMBI+. Así, si Argentina se normaliza, la tasa baja: le da al regulador una razón para aceptar hoy y protege a los usuarios.")]),

      new Paragraph({ children: [run("Referencias: ", { bold: true, size: 14 }), run("Preve, L. y Sarria Allende, V. (2009). Aeropuertos del Neuquén S.A. (A) “El costo de capital”, caso FZ-C-204-IA-1-s, IAE – Universidad Austral (Anexos 1–5). Damodaran, A. e Ibbotson & Associates, primas de mercado de EE.UU. 1926–2003. Hutson, E. y Kearney, C., Dublin Airport Authority’s Cost of Capital. CREG Colombia (2003).", { size: 14, color: "595959" })], spacing: { before: 120 }, alignment: AlignmentType.JUSTIFIED }),
    ],
  }],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync("caso_aeropuerto/Informe_JD_Costo_Capital_AdeN.docx", b);
  console.log("ok");
});
