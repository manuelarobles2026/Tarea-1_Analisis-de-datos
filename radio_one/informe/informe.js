const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, BorderStyle, LevelFormat,
} = require("docx");

const FONT = "Arial";
const BODY = 21; // 10,5 pt
const SMALL = 17; // 8,5 pt
const TBL = 18; // 9 pt
const NAVY = "17365D";
const PAGE_W = 11906, MARGIN = 1134; // A4, 2 cm
const TEXT_W = PAGE_W - 2 * MARGIN; // 9638

// texto con **negritas**
function runs(text, opts = {}) {
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((t) =>
    t.startsWith("**")
      ? new TextRun({ text: t.slice(2, -2), bold: true, font: FONT, size: opts.size || BODY, color: opts.color })
      : new TextRun({ text: t, font: FONT, size: opts.size || BODY, color: opts.color, italics: opts.italics }));
}
const P = (text, o = {}) => new Paragraph({
  children: runs(text, o), alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 80, line: 252 }, keepNext: o.keepNext,
});
const H = (text) => new Paragraph({
  children: [new TextRun({ text, bold: true, font: FONT, size: 24, color: NAVY })],
  spacing: { before: 140, after: 60 }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 } },
});
const B = (text) => new Paragraph({
  children: runs(text), numbering: { reference: "vi", level: 0 },
  alignment: AlignmentType.JUSTIFIED, spacing: { after: 50, line: 252 },
});
const N = (text) => new Paragraph({
  children: runs(text), numbering: { reference: "num", level: 0 },
  alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, line: 252 },
});

const border = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0, cantSplit: true,
      children: r.map((c, j) => {
        const head = i === 0;
        const hl = !head && opts.highlightCol === j;
        const firstCol = j === 0;
        return new TableCell({
          width: { size: widths[j], type: WidthType.DXA }, borders,
          margins: { top: 30, bottom: 30, left: 70, right: 70 },
          shading: head ? { fill: NAVY, type: ShadingType.CLEAR, color: "auto" }
            : hl ? { fill: "E2EFDA", type: ShadingType.CLEAR, color: "auto" } : undefined,
          children: [new Paragraph({
            alignment: firstCol || opts.leftAll ? AlignmentType.LEFT : AlignmentType.CENTER,
            children: runs(c, { size: TBL, color: head ? "FFFFFF" : undefined }).map((x) => {
              if (head) x.root?.forEach?.(() => {});
              return x;
            }),
          })],
        });
      }),
    })),
  });
}
// encabezados en negrita
const hdr = (arr) => arr.map((t) => `**${t}**`);

const children = [
  new Paragraph({ children: [new TextRun({ text: "Radio One: duplicar la empresa sin pagar de más", bold: true, font: FONT, size: 32, color: NAVY })], spacing: { after: 40 } }),
  new Paragraph({ children: [new TextRun({ text: "Grupo 2: Daniela González ׀ Andrés Morales ׀ Manuela Robles ׀ Norberto Rodríguez", font: FONT, size: 18, color: "595959" })], spacing: { after: 100 } }),
  new Paragraph({
    children: runs("**Recomendamos ofertar por las 12 emisoras urbanas que Clear Channel debe vender, abriendo en US$1.244 millones (20 veces la caja que generan en un año) y sin pasar de US$1.337 millones.** Es el precio que hoy paga el mercado y queda por debajo de los US$1.343 millones que estas emisoras valen para Radio One."),
    alignment: AlignmentType.JUSTIFIED, spacing: { after: 100, line: 252 },
    shading: { fill: "E2EFDA", type: ShadingType.CLEAR, color: "auto" },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: "70AD47", space: 6 } },
  }),

  H("1. Situación y problema"),
  P("La fusión de Clear Channel con AMFM obliga a vender emisoras. Debemos decidir si compramos las 12 urbanas que salen a la venta y cuánto pagamos por ellas."),
  B("**Es una oportunidad que no se repite.** El regulador de las comunicaciones (la FCC) exige a Clear Channel vender unas 100 emisoras; 12 son urbanas y están en grandes mercados afroamericanos como Los Ángeles, Houston y Dallas. Con ellas llegaríamos a más hogares afroamericanos que cualquier otro medio y venderíamos publicidad nacional como una sola red. Si no compramos, lo hará un competidor."),
  B("**No son las emisoras que solemos comprar.** Siempre hemos comprado emisoras pequeñas y poco rentables para mejorarlas. Estas ya son grandes y rentables: dejan unos US$62 millones de caja al año, 1,7 veces lo que generó todo Radio One en 1999. Por eso hay menos espacio para mejorar y el precio es mucho mayor."),
  B("**El precio es el riesgo central.** Los compradores recientes pagaron entre 18 y 21,5 veces la caja anual de emisoras parecidas, y nuestra acción subió de 40 a 97 dólares solo por la expectativa de esta compra. Pagar de más borraría el valor que la operación puede crear."),

  H("2. Alternativas y análisis"),
  P("Valoramos las emisoras de tres formas: el dinero que generarán en el futuro traído a hoy, el precio al que cotizan empresas de radio parecidas y lo que se pagó en compras recientes. Con eso comparamos cuatro ofertas posibles."),
  P("Entre 2001 y 2004 las emisoras dejarían libres entre US$43 y 63 millones al año, después de impuestos, de US$1,2 millones anuales de inversión en equipos y del aumento de la cartera. En 2000 debemos aportar unos US$26 millones de capital de trabajo, porque el vendedor no lo incluye. Una operación más grande también sube los gastos corporativos, de US$0,9 a 4,5 millones al año. Además, la compra genera unos US$90 millones anuales de amortización deducible hasta 2015: no es una salida de caja, pero reduce los impuestos en unos US$31 millones al año y hoy vale unos US$216 millones."),
  P("A esos flujos les exigimos una rentabilidad mínima del 11,4% anual, la que pediría un inversionista por un negocio de radio con el bajo endeudamiento de Radio One. Para el valor después de 2004 probamos dos supuestos: que el negocio crezca 3,5% al año para siempre (US$883 millones) o que pueda venderse a 15 veces su utilidad operativa, menos de lo que paga hoy el mercado (US$1.343 millones). Damos más peso al segundo: en 2004 las emisoras aún crecerían casi 13% al año y frenarlas de golpe al 3,5% subestima su valor. Las empresas de radio en bolsa indican unos US$1.306 millones, y las compras de Cox e Infinity, entre US$1.144 y 1.337 millones.", { keepNext: true }),
  table([2700, 1580, 1580, 2178, 1600], [
    hdr(["US$ millones", "A. No comprar", "B. Ofertar 883", "C. Ofertar 1.244 (máx. 1.337)", "D. Pagar 1.866"]),
    ["Veces la caja anual de las emisoras", "–", "14", "20 a 21,5", "30"],
    ["Valor para Radio One menos el precio", "0", "+460", "+99 a +6", "−523"],
    ["Métodos que valen al menos el precio", "–", "5 de 5", "3 de 5", "0 de 5"],
    ["¿Ganamos la compra?", "No", "Difícil: Cox pagó más", "Probable", "Sí"],
    ["Dinero a conseguir además de la caja actual", "0", "620", "980 a 1.075", "1.600"],
    ["**¿Conviene?**", "**No**", "**No**", "**Sí**", "**No**"],
  ], { highlightCol: 3 }),
  P("", { after: 40 }),
  P("**A. No comprar: perdemos la escala.** No arriesgamos dinero, pero un competidor se queda con los únicos activos urbanos grandes que saldrán a la venta, y Radio One pierde la plataforma para crecer hacia la televisión por cable, la música e Internet."),
  P("**B. Ofertar US$883 millones: barato, pero perdemos la compra.** Pagaríamos 14 veces la caja, el valor de las emisoras si dejaran de crecer en 2005. Cox ya pagó 18,4 veces por emisoras parecidas, así que Clear Channel no aceptaría esta oferta."),
  P("**C. Ofertar US$1.244 millones con un máximo de US$1.337: precio de mercado con valor a favor.** Se sitúa entre lo que pagaron Cox e Infinity y deja unos US$99 millones de valor si cerramos en la apertura. A cambio, dependemos de que la caja pase de US$62 a 106 millones entre 2000 y 2004 y de recibir el ahorro de impuestos, casi la sexta parte del precio. Su principal riesgo es conseguir unos US$980 millones."),
  P("**D. Pagar hasta 30 veces la caja (US$1.866 millones): pagamos dos veces la misma expectativa.** Nuestra acción vale 30 veces la caja porque el mercado ya cuenta con esta compra; usar ese múltiplo como precio destruiría unos US$523 millones."),

  H("3. Recomendación"),
  P("Recomendamos la oferta C porque es la única que gana la compra sin destruir valor: frente a la oferta B sí compite con los precios del mercado, y frente a la D ahorra unos US$622 millones.", { keepNext: true }),
  table([2500, 3000, 4138], [
    hdr(["Riesgo", "Cuánto puede afectar", "Cómo se cubre"]),
    ["La caja crece menos de lo previsto", "Si el crecimiento se frena en 2005, el valor baja de US$1.343 a 883 millones", "Precio máximo de US$1.337 millones; ya mejoramos emisoras antes: en Washington la caja pasó de US$6,3 a 14,9 millones en tres años"],
    ["Financiamiento", "Unos US$980 millones por encima de la caja actual (US$263 millones)", "Emitir unos 10 millones de acciones a cerca de 97 dólares cubre todo sin subir la deuda"],
    ["El ahorro de impuestos no se concreta", "Hasta US$216 millones de valor", "Comprar los activos, no las sociedades, para poder amortizar el precio; condicionar la oferta a confirmarlo"],
    ["Integración en mercados nuevos", "Dallas convierte en caja solo el 21% de sus ingresos, frente al 59% del grupo", "Plan de 100 días y retención de gerentes; llevar Dallas al promedio suma unos US$2 millones al año"],
  ], { leftAll: true }),
  P("", { after: 40 }),
  N("**Aprobar** una oferta de US$1.244 millones a Clear Channel por las 12 emisoras antes del 30/04/2000, condicionada a la revisión detallada de cada emisora (due diligence)."),
  N("**Fijar** US$1.337 millones como precio máximo; si la puja lo supera, nos retiramos."),
  N("**Autorizar** la consecución de unos US$980 millones, dando prioridad a la emisión de acciones mientras la acción cotice cerca de 97 dólares."),
  N("**Estructurar** la operación como compra de activos para asegurar unos US$31 millones al año de ahorro de impuestos hasta 2015."),
  N("**Presentar** a la Junta, antes del cierre, un plan de integración de 100 días para Los Ángeles, Houston y Dallas."),

  new Paragraph({
    children: runs("**Supuestos técnicos.** Rentabilidad mínima del 11,4%: tasa sin riesgo del 6,28% (bono del Tesoro a 10 años), beta de 0,79 (promedio del sector reapalancado con Hamada), prima de mercado del 7%, costo del patrimonio del 11,8% y de la deuda del 9,68% (6,4% después de impuestos), con 8% de deuda a valor en libros y 92% de patrimonio a US$40 por acción antes del anuncio. Impuestos del 34%. El capital de trabajo usa la rotación histórica de cartera (26% de los ingresos) y proveedores (3% de los gastos). Valoramos solo las 12 emisoras de Clear Channel; las 9 de Charlotte, Augusta e Indianápolis se negocian aparte con otros vendedores.", { size: SMALL }),
    alignment: AlignmentType.JUSTIFIED, spacing: { before: 100, after: 40, line: 228 },
  }),
  new Paragraph({
    children: runs("**Referencias.** Ruback, R. y Fischer, P. (2000). Radio One, Inc. Harvard Business School, caso 204-S05. Modelo financiero del Grupo 2 en Excel.", { size: SMALL }),
    spacing: { after: 0, line: 228 },
  }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: BODY } } } },
  numbering: {
    config: [
      { reference: "vi", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 300 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    children,
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync(process.argv[2], b));
