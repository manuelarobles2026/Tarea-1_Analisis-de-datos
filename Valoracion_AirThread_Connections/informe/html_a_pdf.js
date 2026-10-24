// Convierte el HTML del informe a PDF con el Chromium preinstalado (Playwright).
const { chromium } = require('playwright');
(async () => {
  const [, , html, pdf] = process.argv;
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + require('path').resolve(html), { waitUntil: 'load' });
  await page.pdf({ path: pdf, format: 'Letter', printBackground: true, preferCSSPageSize: true,
    displayHeaderFooter: true, headerTemplate: '<span></span>',
    footerTemplate: '<div style="font-size:7pt;width:100%;text-align:center;color:#777;font-family:Arial">AirThread Connections – valoración · página <span class="pageNumber"></span> de <span class="totalPages"></span></div>',
    margin: { top: '16mm', bottom: '16mm', left: '15mm', right: '15mm' } });
  await browser.close();
})();
