const fs = require('fs');
const JavaScriptObfuscator = require('javascript-obfuscator');
const CleanCSS = require('clean-css');

const html = fs.readFileSync('/home/user/elevator-app/app-v2.html', 'utf-8');

// ─── COPYRIGHT HEADER ───
const COPYRIGHT = `/*
 * ══════════════════════════════════════════════════════════
 * AscensorTech - Gestión Integral de Servicios de Ascensores
 * Copyright © 2026 AscensorTech. Todos los derechos reservados.
 *
 * Este software está protegido por las leyes de propiedad
 * intelectual de la República Argentina (Ley 11.723) y
 * tratados internacionales de copyright.
 *
 * Queda prohibida la reproducción, distribución, modificación,
 * ingeniería inversa, descompilación o cualquier uso no
 * autorizado de este código fuente.
 *
 * Registro DNDA en trámite.
 * Contacto: info@ascensortech.com
 * ══════════════════════════════════════════════════════════
 */`;

// ─── EXTRACT CSS ───
const cssMatch = html.match(/<style>([\s\S]*?)<\/style>/);
const cssRaw = cssMatch ? cssMatch[1] : '';

// Minify CSS
const cssMin = new CleanCSS({ level: 2 }).minify(cssRaw).styles;

// ─── EXTRACT ALL SCRIPT BLOCKS ───
// Find ALL script blocks
const scriptBlocks = [];
const scriptRegex = /<script>([\s\S]*?)<\/script>/g;
let m;
while ((m = scriptRegex.exec(html)) !== null) {
  scriptBlocks.push(m[1]);
}
const jsRaw = scriptBlocks.join('\n;\n');

console.log(`Found ${scriptBlocks.length} script block(s), total ${jsRaw.length} chars of JS`);

// ─── OBFUSCATE JS ───
const obfResult = JavaScriptObfuscator.obfuscate(jsRaw, {
  // High protection
  compact: true,
  controlFlowFlattening: true,
  controlFlowFlatteningThreshold: 0.6,
  deadCodeInjection: true,
  deadCodeInjectionThreshold: 0.3,
  debugProtection: false,           // avoid breaking for legit debugging
  disableConsoleOutput: false,
  identifierNamesGenerator: 'hexadecimal',
  log: false,
  numbersToExpressions: true,
  renameGlobals: false,             // keep globals so HTML onclick/events work
  selfDefending: false,             // avoid breaking if formatted
  simplify: true,
  splitStrings: true,
  splitStringsChunkLength: 8,
  stringArray: true,
  stringArrayCallsTransform: true,
  stringArrayEncoding: ['rc4'],
  stringArrayIndexShift: true,
  stringArrayRotate: true,
  stringArrayShuffle: true,
  stringArrayWrappersCount: 2,
  stringArrayWrappersChainedCalls: true,
  stringArrayWrappersParametersMaxCount: 4,
  stringArrayWrappersType: 'function',
  stringArrayThreshold: 0.75,
  transformObjectKeys: true,
  unicodeEscapeSequence: false,
});

const jsObf = obfResult.getObfuscatedCode();
console.log(`Obfuscated JS: ${jsObf.length} chars`);

// ─── EXTRACT THE HTML BODY PORTION (between </style> and <script>) ───
const afterStyle = html.indexOf('</style>') + '</style>'.length;
const firstScript = html.indexOf('<script>');
const htmlBody = html.substring(afterStyle, firstScript).trim();

// ─── WATERMARKS: hidden identifiers embedded in the code ───
const WATERMARK = `var _wm_ascensortech="${Buffer.from('AscensorTech-Copyright-2026-Registro-DNDA-' + Date.now()).toString('base64')}";`;

// ─── REASSEMBLE ───
const output = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AscensorTech</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap">
${COPYRIGHT}
<style>${cssMin}</style>
</head>
<body>
${htmlBody}
<script>
${COPYRIGHT}
${WATERMARK}
${jsObf}
</script>
</body>
</html>`;

const outPath = '/home/user/elevator-app/app-v2-protected.html';
fs.writeFileSync(outPath, output, 'utf-8');

console.log(`\n✓ Output: ${outPath}`);
console.log(`  Original: ${html.length} chars`);
console.log(`  Protected: ${output.length} chars`);
console.log(`  CSS: ${cssRaw.length} → ${cssMin.length} (minified)`);
console.log(`  JS: ${jsRaw.length} → ${jsObf.length} (obfuscated)`);
