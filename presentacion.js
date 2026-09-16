const pptxgen = require('pptxgenjs');
const pres = new pptxgen();

// ─── CONFIG ───
pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
pres.author = 'AscensorTech';
pres.title = 'AscensorTech — Gestión Integral de Ascensores';

// ─── PALETTE ───
const C = {
  navy:    '0B2447',
  blue:    '19508C',
  teal:    '1B7A8C',
  cyan:    '00C9A7',
  light:   'EFF6FB',
  white:   'FFFFFF',
  offWhite:'F5F9FC',
  dark:    '0A1628',
  grey:    '64748B',
  greyLt:  'CBD5E1',
  accent:  'FF6B35',
  success: '10B981',
  warning: 'F59E0B',
  danger:  'EF4444',
};

const FONT_TITLE = 'Cambria';
const FONT_BODY = 'Calibri';

// ─── HELPERS ───
function addBg(slide, color) {
  slide.background = { color };
}
function addPageNum(slide, light) {
  slide.addText([{ text: '', options: { field: 'slidenum' } }], {
    x: 12.4, y: 7.05, w: 0.6, h: 0.3, fontSize: 9,
    fontFace: FONT_BODY, color: light ? '8899AA' : C.grey,
    align: 'right', isTextBox: true,
  });
}
function addFooter(slide, light) {
  slide.addText('AscensorTech — Gestión Integral de Ascensores', {
    x: 0.6, y: 7.05, w: 5, h: 0.3, fontSize: 9,
    fontFace: FONT_BODY, color: light ? '6688AA' : C.greyLt,
    isTextBox: true, margin: 0,
  });
  addPageNum(slide, light);
}

function iconCircle(slide, x, y, size, bgColor, iconText, iconColor) {
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: size, h: size, fill: { color: bgColor },
  });
  slide.addText(iconText, {
    x, y, w: size, h: size, fontSize: Math.round(size * 28),
    fontFace: FONT_BODY, color: iconColor || C.white,
    align: 'center', valign: 'middle', bold: true, isTextBox: true, margin: 0,
  });
}

// ══════════════════════════════════════════════════
// SLIDE 1: COVER
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.dark);

  // Gradient-like layered shapes
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 7.5,
    fill: { color: C.navy },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 7, y: -2.5, w: 10, h: 10,
    fill: { color: C.blue }, transparency: 70,
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 9, y: 2, w: 7, h: 7,
    fill: { color: C.teal }, transparency: 75,
  });

  // Title block
  s.addText('ASCENSOR', {
    x: 0.8, y: 1.6, w: 8, h: 0.9, fontSize: 52,
    fontFace: FONT_TITLE, color: C.white, bold: true,
    charSpacing: 6, isTextBox: true, margin: 0,
  });
  s.addText('TECH', {
    x: 0.8, y: 2.4, w: 8, h: 0.9, fontSize: 52,
    fontFace: FONT_TITLE, color: C.cyan, bold: true,
    charSpacing: 6, isTextBox: true, margin: 0,
  });

  s.addText('Gestión Integral de Servicios de Ascensores', {
    x: 0.8, y: 3.6, w: 7, h: 0.5, fontSize: 22,
    fontFace: FONT_BODY, color: C.greyLt, isTextBox: true, margin: 0,
  });

  s.addText('Plataforma unificada para la gestión completa de\nmantenimiento, reparaciones y relación con clientes', {
    x: 0.8, y: 4.4, w: 6.5, h: 0.8, fontSize: 14,
    fontFace: FONT_BODY, color: '8899BB', lineSpacingMultiple: 1.4,
    isTextBox: true, margin: 0,
  });

  // Stats bar at bottom
  const stats = [
    { n: '100%', label: 'Digital' },
    { n: '4', label: 'Roles integrados' },
    { n: '24/7', label: 'Trazabilidad' },
    { n: '0', label: 'Puntos ciegos' },
  ];
  const barY = 5.8;
  s.addShape(pres.ShapeType.rect, {
    x: 0.8, y: barY - 0.15, w: 11.6, h: 1.2,
    fill: { color: C.dark }, transparency: 40,
    rectRadius: 0.12, shapeName: 'ROUNDED_RECTANGLE',
  });
  stats.forEach((st, i) => {
    const sx = 1.2 + i * 2.85;
    s.addText(st.n, {
      x: sx, y: barY, w: 2.2, h: 0.5, fontSize: 28,
      fontFace: FONT_TITLE, color: C.cyan, bold: true,
      align: 'center', isTextBox: true, margin: 0,
    });
    s.addText(st.label, {
      x: sx, y: barY + 0.5, w: 2.2, h: 0.3, fontSize: 11,
      fontFace: FONT_BODY, color: '7788AA', align: 'center',
      isTextBox: true, margin: 0,
    });
  });

  s.addNotes('Slide de apertura. Presentar AscensorTech como plataforma integral.');
}

// ══════════════════════════════════════════════════
// SLIDE 2: EL PROBLEMA
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.white);

  s.addText('El problema actual', {
    x: 0.8, y: 0.5, w: 6, h: 0.7, fontSize: 36,
    fontFace: FONT_TITLE, color: C.navy, bold: true,
    isTextBox: true, margin: 0,
  });
  s.addText('La gestión manual genera pérdidas, demoras y falta de control', {
    x: 0.8, y: 1.15, w: 8, h: 0.4, fontSize: 14,
    fontFace: FONT_BODY, color: C.grey, isTextBox: true, margin: 0,
  });

  const problems = [
    { icon: '!', color: C.danger, title: 'Procesos desconectados', desc: 'Cada área trabaja aislada: el técnico no sabe si hay stock, compras no sabe qué aprobó el jefe, el cliente no sabe el estado.' },
    { icon: '?', color: C.warning, title: 'Sin visibilidad en tiempo real', desc: 'Nadie tiene un panorama completo de cada intervención. La información viaja por WhatsApp o papel.' },
    { icon: '×', color: C.danger, title: 'Presupuestos manuales y lentos', desc: 'Generar un presupuesto requiere llamadas, consultar stock a mano y esperar aprobaciones verbales.' },
    { icon: '$', color: C.accent, title: 'Pérdida de ingresos', desc: 'Demoras en facturación, morosidad no detectada a tiempo, repuestos comprados de más o de menos.' },
    { icon: '⏱', color: C.blue, title: 'Demoras en reparaciones', desc: 'Sin un flujo claro, las solicitudes se pierden o se demoran semanas por falta de seguimiento.' },
    { icon: '✗', color: C.grey, title: 'Clientes insatisfechos', desc: 'El cliente no puede ver el avance de su pedido ni aprobar presupuestos de forma ágil.' },
  ];

  problems.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const px = 0.8 + col * 6.1;
    const py = 1.9 + row * 1.7;

    s.addShape(pres.ShapeType.roundRect, {
      x: px, y: py, w: 5.8, h: 1.45,
      fill: { color: C.offWhite },
      rectRadius: 0.08,
    });
    iconCircle(s, px + 0.25, py + 0.35, 0.7, p.color, p.icon, C.white);
    s.addText(p.title, {
      x: px + 1.15, y: py + 0.2, w: 4.3, h: 0.35, fontSize: 15,
      fontFace: FONT_BODY, color: C.navy, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(p.desc, {
      x: px + 1.15, y: py + 0.55, w: 4.3, h: 0.7, fontSize: 11,
      fontFace: FONT_BODY, color: C.grey, lineSpacingMultiple: 1.3,
      isTextBox: true, margin: 0,
    });
  });

  addFooter(s, false);
  s.addNotes('Describir los pain points que sufren las empresas de ascensores con gestión manual.');
}

// ══════════════════════════════════════════════════
// SLIDE 3: LA SOLUCIÓN
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.white);

  // Left half — text
  s.addText('La solución', {
    x: 0.8, y: 0.5, w: 5.5, h: 0.7, fontSize: 36,
    fontFace: FONT_TITLE, color: C.navy, bold: true,
    isTextBox: true, margin: 0,
  });
  s.addText('Una plataforma que conecta a todos los actores\nen un flujo continuo y transparente', {
    x: 0.8, y: 1.15, w: 5.5, h: 0.6, fontSize: 14,
    fontFace: FONT_BODY, color: C.grey, lineSpacingMultiple: 1.4,
    isTextBox: true, margin: 0,
  });

  const features = [
    { title: 'Gestión de solicitudes', desc: 'Recepción, clasificación y asignación automática' },
    { title: 'Informes técnicos digitales', desc: 'Con fotos, repuestos y cálculo de presupuesto' },
    { title: 'Flujo de aprobaciones', desc: 'Jefe técnico → depósito/compras → cliente' },
    { title: 'Portal de clientes', desc: 'Aprobación online y seguimiento en tiempo real' },
    { title: 'Control de stock', desc: 'Inventario con alertas y movimientos automáticos' },
    { title: 'Facturación y cobros', desc: 'Generación automática, control de morosidad' },
  ];

  features.forEach((f, i) => {
    const fy = 2.0 + i * 0.82;
    s.addShape(pres.ShapeType.ellipse, {
      x: 1.0, y: fy + 0.08, w: 0.32, h: 0.32,
      fill: { color: C.cyan },
    });
    s.addText(String(i + 1), {
      x: 1.0, y: fy + 0.08, w: 0.32, h: 0.32, fontSize: 12,
      fontFace: FONT_BODY, color: C.white, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
    s.addText(f.title, {
      x: 1.55, y: fy, w: 4.5, h: 0.3, fontSize: 14,
      fontFace: FONT_BODY, color: C.navy, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(f.desc, {
      x: 1.55, y: fy + 0.3, w: 4.5, h: 0.3, fontSize: 11,
      fontFace: FONT_BODY, color: C.grey,
      isTextBox: true, margin: 0,
    });
  });

  // Right half — big visual block simulating app
  s.addShape(pres.ShapeType.roundRect, {
    x: 6.8, y: 0.5, w: 5.8, h: 6.5,
    fill: { color: C.navy },
    rectRadius: 0.15,
  });
  // Simulated sidebar
  s.addShape(pres.ShapeType.rect, {
    x: 6.95, y: 0.7, w: 1.5, h: 6.1,
    fill: { color: C.dark },
  });
  const sideItems = ['Dashboard', 'Recepción', 'Informes', 'Depósito', 'Pedidos', 'Presupuestos', 'Órdenes', 'Facturación'];
  sideItems.forEach((item, i) => {
    const active = i === 1;
    if (active) {
      s.addShape(pres.ShapeType.rect, {
        x: 7.05, y: 1.1 + i * 0.6, w: 1.3, h: 0.45,
        fill: { color: C.blue },
        rectRadius: 0.06, shapeName: 'ROUNDED_RECTANGLE',
      });
    }
    s.addText(item, {
      x: 7.05, y: 1.1 + i * 0.6, w: 1.3, h: 0.45, fontSize: 8,
      fontFace: FONT_BODY, color: active ? C.white : '6688AA',
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
  });

  // Simulated content area with KPI cards
  const kpis = [
    { val: '9', lbl: 'Edificios', bg: C.teal },
    { val: '15', lbl: 'Ascensores', bg: C.blue },
    { val: '5', lbl: 'Abiertas', bg: C.warning },
  ];
  kpis.forEach((k, i) => {
    const kx = 8.7 + i * 1.3;
    s.addShape(pres.ShapeType.roundRect, {
      x: kx, y: 1.0, w: 1.15, h: 0.9,
      fill: { color: k.bg }, transparency: 20,
      rectRadius: 0.08,
    });
    s.addText(k.val, {
      x: kx, y: 1.0, w: 1.15, h: 0.55, fontSize: 20,
      fontFace: FONT_TITLE, color: C.white, bold: true,
      align: 'center', valign: 'bottom', isTextBox: true, margin: 0,
    });
    s.addText(k.lbl, {
      x: kx, y: 1.5, w: 1.15, h: 0.35, fontSize: 8,
      fontFace: FONT_BODY, color: 'AABBCC',
      align: 'center', isTextBox: true, margin: 0,
    });
  });

  // Simulated table rows
  for (let i = 0; i < 6; i++) {
    s.addShape(pres.ShapeType.rect, {
      x: 8.6, y: 2.2 + i * 0.65, w: 3.85, h: 0.5,
      fill: { color: i % 2 === 0 ? '0D1F3C' : C.dark },
    });
    s.addText(`REQ-000${i + 9}   Edificio ${['Torre Palermo', 'Callao Res.', 'Lavalle Office', 'Palermo Suites', 'Reconquista', 'Belgrano'][i]}`, {
      x: 8.7, y: 2.2 + i * 0.65, w: 3.6, h: 0.5, fontSize: 7,
      fontFace: FONT_BODY, color: '8899BB',
      valign: 'middle', isTextBox: true, margin: 0,
    });
  }

  // App label
  s.addText('Vista real de la aplicación', {
    x: 8.6, y: 6.4, w: 3.9, h: 0.3, fontSize: 9,
    fontFace: FONT_BODY, color: '5566AA', italic: true,
    align: 'center', isTextBox: true, margin: 0,
  });

  addFooter(s, false);
  s.addNotes('Mostrar las 6 áreas clave y la vista de la app.');
}

// ══════════════════════════════════════════════════
// SLIDE 4: FLUJO INTEGRADO — El diagrama de interrelación
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.navy);

  s.addText('Flujo integrado entre todos los actores', {
    x: 0.8, y: 0.35, w: 11, h: 0.6, fontSize: 32,
    fontFace: FONT_TITLE, color: C.white, bold: true,
    isTextBox: true, margin: 0,
  });
  s.addText('Cada paso genera notificaciones automáticas — nadie queda fuera del circuito', {
    x: 0.8, y: 0.9, w: 11, h: 0.35, fontSize: 13,
    fontFace: FONT_BODY, color: '8899CC', isTextBox: true, margin: 0,
  });

  // 4 Actor nodes — compact cards to avoid overlap
  const actors = [
    { label: 'CLIENTE', sub: 'Solicita reparación\nAprueba presupuesto', x: 0.6, y: 1.5, color: C.accent },
    { label: 'TÉCNICO', sub: 'Diagnostica\nSolicita repuestos', x: 4.4, y: 1.5, color: C.teal },
    { label: 'JEFE TÉCNICO', sub: 'Aprueba pedidos\nSupervisa todo', x: 8.2, y: 1.5, color: C.blue },
    { label: 'DEPÓSITO / COMPRAS', sub: 'Verifica stock\nGestiona compras', x: 0.6, y: 4.2, color: C.success },
  ];

  actors.forEach(a => {
    s.addShape(pres.ShapeType.roundRect, {
      x: a.x, y: a.y, w: 3.3, h: 1.9,
      fill: { color: a.color }, transparency: 80,
      line: { color: a.color, width: 2 },
      rectRadius: 0.15,
    });
    s.addShape(pres.ShapeType.ellipse, {
      x: a.x + 0.2, y: a.y + 0.3, w: 0.8, h: 0.8,
      fill: { color: a.color },
    });
    s.addText(a.label.charAt(0), {
      x: a.x + 0.2, y: a.y + 0.3, w: 0.8, h: 0.8, fontSize: 24,
      fontFace: FONT_TITLE, color: C.white, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
    s.addText(a.label, {
      x: a.x + 1.15, y: a.y + 0.2, w: 2.0, h: 0.4, fontSize: 12,
      fontFace: FONT_BODY, color: C.white, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(a.sub, {
      x: a.x + 1.15, y: a.y + 0.6, w: 2.0, h: 0.55, fontSize: 10,
      fontFace: FONT_BODY, color: 'AABBDD',
      lineSpacingMultiple: 1.3, isTextBox: true, margin: 0,
    });
  });

  // Flow arrows (using shapes + text labels)
  // Arrow 1: Cliente → Técnico
  s.addShape(pres.ShapeType.line, {
    x: 3.9, y: 2.4, w: 0.5, h: 0,
    line: { color: C.cyan, width: 2.5, dashType: 'solid', endArrowType: 'arrow' },
  });
  s.addText('1. Solicitud', {
    x: 3.5, y: 2.0, w: 1.3, h: 0.3, fontSize: 8,
    fontFace: FONT_BODY, color: C.cyan, bold: true,
    align: 'center', isTextBox: true, margin: 0,
  });

  // Arrow 2: Técnico → Jefe
  s.addShape(pres.ShapeType.line, {
    x: 7.7, y: 2.4, w: 0.5, h: 0,
    line: { color: C.cyan, width: 2.5, dashType: 'solid', endArrowType: 'arrow' },
  });
  s.addText('2. Pide\nrepuestos', {
    x: 7.3, y: 1.9, w: 1.3, h: 0.5, fontSize: 8,
    fontFace: FONT_BODY, color: C.cyan, bold: true,
    align: 'center', lineSpacingMultiple: 1.2, isTextBox: true, margin: 0,
  });

  // Arrow 3: Jefe → Depósito (diagonal down-left)
  s.addShape(pres.ShapeType.line, {
    x: 3.9, y: 3.4, w: 4.3, h: 1.2,
    line: { color: C.cyan, width: 2.5, dashType: 'solid', endArrowType: 'arrow' },
    flipH: true,
  });
  s.addText('3. Aprueba y chequea stock', {
    x: 4.3, y: 3.55, w: 3.5, h: 0.3, fontSize: 8,
    fontFace: FONT_BODY, color: C.cyan, bold: true,
    align: 'center', isTextBox: true, margin: 0,
  });

  // Arrow 4: Jefe → Cliente (presupuesto bidirectional)
  s.addShape(pres.ShapeType.line, {
    x: 2.2, y: 3.6, w: 6.4, h: 0,
    line: { color: C.warning, width: 2, dashType: 'dash', endArrowType: 'arrow', beginArrowType: 'arrow' },
  });
  s.addText('4. Presupuesto auto-generado → Cliente aprueba → Jefe notificado', {
    x: 2.2, y: 3.8, w: 6.4, h: 0.3, fontSize: 9,
    fontFace: FONT_BODY, color: C.warning, bold: true,
    align: 'center', isTextBox: true, margin: 0,
  });

  // Arrow 5: Depósito → Técnico (to the right)
  s.addShape(pres.ShapeType.line, {
    x: 3.9, y: 5.1, w: 0.95, h: -1.5,
    line: { color: C.success, width: 2, dashType: 'dash', endArrowType: 'arrow' },
  });
  s.addText('5. Entrega\nrepuestos', {
    x: 4.3, y: 4.5, w: 1.5, h: 0.5, fontSize: 8,
    fontFace: FONT_BODY, color: C.success, bold: true,
    lineSpacingMultiple: 1.2, isTextBox: true, margin: 0,
  });

  // Central label — RETROALIMENTACIÓN
  s.addShape(pres.ShapeType.roundRect, {
    x: 4.6, y: 5.5, w: 7.6, h: 0.9,
    fill: { color: C.dark }, transparency: 30,
    rectRadius: 0.1,
  });
  s.addText('RETROALIMENTACIÓN CONTINUA', {
    x: 4.6, y: 5.52, w: 7.6, h: 0.4, fontSize: 15,
    fontFace: FONT_TITLE, color: C.cyan, bold: true,
    align: 'center', isTextBox: true, margin: 0,
  });
  s.addText('Cada actor ve en tiempo real el estado de la intervención a través del Timeline integrado', {
    x: 4.6, y: 5.92, w: 7.6, h: 0.35, fontSize: 10,
    fontFace: FONT_BODY, color: '99AACC',
    align: 'center', isTextBox: true, margin: 0,
  });

  // Bottom note bar
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 6.6, w: 12.1, h: 0.5,
    fill: { color: C.cyan }, transparency: 85,
    rectRadius: 0.06,
  });
  s.addText('Alertas automáticas  ·  Badges en navegación  ·  Timeline de intervención  ·  Aprobaciones online  ·  Trazabilidad completa', {
    x: 0.8, y: 6.6, w: 11.7, h: 0.5, fontSize: 11,
    fontFace: FONT_BODY, color: C.cyan, bold: true,
    align: 'center', valign: 'middle', isTextBox: true, margin: 0,
  });

  addPageNum(s, true);
  s.addNotes('El slide más importante: mostrar cómo cada actor está conectado en un circuito cerrado.');
}

// ══════════════════════════════════════════════════
// SLIDE 5: LOS 4 ROLES
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.white);

  s.addText('Cada rol ve exactamente lo que necesita', {
    x: 0.8, y: 0.45, w: 10, h: 0.7, fontSize: 32,
    fontFace: FONT_TITLE, color: C.navy, bold: true,
    isTextBox: true, margin: 0,
  });
  s.addText('Acceso controlado por permisos — sin distracciones, sin riesgo', {
    x: 0.8, y: 1.05, w: 10, h: 0.35, fontSize: 14,
    fontFace: FONT_BODY, color: C.grey, isTextBox: true, margin: 0,
  });

  const roles = [
    {
      title: 'Administrador',
      color: C.navy,
      items: ['Dashboard completo', 'Gestión de personal', 'Edificios y contratos', 'Facturación y cobros', 'Índices de actualización', 'Control total'],
    },
    {
      title: 'Jefe Técnico',
      color: C.blue,
      items: ['Aprobación de pedidos', 'Supervisión de informes', 'Alertas de aprobación', 'Órdenes de trabajo', 'Vista de stock', 'Reportes estadísticos'],
    },
    {
      title: 'Técnico',
      color: C.teal,
      items: ['Mis solicitudes', 'Carga de informes', 'Pedido de repuestos', 'Subir fotos', 'Consulta de historial', 'Turnos y guardias'],
    },
    {
      title: 'Cliente',
      color: C.accent,
      items: ['Estado de ascensores', 'Solicitar reparación', 'Aprobar presupuestos', 'Ver facturas', 'Timeline de avance', 'Subir fotos'],
    },
  ];

  roles.forEach((r, i) => {
    const rx = 0.6 + i * 3.1;
    const ry = 1.7;

    // Card
    s.addShape(pres.ShapeType.roundRect, {
      x: rx, y: ry, w: 2.9, h: 5.3,
      fill: { color: C.offWhite },
      rectRadius: 0.12,
    });

    // Header
    s.addShape(pres.ShapeType.roundRect, {
      x: rx, y: ry, w: 2.9, h: 1.3,
      fill: { color: r.color },
      rectRadius: 0.12,
    });
    // Fix bottom corners of header
    s.addShape(pres.ShapeType.rect, {
      x: rx, y: ry + 0.9, w: 2.9, h: 0.4,
      fill: { color: r.color },
    });

    s.addText(r.title.charAt(0), {
      x: rx + 0.9, y: ry + 0.1, w: 1.1, h: 0.7, fontSize: 30,
      fontFace: FONT_TITLE, color: C.white, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
    s.addText(r.title, {
      x: rx + 0.1, y: ry + 0.8, w: 2.7, h: 0.4, fontSize: 14,
      fontFace: FONT_BODY, color: C.white, bold: true,
      align: 'center', isTextBox: true, margin: 0,
    });

    // Items
    r.items.forEach((item, j) => {
      const iy = ry + 1.55 + j * 0.6;
      s.addShape(pres.ShapeType.ellipse, {
        x: rx + 0.25, y: iy + 0.06, w: 0.22, h: 0.22,
        fill: { color: r.color },
      });
      s.addText(item, {
        x: rx + 0.6, y: iy, w: 2.1, h: 0.35, fontSize: 11,
        fontFace: FONT_BODY, color: C.navy,
        valign: 'middle', isTextBox: true, margin: 0,
      });
    });
  });

  addFooter(s, false);
  s.addNotes('Explicar que cada persona ve solo su área. Seguridad y eficiencia.');
}

// ══════════════════════════════════════════════════
// SLIDE 6: EL CICLO COMPLETO — Timeline vertical
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.white);

  s.addText('El ciclo completo de una intervención', {
    x: 0.8, y: 0.35, w: 11, h: 0.6, fontSize: 32,
    fontFace: FONT_TITLE, color: C.navy, bold: true,
    isTextBox: true, margin: 0,
  });
  s.addText('De la solicitud del cliente al cierre — todo trazado', {
    x: 0.8, y: 0.9, w: 11, h: 0.3, fontSize: 13,
    fontFace: FONT_BODY, color: C.grey, isTextBox: true, margin: 0,
  });

  const timeline = [
    { step: '1', label: 'Cliente solicita', who: 'Cliente', desc: 'Crea solicitud con fotos y descripción', color: C.accent },
    { step: '2', label: 'Recepción y triaje', who: 'Admin/JT', desc: 'Clasificación por urgencia, asignación de técnico', color: C.navy },
    { step: '3', label: 'Diagnóstico técnico', who: 'Técnico', desc: 'Informe con fotos, repuestos necesarios y horas', color: C.teal },
    { step: '4', label: 'Pedido de repuestos', who: 'Técnico', desc: 'Se genera automáticamente al guardar el informe', color: C.teal },
    { step: '5', label: 'Aprobación del jefe', who: 'Jefe Técnico', desc: 'Aprueba, verifica stock → depósito o compras', color: C.blue },
    { step: '6', label: 'Presupuesto al cliente', who: 'Sistema', desc: 'Se genera y envía automáticamente al aprobar', color: C.warning },
    { step: '7', label: 'Aprobación del cliente', who: 'Cliente', desc: 'Aprueba online → el jefe recibe la notificación', color: C.accent },
    { step: '8', label: 'OT + Ejecución + Cierre', who: 'Admin/JT', desc: 'Orden de trabajo, factura automática, completado', color: C.success },
  ];

  // Two columns of 4
  timeline.forEach((t, i) => {
    const col = i < 4 ? 0 : 1;
    const row = i % 4;
    const tx = 0.6 + col * 6.5;
    const ty = 1.5 + row * 1.42;

    // Step circle
    s.addShape(pres.ShapeType.ellipse, {
      x: tx, y: ty + 0.05, w: 0.65, h: 0.65,
      fill: { color: t.color },
    });
    s.addText(t.step, {
      x: tx, y: ty + 0.05, w: 0.65, h: 0.65, fontSize: 20,
      fontFace: FONT_TITLE, color: C.white, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });

    // Connector line (not on last of each column)
    if (row < 3) {
      s.addShape(pres.ShapeType.line, {
        x: tx + 0.325, y: ty + 0.7, w: 0, h: 0.72,
        line: { color: C.greyLt, width: 2, dashType: 'dash' },
      });
    }

    // Text
    s.addText(t.label, {
      x: tx + 0.85, y: ty, w: 4.8, h: 0.35, fontSize: 15,
      fontFace: FONT_BODY, color: C.navy, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(t.desc, {
      x: tx + 0.85, y: ty + 0.32, w: 4.8, h: 0.3, fontSize: 11,
      fontFace: FONT_BODY, color: C.grey,
      isTextBox: true, margin: 0,
    });
    // WHO badge
    s.addShape(pres.ShapeType.roundRect, {
      x: tx + 0.85, y: ty + 0.62, w: 1.0, h: 0.28,
      fill: { color: t.color }, transparency: 85,
      rectRadius: 0.04,
    });
    s.addText(t.who, {
      x: tx + 0.85, y: ty + 0.62, w: 1.0, h: 0.28, fontSize: 8,
      fontFace: FONT_BODY, color: t.color, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
  });

  // Connector between columns (arrow from step 4 to 5)
  s.addShape(pres.ShapeType.line, {
    x: 5.7, y: 5.9, w: 1.4, h: -4.4,
    line: { color: C.cyan, width: 2.5, dashType: 'dash', endArrowType: 'arrow' },
    flipH: false,
  });

  addFooter(s, false);
  s.addNotes('Recorrer cada paso del ciclo. Enfatizar la automatización en pasos 5-6-7.');
}

// ══════════════════════════════════════════════════
// SLIDE 7: FUNCIONALIDADES CLAVE
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.navy);

  s.addText('Funcionalidades que marcan la diferencia', {
    x: 0.8, y: 0.4, w: 11, h: 0.6, fontSize: 32,
    fontFace: FONT_TITLE, color: C.white, bold: true,
    isTextBox: true, margin: 0,
  });

  const feats = [
    { icon: '◉', title: 'Timeline de intervención', desc: 'Cada solicitud muestra un timeline vertical con todos los eventos, visible para todos los actores involucrados.', color: C.cyan },
    { icon: '⚡', title: 'Aprobaciones automáticas', desc: 'Cuando el jefe aprueba repuestos, se verifica stock y se genera presupuesto sin pasos manuales.', color: C.warning },
    { icon: '📊', title: 'Reportes y estadísticas', desc: 'Dashboard con métricas diarias, semanales y mensuales. Distribución por tipo, tiempos de respuesta.', color: C.accent },
    { icon: '🖨', title: 'Impresión de documentos', desc: 'Presupuestos, informes técnicos, órdenes de trabajo y remitos listos para imprimir con formato profesional.', color: C.blue },
    { icon: '📸', title: 'Fotos integradas', desc: 'Técnicos y clientes adjuntan fotos en cada paso. Las fotos viajan con el informe y el presupuesto.', color: C.teal },
    { icon: '🔔', title: 'Alertas inteligentes', desc: 'Badges en la navegación, alertas en el dashboard. Stock bajo, morosidad, inspecciones por vencer.', color: C.danger },
    { icon: '📋', title: 'Control de stock', desc: 'Inventario con movimientos automáticos al entregar repuestos. Alertas de stock mínimo.', color: C.success },
    { icon: '💰', title: 'Facturación integrada', desc: 'Al completar una OT se genera factura automáticamente. Control de cobros y morosidad.', color: C.accent },
  ];

  feats.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const fx = 0.6 + col * 6.3;
    const fy = 1.3 + row * 1.4;

    s.addShape(pres.ShapeType.roundRect, {
      x: fx, y: fy, w: 6.0, h: 1.2,
      fill: { color: C.dark }, transparency: 30,
      rectRadius: 0.1,
    });
    iconCircle(s, fx + 0.2, fy + 0.25, 0.7, f.color, f.icon.charAt(0), C.white);
    s.addText(f.title, {
      x: fx + 1.1, y: fy + 0.15, w: 4.6, h: 0.35, fontSize: 14,
      fontFace: FONT_BODY, color: C.white, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(f.desc, {
      x: fx + 1.1, y: fy + 0.5, w: 4.6, h: 0.55, fontSize: 10,
      fontFace: FONT_BODY, color: '99AACC', lineSpacingMultiple: 1.3,
      isTextBox: true, margin: 0,
    });
  });

  addPageNum(s, true);
  s.addNotes('Recorrer las funcionalidades diferenciadores.');
}

// ══════════════════════════════════════════════════
// SLIDE 8: BENEFICIOS CUANTITATIVOS
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.white);

  s.addText('Beneficios concretos', {
    x: 0.8, y: 0.45, w: 11, h: 0.7, fontSize: 36,
    fontFace: FONT_TITLE, color: C.navy, bold: true,
    isTextBox: true, margin: 0,
  });

  // Big stat callouts - 3 across top
  const bigStats = [
    { value: '70%', label: 'menos tiempo\nen gestión de pedidos', color: C.teal },
    { value: '0', label: 'solicitudes\nperdidas', color: C.success },
    { value: '100%', label: 'trazabilidad\nen cada intervención', color: C.blue },
  ];

  bigStats.forEach((st, i) => {
    const bx = 0.6 + i * 4.15;
    s.addShape(pres.ShapeType.roundRect, {
      x: bx, y: 1.4, w: 3.85, h: 2.2,
      fill: { color: st.color },
      rectRadius: 0.1,
    });
    s.addText(st.value, {
      x: bx, y: 1.5, w: 3.85, h: 1.1, fontSize: 56,
      fontFace: FONT_TITLE, color: C.white, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
    s.addText(st.label, {
      x: bx, y: 2.65, w: 3.85, h: 0.7, fontSize: 14,
      fontFace: FONT_BODY, color: C.white,
      align: 'center', lineSpacingMultiple: 1.3, isTextBox: true, margin: 0,
    });
  });

  // Benefit rows
  const benefits = [
    { title: 'Presupuestos automáticos', desc: 'Se generan al aprobar el pedido de repuestos, con cálculo exacto de materiales + mano de obra + 22% gastos generales' },
    { title: 'Cliente informado 24/7', desc: 'Portal propio con timeline de avance, aprobación online y acceso a facturas — sin necesidad de llamar' },
    { title: 'Stock siempre actualizado', desc: 'Al entregar un repuesto de depósito se descuenta automáticamente. Alertas de stock mínimo en tiempo real' },
    { title: 'Facturación sin demoras', desc: 'Al completar una orden de trabajo se genera la factura. Control de morosidad con alertas graduales' },
  ];

  benefits.forEach((b, i) => {
    const by = 4.0 + i * 0.8;
    s.addShape(pres.ShapeType.ellipse, {
      x: 1.0, y: by + 0.1, w: 0.25, h: 0.25,
      fill: { color: C.cyan },
    });
    s.addText(b.title, {
      x: 1.5, y: by, w: 3.5, h: 0.35, fontSize: 13,
      fontFace: FONT_BODY, color: C.navy, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(b.desc, {
      x: 5.0, y: by, w: 7.5, h: 0.5, fontSize: 11,
      fontFace: FONT_BODY, color: C.grey, lineSpacingMultiple: 1.3,
      isTextBox: true, margin: 0,
    });
  });

  addFooter(s, false);
  s.addNotes('Usar estos números para anclar el valor. El 70% puede ajustarse según métricas reales.');
}

// ══════════════════════════════════════════════════
// SLIDE 9: TECNOLOGÍA
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.offWhite);

  s.addText('Tecnología pensada para simplicidad', {
    x: 0.8, y: 0.45, w: 11, h: 0.7, fontSize: 32,
    fontFace: FONT_TITLE, color: C.navy, bold: true,
    isTextBox: true, margin: 0,
  });

  const techItems = [
    { title: 'Aplicación web única', desc: 'Un solo archivo HTML que corre en cualquier navegador. Sin instalaciones, sin servidores complejos, sin dependencias externas.', color: C.blue },
    { title: 'Funciona offline', desc: 'Los datos se guardan en el navegador (localStorage). La app funciona sin conexión a internet — ideal para técnicos en campo.', color: C.teal },
    { title: 'Responsive', desc: 'Se adapta a celulares, tablets y escritorio. Los técnicos pueden cargar informes desde el teléfono en el campo.', color: C.accent },
    { title: 'Preparada para escalar', desc: 'Arquitectura modular lista para conectar a base de datos en la nube (Supabase, Firebase) cuando la empresa lo requiera.', color: C.success },
  ];

  techItems.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const tx = 0.6 + col * 6.3;
    const ty = 1.5 + row * 2.6;

    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: ty, w: 5.9, h: 2.2,
      fill: { color: C.white },
      shadow: { type: 'outer', blur: 8, offset: 3, angle: 135, color: '000000', opacity: 0.08 },
      rectRadius: 0.12,
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: tx + 0.2, y: ty + 0.3, w: 0.8, h: 0.8,
      fill: { color: t.color }, transparency: 85,
      rectRadius: 0.08,
    });
    s.addText(String(i + 1), {
      x: tx + 0.2, y: ty + 0.3, w: 0.8, h: 0.8, fontSize: 24,
      fontFace: FONT_TITLE, color: t.color, bold: true,
      align: 'center', valign: 'middle', isTextBox: true, margin: 0,
    });
    s.addText(t.title, {
      x: tx + 1.2, y: ty + 0.3, w: 4.3, h: 0.4, fontSize: 16,
      fontFace: FONT_BODY, color: C.navy, bold: true,
      isTextBox: true, margin: 0,
    });
    s.addText(t.desc, {
      x: tx + 1.2, y: ty + 0.8, w: 4.3, h: 1.1, fontSize: 12,
      fontFace: FONT_BODY, color: C.grey, lineSpacingMultiple: 1.4,
      isTextBox: true, margin: 0,
    });
  });

  addFooter(s, false);
  s.addNotes('Enfatizar que NO necesita infraestructura compleja. Un archivo HTML.');
}

// ══════════════════════════════════════════════════
// SLIDE 10: CIERRE
// ══════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.dark);

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 7.5,
    fill: { color: C.navy },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: -2, y: -3, w: 12, h: 12,
    fill: { color: C.blue }, transparency: 80,
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 8, y: 3, w: 8, h: 8,
    fill: { color: C.teal }, transparency: 80,
  });

  s.addText('ASCENSOR', {
    x: 2, y: 1.2, w: 9.3, h: 0.8, fontSize: 48,
    fontFace: FONT_TITLE, color: C.white, bold: true,
    align: 'center', charSpacing: 6, isTextBox: true, margin: 0,
  });
  s.addText('TECH', {
    x: 2, y: 1.9, w: 9.3, h: 0.8, fontSize: 48,
    fontFace: FONT_TITLE, color: C.cyan, bold: true,
    align: 'center', charSpacing: 6, isTextBox: true, margin: 0,
  });

  s.addText('Conectamos cada parte de su operación\nen un flujo inteligente y transparente', {
    x: 2, y: 3.3, w: 9.3, h: 0.8, fontSize: 18,
    fontFace: FONT_BODY, color: C.greyLt,
    align: 'center', lineSpacingMultiple: 1.4, isTextBox: true, margin: 0,
  });

  // CTA box
  s.addShape(pres.ShapeType.roundRect, {
    x: 4, y: 4.5, w: 5.3, h: 1.3,
    fill: { color: C.cyan },
    rectRadius: 0.12,
  });
  s.addText('Solicite una demostración', {
    x: 4, y: 4.6, w: 5.3, h: 0.6, fontSize: 22,
    fontFace: FONT_BODY, color: C.navy, bold: true,
    align: 'center', isTextBox: true, margin: 0,
  });
  s.addText('Personalizable para su empresa', {
    x: 4, y: 5.15, w: 5.3, h: 0.4, fontSize: 13,
    fontFace: FONT_BODY, color: C.dark,
    align: 'center', isTextBox: true, margin: 0,
  });

  // Contact area
  s.addText('info@ascensortech.com  ·  +54 11 XXXX-XXXX  ·  ascensortech.com', {
    x: 2, y: 6.2, w: 9.3, h: 0.4, fontSize: 13,
    fontFace: FONT_BODY, color: '7788BB',
    align: 'center', isTextBox: true, margin: 0,
  });

  s.addNotes('Slide de cierre. Invitar a una demo personalizada.');
}

// ─── SAVE ───
const outPath = '/home/user/elevator-app/AscensorTech-Presentacion.pptx';
pres.writeFile({ fileName: outPath }).then(() => {
  console.log('Saved to ' + outPath);
}).catch(err => {
  console.error('Error:', err);
});
