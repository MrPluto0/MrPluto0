const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const COLORS = {
  navy: "0B1B2B",
  deep: "065A82",
  teal: "1C7293",
  mint: "02C39A",
  ink: "0F172A",
  text: "1F2937",
  muted: "64748B",
  bg: "F2F7FA",
  card: "FFFFFF",
  line: "D6E1EA",
  warn: "F59E0B",
  ok: "10B981",
  bad: "EF4444",
};

const FONTS = {
  head: "Trebuchet MS",
  body: "Calibri",
};

function addBg(slide, fill) {
  slide.background = { color: fill };
}

function addLeftRail(slide, color) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 0.22,
    h: SLIDE_H,
    fill: { color },
    line: { color },
  });
}

function addTopMeta(slide, label) {
  const x = 0.55;
  const y = 0.42;
  const w = SLIDE_W - x - 0.6;
  slide.addText(label, {
    x,
    y,
    w,
    h: 0.3,
    fontFace: FONTS.body,
    fontSize: 11,
    color: COLORS.muted,
  });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.55,
    y: 0.75,
    w: SLIDE_W - 1.1,
    h: 0.75,
    fontFace: FONTS.head,
    fontSize: 38,
    bold: true,
    color: COLORS.ink,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.55,
      y: 1.55,
      w: SLIDE_W - 1.1,
      h: 0.5,
      fontFace: FONTS.body,
      fontSize: 16,
      color: COLORS.muted,
    });
  }
}

function addCard(slide, { x, y, w, h, title, body, accent = COLORS.mint }) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    fill: { color: COLORS.card },
    line: { color: COLORS.line, width: 1 },
    radius: 10,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.25,
    y: y + 0.26,
    w: 0.18,
    h: h - 0.52,
    fill: { color: accent },
    line: { color: accent },
    radius: 4,
  });

  slide.addText(title, {
    x: x + 0.55,
    y: y + 0.28,
    w: w - 0.8,
    h: 0.45,
    fontFace: FONTS.head,
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
  });

  if (body) {
    slide.addText(body, {
      x: x + 0.55,
      y: y + 0.78,
      w: w - 0.8,
      h: h - 1.05,
      fontFace: FONTS.body,
      fontSize: 14,
      color: COLORS.text,
      valign: "top",
    });
  }
}

function addIconCircle(slide, { x, y, r = 0.2, fill = COLORS.mint, glyph = "•" }) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x - r,
    y: y - r,
    w: r * 2,
    h: r * 2,
    fill: { color: fill },
    line: { color: fill },
  });
  slide.addText(glyph, {
    x: x - r,
    y: y - r + 0.02,
    w: r * 2,
    h: r * 2,
    fontFace: FONTS.head,
    fontSize: Math.max(12, Math.round(r * 60)),
    bold: true,
    color: "FFFFFF",
    align: "center",
    valign: "mid",
  });
}

function slideTitleOnly(slide, label, title, subtitle) {
  addBg(slide, COLORS.bg);
  addLeftRail(slide, COLORS.deep);
  addTopMeta(slide, label);
  addTitle(slide, title, subtitle);
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.55,
    y: SLIDE_H - 0.45,
    w: SLIDE_W - 1.1,
    h: 0.28,
    fontFace: FONTS.body,
    fontSize: 10,
    color: COLORS.muted,
  });
}

function s1() {
  const slide = pptx.addSlide();
  addBg(slide, COLORS.navy);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: SLIDE_W,
    h: SLIDE_H,
    fill: { color: COLORS.navy },
    line: { color: COLORS.navy },
  });

  slide.addShape(pptx.ShapeType.ellipse, {
    x: -2,
    y: -1.8,
    w: 7.2,
    h: 7.2,
    fill: { color: COLORS.deep, transparency: 35 },
    line: { color: COLORS.deep, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: SLIDE_W - 6.2,
    y: SLIDE_H - 5.6,
    w: 7.0,
    h: 7.0,
    fill: { color: COLORS.mint, transparency: 60 },
    line: { color: COLORS.mint, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0,
    y: 0,
    w: 0.28,
    h: SLIDE_H,
    fill: { color: COLORS.mint },
    line: { color: COLORS.mint },
    radius: 10,
  });

  slide.addText("无人机技术与应用概览", {
    x: 0.75,
    y: 2.25,
    w: SLIDE_W - 1.5,
    h: 1.0,
    fontFace: FONTS.head,
    fontSize: 50,
    bold: true,
    color: "FFFFFF",
  });
  slide.addText("平台 · 载荷 · 算法 · 合规 · 交付", {
    x: 0.78,
    y: 3.25,
    w: SLIDE_W - 1.56,
    h: 0.5,
    fontFace: FONTS.body,
    fontSize: 18,
    color: "DCE7F1",
  });
  slide.addText(`2026-03`, {
    x: 0.78,
    y: 4.0,
    w: SLIDE_W - 1.56,
    h: 0.35,
    fontFace: FONTS.body,
    fontSize: 12,
    color: "A9C2D6",
  });
}

function s2() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "目录", "本次分享覆盖什么", "从技术底座到落地交付的关键点");

  const startX = 0.55;
  const startY = 2.35;
  const gap = 0.35;
  const w = (SLIDE_W - 1.1 - gap) / 2;
  const h = 1.35;

  const items = [
    ["技术栈分层", "平台/飞控/导航/通信/载荷/软件"],
    ["关键指标", "续航、载重、精度、抗风、冗余"],
    ["应用场景", "测绘、巡检、农业、安防、应急"],
    ["安全与合规", "空域、隐私、数据、失控保护"],
    ["选型与实施", "需求→方案→试点→规模化运维"],
    ["趋势与机会", "AI视觉、蜂群、VTOL、卫星通信"],
  ];

  items.forEach(([t, b], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    addCard(slide, {
      x: startX + col * (w + gap),
      y: startY + row * (h + gap),
      w,
      h,
      title: `${i + 1}. ${t}`,
      body: b,
      accent: col === 0 ? COLORS.mint : COLORS.teal,
    });
  });

  addFooter(slide, "建议时长：10–15 分钟 · 可按需要删减或加深某一块");
}

function s3() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "技术栈", "无人机系统分层", "硬件平台 + 任务载荷 + 软件能力共同决定交付效果");

  const cx = SLIDE_W * 0.62;
  const cy = 4.15;

  slide.addShape(pptx.ShapeType.roundRect, {
    x: cx - 1.4,
    y: cy - 0.55,
    w: 2.8,
    h: 1.1,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.line, width: 1 },
    radius: 14,
  });
  slide.addText("无人机平台", {
    x: cx - 1.4,
    y: cy - 0.4,
    w: 2.8,
    h: 0.8,
    fontFace: FONTS.head,
    fontSize: 22,
    bold: true,
    color: COLORS.ink,
    align: "center",
    valign: "mid",
  });

  const nodes = [
    { t: "机体/动力", b: "结构、桨叶、电机、续航", x: cx - 3.9, y: cy - 2.35, c: COLORS.teal, g: "P" },
    { t: "飞控", b: "姿态控制、冗余、故障保护", x: cx + 2.0, y: cy - 2.35, c: COLORS.mint, g: "F" },
    { t: "导航定位", b: "GNSS/RTK、IMU、视觉", x: cx - 4.25, y: cy + 0.25, c: COLORS.mint, g: "N" },
    { t: "通信链路", b: "数传/图传、抗干扰、加密", x: cx + 2.35, y: cy + 0.25, c: COLORS.teal, g: "C" },
    { t: "任务载荷", b: "相机、激光雷达、喷洒", x: cx - 3.9, y: cy + 2.85, c: COLORS.teal, g: "L" },
    { t: "地面站/云", b: "规划、任务管理、数据闭环", x: cx + 2.0, y: cy + 2.85, c: COLORS.mint, g: "S" },
  ];

  nodes.forEach((n) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: n.x,
      y: n.y,
      w: 3.2,
      h: 1.45,
      fill: { color: "FFFFFF" },
      line: { color: COLORS.line, width: 1 },
      radius: 14,
    });
    addIconCircle(slide, { x: n.x + 0.45, y: n.y + 0.5, r: 0.23, fill: n.c, glyph: n.g });
    slide.addText(n.t, {
      x: n.x + 0.8,
      y: n.y + 0.25,
      w: 2.25,
      h: 0.45,
      fontFace: FONTS.head,
      fontSize: 16,
      bold: true,
      color: COLORS.ink,
    });
    slide.addText(n.b, {
      x: n.x + 0.8,
      y: n.y + 0.68,
      w: 2.25,
      h: 0.68,
      fontFace: FONTS.body,
      fontSize: 12,
      color: COLORS.muted,
    });

    const nodeY = n.y + 0.72;
    const centerLeft = cx - 1.4;
    const centerRight = cx + 1.4;
    const nodeLeft = n.x;
    const nodeRight = n.x + 3.2;
    const fromX = n.x < cx ? nodeRight : centerRight;
    const toX = n.x < cx ? centerLeft : nodeLeft;
    const lineX = Math.min(fromX, toX);
    const lineW = Math.max(0, Math.abs(toX - fromX));

    slide.addShape(pptx.ShapeType.line, {
      x: lineX,
      y: nodeY,
      w: lineW,
      h: 0,
      line: { color: "BFD3E2", width: 2, dash: "dash" },
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 2.25,
    w: 5.4,
    h: 4.95,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.line, width: 1 },
    radius: 14,
  });
  slide.addText("交付视角：三条主线", {
    x: 0.95,
    y: 2.5,
    w: 4.9,
    h: 0.45,
    fontFace: FONTS.head,
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
  });
  const bullets = [
    ["可靠性", "失控保护、冗余、抗风与维护"],
    ["数据链路", "采集→处理→标注→分析→归档"],
    ["任务闭环", "规划、执行、复盘、SOP、报表"],
  ];
  bullets.forEach(([t, b], i) => {
    const y = 3.2 + i * 1.35;
    addIconCircle(slide, { x: 1.15, y: y + 0.2, r: 0.18, fill: i === 1 ? COLORS.teal : COLORS.mint, glyph: String(i + 1) });
    slide.addText(t, {
      x: 1.45,
      y,
      w: 4.3,
      h: 0.35,
      fontFace: FONTS.head,
      fontSize: 16,
      bold: true,
      color: COLORS.ink,
    });
    slide.addText(b, {
      x: 1.45,
      y: y + 0.42,
      w: 4.6,
      h: 0.55,
      fontFace: FONTS.body,
      fontSize: 12,
      color: COLORS.muted,
    });
  });
}

function s4() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "指标", "怎么衡量一台无人机是否适配任务", "先定义工况，再谈参数：环境、载荷、作业模式决定答案");

  const gridX = 0.55;
  const gridY = 2.35;
  const cols = 3;
  const gap = 0.32;
  const cardW = (SLIDE_W - 1.1 - gap * (cols - 1)) / cols;
  const cardH = 1.55;

  const items = [
    ["续航/航程", "电池效率、巡航速度、留安全余量", "45–90 分钟"],
    ["载重/载荷", "相机/雷达/喷洒与供电接口", "0.5–10 kg"],
    ["定位精度", "RTK/PPP 与标定流程", "厘米级"],
    ["抗风/稳定", "风场、控制律、螺旋桨匹配", "6–10 m/s"],
    ["链路与带宽", "图传延迟、数传可靠性、加密", "< 200 ms"],
    ["安全冗余", "双 IMU、双 GNSS、失效保护", "可控降落"],
  ];

  items.forEach(([t, b, stat], i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = gridX + col * (cardW + gap);
    const y = gridY + row * (cardH + gap);
    addCard(slide, {
      x,
      y,
      w: cardW,
      h: cardH,
      title: t,
      body: b,
      accent: col === 1 ? COLORS.teal : COLORS.mint,
    });
    slide.addText(stat, {
      x: x + 0.55,
      y: y + 1.05,
      w: cardW - 0.85,
      h: 0.45,
      fontFace: FONTS.head,
      fontSize: 20,
      bold: true,
      color: COLORS.deep,
    });
  });

  addFooter(slide, "提示：参数来自典型行业范围，最终以测试工况与合规约束为准");
}

function s5() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "场景", "典型行业应用", "每个场景都要绑定：目标、覆盖范围、数据产出与SLA");

  const x0 = 0.55;
  const y0 = 2.35;
  const gap = 0.3;
  const cols = 3;
  const w = (SLIDE_W - 1.1 - gap * (cols - 1)) / cols;
  const h = 1.65;

  const apps = [
    ["测绘建模", "航测/倾斜摄影、点云与正射", "M"],
    ["电力巡检", "可见光+红外，缺陷识别与复核", "E"],
    ["安防应急", "热成像、喊话照明、快速部署", "A"],
    ["农业植保", "变量喷洒、地块管理与合规", "G"],
    ["物流配送", "短距高频、末端交付与调度", "D"],
    ["园区运维", "例行巡航、告警联动、报表", "O"],
  ];

  apps.forEach(([t, b, g], i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = x0 + col * (w + gap);
    const y = y0 + row * (h + gap);
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h,
      fill: { color: "FFFFFF" },
      line: { color: COLORS.line, width: 1 },
      radius: 14,
    });
    addIconCircle(slide, { x: x + 0.55, y: y + 0.55, r: 0.26, fill: col === 1 ? COLORS.teal : COLORS.mint, glyph: g });
    slide.addText(t, {
      x: x + 0.95,
      y: y + 0.25,
      w: w - 1.15,
      h: 0.45,
      fontFace: FONTS.head,
      fontSize: 18,
      bold: true,
      color: COLORS.ink,
    });
    slide.addText(b, {
      x: x + 0.95,
      y: y + 0.72,
      w: w - 1.15,
      h: 0.85,
      fontFace: FONTS.body,
      fontSize: 13,
      color: COLORS.muted,
    });
  });
}

function s6() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "趋势", "2026 前后值得关注的能力升级", "硬件性能提升放缓，差异化更多来自软件与系统工程");

  const x = 0.75;
  const y = 2.6;
  const w = SLIDE_W - 1.5;

  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 3.9,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.line, width: 1 },
    radius: 14,
  });

  slide.addText("能力路线（从可用到可规模化）", {
    x: x + 0.35,
    y: y + 0.35,
    w: w - 0.7,
    h: 0.4,
    fontFace: FONTS.head,
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
  });

  const points = [
    ["AI 视觉感知", "目标识别、缺陷检测、避障与自定位"],
    ["高精定位普及", "RTK/PPP 标准化、标定流程产品化"],
    ["蜂群/协同", "多机调度、任务分配、冲突消解"],
    ["通信升级", "5G 专网、卫星链路、端到端加密"],
    ["VTOL 与复合翼", "长航程 + 垂直起降，拓展工况"],
  ];

  const left = x + 0.6;
  const top = y + 1.05;
  const rowH = 0.58;
  const lineX = left + 0.25;

  slide.addShape(pptx.ShapeType.line, {
    x: lineX,
    y: top - 0.05,
    w: 0,
    h: rowH * (points.length - 1) + 0.1,
    line: { color: "BFD3E2", width: 3 },
  });

  points.forEach(([t, b], i) => {
    const yy = top + i * rowH;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: lineX - 0.12,
      y: yy + 0.03,
      w: 0.24,
      h: 0.24,
      fill: { color: i % 2 === 0 ? COLORS.mint : COLORS.teal },
      line: { color: i % 2 === 0 ? COLORS.mint : COLORS.teal },
    });
    slide.addText(t, {
      x: left + 0.55,
      y: yy - 0.02,
      w: w - 1.3,
      h: 0.28,
      fontFace: FONTS.head,
      fontSize: 16,
      bold: true,
      color: COLORS.ink,
    });
    slide.addText(b, {
      x: left + 0.55,
      y: yy + 0.25,
      w: w - 1.3,
      h: 0.32,
      fontFace: FONTS.body,
      fontSize: 12,
      color: COLORS.muted,
    });
  });
}

function s7() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "安全与合规", "把风险降到可控：流程 + 技术 + 证据", "合规不是“能不能飞”，而是“能否持续交付并承担责任”");

  const leftX = 0.7;
  const topY = 2.4;
  const leftW = 6.2;
  const rightX = 7.15;
  const rightW = SLIDE_W - rightX - 0.6;

  slide.addShape(pptx.ShapeType.roundRect, {
    x: leftX,
    y: topY,
    w: leftW,
    h: 4.8,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.line, width: 1 },
    radius: 14,
  });
  slide.addText("合规清单（示例）", {
    x: leftX + 0.35,
    y: topY + 0.35,
    w: leftW - 0.7,
    h: 0.4,
    fontFace: FONTS.head,
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
  });

  const checklist = [
    ["空域/审批", "飞行计划、禁飞区、临时管制"],
    ["人员资质", "机长/操作员培训与记录"],
    ["数据合规", "隐私处理、留存期限、访问控制"],
    ["设备安全", "失控保护、返航策略、冗余验证"],
    ["现场SOP", "起降区、围界、告知、应急预案"],
  ];
  checklist.forEach(([t, b], i) => {
    const yy = topY + 1.05 + i * 0.72;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: leftX + 0.35,
      y: yy + 0.06,
      w: 0.22,
      h: 0.22,
      fill: { color: COLORS.ok },
      line: { color: COLORS.ok },
    });
    slide.addText("✓", {
      x: leftX + 0.35,
      y: yy + 0.045,
      w: 0.22,
      h: 0.22,
      fontFace: FONTS.head,
      fontSize: 12,
      bold: true,
      color: "FFFFFF",
      align: "center",
      valign: "mid",
    });
    slide.addText(t, {
      x: leftX + 0.65,
      y: yy,
      w: leftW - 1.0,
      h: 0.28,
      fontFace: FONTS.head,
      fontSize: 15,
      bold: true,
      color: COLORS.ink,
    });
    slide.addText(b, {
      x: leftX + 0.65,
      y: yy + 0.28,
      w: leftW - 1.0,
      h: 0.28,
      fontFace: FONTS.body,
      fontSize: 11,
      color: COLORS.muted,
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: rightX,
    y: topY,
    w: rightW,
    h: 4.8,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.line, width: 1 },
    radius: 14,
  });
  slide.addText("风险矩阵（示例）", {
    x: rightX + 0.35,
    y: topY + 0.35,
    w: rightW - 0.7,
    h: 0.4,
    fontFace: FONTS.head,
    fontSize: 18,
    bold: true,
    color: COLORS.ink,
  });

  const mx = rightX + 0.55;
  const my = topY + 1.15;
  const mw = rightW - 1.1;
  const mh = 3.25;

  slide.addShape(pptx.ShapeType.rect, { x: mx, y: my, w: mw, h: mh, fill: { color: "FFFFFF" }, line: { color: COLORS.line, width: 1 } });
  slide.addShape(pptx.ShapeType.line, { x: mx + mw / 2, y: my, w: 0, h: mh, line: { color: COLORS.line, width: 1 } });
  slide.addShape(pptx.ShapeType.line, { x: mx, y: my + mh / 2, w: mw, h: 0, line: { color: COLORS.line, width: 1 } });

  slide.addShape(pptx.ShapeType.rect, { x: mx, y: my, w: mw / 2, h: mh / 2, fill: { color: COLORS.ok, transparency: 78 }, line: { color: COLORS.ok, transparency: 100 } });
  slide.addShape(pptx.ShapeType.rect, { x: mx + mw / 2, y: my, w: mw / 2, h: mh / 2, fill: { color: COLORS.warn, transparency: 78 }, line: { color: COLORS.warn, transparency: 100 } });
  slide.addShape(pptx.ShapeType.rect, { x: mx, y: my + mh / 2, w: mw / 2, h: mh / 2, fill: { color: COLORS.warn, transparency: 78 }, line: { color: COLORS.warn, transparency: 100 } });
  slide.addShape(pptx.ShapeType.rect, { x: mx + mw / 2, y: my + mh / 2, w: mw / 2, h: mh / 2, fill: { color: COLORS.bad, transparency: 80 }, line: { color: COLORS.bad, transparency: 100 } });

  slide.addText("影响", { x: mx + mw + 0.05, y: my + mh / 2 - 0.1, w: 0.35, h: 0.3, rotate: 90, fontFace: FONTS.body, fontSize: 11, color: COLORS.muted, align: "center" });
  slide.addText("发生概率", { x: mx + mw / 2 - 0.6, y: my + mh + 0.1, w: 1.2, h: 0.3, fontFace: FONTS.body, fontSize: 11, color: COLORS.muted, align: "center" });

  const tags = [
    { t: "低：SOP", x: mx + 0.25, y: my + 0.2, c: COLORS.ok },
    { t: "中：冗余", x: mx + mw / 2 + 0.2, y: my + mh / 2 + 0.2, c: COLORS.warn },
    { t: "高：禁飞", x: mx + mw / 2 + 0.2, y: my + 0.2, c: COLORS.bad },
  ];
  tags.forEach((tag) => {
    slide.addShape(pptx.ShapeType.roundRect, { x: tag.x, y: tag.y, w: 1.35, h: 0.35, fill: { color: tag.c }, line: { color: tag.c }, radius: 8 });
    slide.addText(tag.t, { x: tag.x, y: tag.y + 0.03, w: 1.35, h: 0.3, fontFace: FONTS.body, fontSize: 11, color: "FFFFFF", align: "center", valign: "mid" });
  });
}

function s8() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "实施", "从需求到规模化：建议的交付路径", "用试点验证闭环，用指标驱动扩容与运维");

  const x = 0.75;
  const y = 2.6;
  const w = SLIDE_W - 1.5;
  const h = 3.9;
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: "FFFFFF" }, line: { color: COLORS.line, width: 1 }, radius: 14 });

  const steps = [
    ["需求定义", "场景/范围/产出/频率"],
    ["平台选型", "机型/载荷/链路/合规"],
    ["试点验证", "SOP、指标、数据闭环"],
    ["规模部署", "人员、备件、工单体系"],
    ["持续优化", "算法迭代、报表、审计"],
  ];

  const stepW = (w - 0.8) / steps.length;
  const baseY = y + 1.2;
  steps.forEach(([t, b], i) => {
    const sx = x + 0.4 + i * stepW;
    const accent = i % 2 === 0 ? COLORS.mint : COLORS.teal;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: sx,
      y: baseY,
      w: stepW - 0.2,
      h: 2.25,
      fill: { color: COLORS.bg },
      line: { color: COLORS.line, width: 1 },
      radius: 12,
    });
    addIconCircle(slide, { x: sx + 0.35, y: baseY + 0.42, r: 0.2, fill: accent, glyph: String(i + 1) });
    slide.addText(t, {
      x: sx + 0.62,
      y: baseY + 0.22,
      w: stepW - 0.9,
      h: 0.4,
      fontFace: FONTS.head,
      fontSize: 15,
      bold: true,
      color: COLORS.ink,
    });
    slide.addText(b, {
      x: sx + 0.35,
      y: baseY + 0.72,
      w: stepW - 0.55,
      h: 1.4,
      fontFace: FONTS.body,
      fontSize: 12,
      color: COLORS.muted,
      valign: "top",
    });
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.rightArrow, {
        x: sx + stepW - 0.26,
        y: baseY + 0.95,
        w: 0.34,
        h: 0.35,
        fill: { color: "BFD3E2" },
        line: { color: "BFD3E2" },
      });
    }
  });

  slide.addText("建议交付物：测试报告、SOP、作业报表模板、风险评估、培训与运维手册", {
    x: x + 0.4,
    y: y + 0.55,
    w: w - 0.8,
    h: 0.45,
    fontFace: FONTS.body,
    fontSize: 13,
    color: COLORS.muted,
  });
}

function s9() {
  const slide = pptx.addSlide();
  slideTitleOnly(slide, "ROI", "成本与收益的常见计算方式", "把“飞一次”拆成可计量的工时、风险与数据价值");

  const x = 0.7;
  const y = 2.35;

  slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 6.35, h: 4.85, fill: { color: "FFFFFF" }, line: { color: COLORS.line, width: 1 }, radius: 14 });
  slide.addText("对比（示例：巡检/测绘类任务）", { x: x + 0.35, y: y + 0.35, w: 5.8, h: 0.35, fontFace: FONTS.head, fontSize: 18, bold: true, color: COLORS.ink });

  const rows = [
    ["指标", "人工/车辆", "无人机方案"],
    ["单次覆盖", "低–中", "中–高"],
    ["作业时长", "长", "短"],
    ["风险暴露", "高（高空/道路）", "低"],
    ["数据粒度", "离散/主观", "连续/可追溯"],
    ["复检与留证", "成本高", "成本低"],
  ];
  const tx = x + 0.35;
  const ty = y + 0.95;
  const tw = 5.8;
  const rh = 0.52;

  rows.forEach((r, i) => {
    const yy = ty + i * rh;
    const header = i === 0;
    slide.addShape(pptx.ShapeType.rect, {
      x: tx,
      y: yy,
      w: tw,
      h: rh,
      fill: { color: header ? COLORS.bg : "FFFFFF" },
      line: { color: COLORS.line, width: 1 },
    });
    const colW = [0.33, 0.33, 0.34].map((p) => tw * p);
    const xs = [tx, tx + colW[0], tx + colW[0] + colW[1]];
    for (let j = 0; j < 2; j++) {
      slide.addShape(pptx.ShapeType.line, { x: xs[j + 1], y: yy, w: 0, h: rh, line: { color: COLORS.line, width: 1 } });
    }
    const style = {
      fontFace: header ? FONTS.head : FONTS.body,
      fontSize: header ? 13 : 12,
      bold: header,
      color: header ? COLORS.ink : COLORS.text,
      valign: "mid",
    };
    slide.addText(r[0], { x: tx + 0.15, y: yy + 0.09, w: colW[0] - 0.25, h: rh - 0.18, ...style });
    slide.addText(r[1], { x: tx + colW[0] + 0.15, y: yy + 0.09, w: colW[1] - 0.25, h: rh - 0.18, ...style });
    slide.addText(r[2], { x: tx + colW[0] + colW[1] + 0.15, y: yy + 0.09, w: colW[2] - 0.25, h: rh - 0.18, ...style });
  });

  const rx = 7.25;
  const rw = SLIDE_W - rx - 0.6;
  slide.addShape(pptx.ShapeType.roundRect, { x: rx, y, w: rw, h: 4.85, fill: { color: "FFFFFF" }, line: { color: COLORS.line, width: 1 }, radius: 14 });
  slide.addText("价值分解（示例）", { x: rx + 0.35, y: y + 0.35, w: rw - 0.7, h: 0.35, fontFace: FONTS.head, fontSize: 18, bold: true, color: COLORS.ink });
  slide.addText("单位：相对贡献（0–100）", { x: rx + 0.35, y: y + 0.75, w: rw - 0.7, h: 0.28, fontFace: FONTS.body, fontSize: 11, color: COLORS.muted });

  const bars = [
    ["节省工时", 70, COLORS.mint],
    ["降低风险", 55, COLORS.teal],
    ["提升覆盖", 60, COLORS.mint],
    ["数据资产", 45, COLORS.teal],
  ];
  const bx = rx + 0.45;
  const by = y + 1.2;
  const bw = rw - 0.9;
  const barH = 0.55;
  bars.forEach(([t, v, c], i) => {
    const yy = by + i * 0.8;
    slide.addText(t, { x: bx, y: yy + 0.06, w: 1.4, h: 0.3, fontFace: FONTS.body, fontSize: 12, color: COLORS.text });
    slide.addShape(pptx.ShapeType.roundRect, { x: bx + 1.45, y: yy, w: bw - 1.7, h: barH, fill: { color: COLORS.bg }, line: { color: COLORS.bg }, radius: 10 });
    slide.addShape(pptx.ShapeType.roundRect, { x: bx + 1.45, y: yy, w: ((bw - 1.7) * v) / 100, h: barH, fill: { color: c }, line: { color: c }, radius: 10 });
    slide.addText(`${v}`, { x: bx + 1.45 + (bw - 1.7) + 0.05, y: yy + 0.06, w: 0.4, h: 0.3, fontFace: FONTS.body, fontSize: 11, color: COLORS.muted });
  });
}

function s10() {
  const slide = pptx.addSlide();
  addBg(slide, COLORS.navy);

  slide.addShape(pptx.ShapeType.roundRect, { x: 0, y: 0, w: 0.28, h: SLIDE_H, fill: { color: COLORS.mint }, line: { color: COLORS.mint }, radius: 10 });
  slide.addShape(pptx.ShapeType.ellipse, { x: SLIDE_W - 7.2, y: -1.6, w: 7.8, h: 7.8, fill: { color: COLORS.deep, transparency: 35 }, line: { color: COLORS.deep, transparency: 100 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: -2.2, y: SLIDE_H - 5.9, w: 7.4, h: 7.4, fill: { color: COLORS.mint, transparency: 62 }, line: { color: COLORS.mint, transparency: 100 } });

  slide.addText("结论", { x: 0.75, y: 1.25, w: 5.5, h: 0.6, fontFace: FONTS.head, fontSize: 44, bold: true, color: "FFFFFF" });
  slide.addText("三句话带走", { x: 0.78, y: 1.95, w: 5.5, h: 0.35, fontFace: FONTS.body, fontSize: 14, color: "CFE0EE" });

  const items = [
    ["先定义任务工况", "覆盖范围、环境、频率、数据产出与SLA"],
    ["交付要看闭环", "规划→执行→复盘→报表→审计，缺一不可"],
    ["合规是硬约束", "流程与证据体系比“参数更强”更重要"],
  ];
  items.forEach(([t, b], i) => {
    const y = 2.65 + i * 1.25;
    slide.addShape(pptx.ShapeType.roundRect, { x: 0.75, y, w: 11.8, h: 1.0, fill: { color: "10263A", transparency: 15 }, line: { color: "10263A", transparency: 100 }, radius: 14 });
    addIconCircle(slide, { x: 1.25, y: y + 0.5, r: 0.22, fill: i === 1 ? COLORS.teal : COLORS.mint, glyph: String(i + 1) });
    slide.addText(t, { x: 1.6, y: y + 0.2, w: 10.8, h: 0.35, fontFace: FONTS.head, fontSize: 20, bold: true, color: "FFFFFF" });
    slide.addText(b, { x: 1.6, y: y + 0.55, w: 10.8, h: 0.32, fontFace: FONTS.body, fontSize: 14, color: "DCE7F1" });
  });

  slide.addText("Q&A", { x: 0.78, y: 6.95, w: 3.0, h: 0.4, fontFace: FONTS.head, fontSize: 18, bold: true, color: "FFFFFF" });
}

s1();
s2();
s3();
s4();
s5();
s6();
s7();
s8();
s9();
s10();

const outPath = "/workspace/无人机_概览.pptx";
pptx.writeFile({ fileName: outPath });
