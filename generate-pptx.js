const PptxGenJS = require('pptxgenjs');

// 创建新的 PPTX 实例
const pptx = new PptxGenJS();

// 设置演示文稿属性
pptx.author = 'WSM';
pptx.title = 'AI 编程工具实战分享';
pptx.subject = 'AI DEV SHARE';
pptx.company = 'AI Developer';

// 定义配色方案 - 延续 HTML 的深色科技风格
const COLORS = {
  bgDark: '0b1020',
  bgSecondary: '121933',
  textPrimary: 'e7ecf5',
  textSecondary: '9eaac1',
  accent: '7cc8ff',
  accent2: '7ee787',
  accent2Dark: '2a8a4a',
  danger: 'ff8e8e',
  panelBg: '1a1f38'
};

// 创建深色背景的工厂函数
const makeDarkSlide = () => {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.bgDark };
  return slide;
};

// ==================== SLIDE 1: 封面页 ====================
const slide1 = makeDarkSlide();

// 顶部标签
slide1.addText('AI DEV SHARE / FRONTEND SLIDES', {
  x: 0.8, y: 0.5, w: 4, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});

// 状态标签
slide1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 4.9, y: 0.5, w: 1.2, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide1.addText('我的实战版', {
  x: 4.9, y: 0.5, w: 1.2, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

// 主标题
slide1.addText('我怎么把 AI 编程工具\n一点点接进研发工作流', {
  x: 0.8, y: 1.2, w: 8.5, h: 1.2,
  fontSize: 36,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.2
});

// 摘要
slide1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 2.6, w: 8.5, h: 0.8,
  fill: { color: COLORS.panelBg },
  line: { color: COLORS.accent, width: 0.5 },
  rectRadius: 0.15
});
slide1.addText('这不是工具说明书，也不是模型横评。更像一次阶段性复盘：我最近怎么用 Codex、Claude Code、Skills、MCP、qclaw、WorkBuddy 去做事、踩坑、协作和收口。', {
  x: 1.0, y: 2.7, w: 8.1, h: 0.6,
  fontSize: 12,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  valign: 'middle'
});

// 四个卡片
const cards = [
  { title: 'Codex', subtitle: '更像执行位', desc: '我更关心的不是谁最聪明，而是谁更适合放在工作流的哪个位置。' },
  { title: 'Claude Code', subtitle: '更像拆解位', desc: '这套分享会讲真实成本、桥接体验、协作分工和我目前的判断。' },
  { title: 'Skills + MCP', subtitle: '补能力缺口', desc: '不是先装满生态，而是先知道什么问题该交给什么能力去补。' },
  { title: 'qclaw + WorkBuddy', subtitle: '往真实任务里接', desc: '有些地方我已经很确定，有些地方我还在试，但都值得拿出来讲真话。' }
];

cards.forEach((card, i) => {
  const x = 0.8 + (i * 2.3);
  const y = 3.7;

  slide1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 2.1, h: 1.8,
    fill: { color: COLORS.panelBg },
    line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
    rectRadius: 0.2
  });

  slide1.addText(card.title, {
    x: x + 0.15, y: y + 0.15, w: 1.6, h: 0.3,
    fontSize: 12,
    color: COLORS.textPrimary,
    fontFace: 'Arial',
    bold: true
  });

  slide1.addText(card.subtitle, {
    x: x + 0.15, y: y + 0.4, w: 1.8, h: 0.25,
    fontSize: 8,
    color: COLORS.accent,
    fontFace: 'Arial'
  });

  slide1.addText(card.desc, {
    x: x + 0.15, y: y + 0.7, w: 1.8, h: 0.9,
    fontSize: 9,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

// 页脚
slide1.addText('01 / Cover', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide1.addText('Arrow keys / Space / Swipe', {
  x: 6.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 2: 48小时事件线 ====================
const slide2 = makeDarkSlide();

slide2.addText('PAGE 02', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide2.addText('48 小时事件线', {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide2.addText('我不是先有成熟方法，\n而是先在混乱里把东西跑起来', {
  x: 0.8, y: 0.9, w: 8.5, h: 0.8,
  fontSize: 24,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.1
});

slide2.addText('这页我想先把背景说清楚：这套分享不是事后总结出来的漂亮框架，而是我最近 48 小时边折腾边形成的真实经验。', {
  x: 0.8, y: 1.8, w: 8.5, h: 0.4,
  fontSize: 11,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

// 左侧：时间线
slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 2.4, w: 4.3, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

const timelineItems = [
  '昨天开始正式折腾 Codex 和 Claude Code 桥接，多 agent 互相讨论和 battle。',
  '开始尝试 qclaw / OpenClaw / WorkBuddy 的联通，希望它们能交接任务、互相评审。',
  '把话题从"工具能不能聊"切到"能不能看 bug 文档、分析 bug、改 bug"。',
  '过程中出现了 token 成本、权限、桥接稳定性、上下文膨胀这些很真实的问题。',
  '最后我意识到，应该把这些经验沉淀成一套能拿去分享、还能继续迭代的网页 PPT。'
];

timelineItems.forEach((item, i) => {
  const y = 2.55 + (i * 0.55);
  slide2.addShape(pptx.shapes.OVAL, {
    x: 1.0, y: y, w: 0.3, h: 0.3,
    fill: { color: COLORS.accent, transparency: 70 },
    line: { color: COLORS.accent, width: 0.5 }
  });
  slide2.addText(String(i + 1), {
    x: 1.0, y: y, w: 0.3, h: 0.3,
    fontSize: 10,
    color: COLORS.accent,
    fontFace: 'Arial',
    align: 'center',
    valign: 'middle'
  });
  slide2.addText(item, {
    x: 1.45, y: y - 0.05, w: 3.5, h: 0.5,
    fontSize: 9,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

// 右侧：判断
slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 2.4, w: 4, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide2.addText('我这页的判断', {
  x: 5.5, y: 2.55, w: 3.6, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const judgments = [
  '这些内容的可信度，不来自定义，而来自它们刚刚真的发生过。',
  '我自己也不是一开始就清楚，只是先有动作，再慢慢收成方法。',
  '所以后面每一页，我都尽量讲"我怎么做"和"我现在怎么想"。'
];

judgments.forEach((item, i) => {
  slide2.addText('• ' + item, {
    x: 5.5, y: 2.95 + (i * 0.65), w: 3.6, h: 0.6,
    fontSize: 10,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

slide2.addText('02 / Timeline', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide2.addText('从折腾到项目化沉淀', {
  x: 6.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 3: 工作流总图 ====================
const slide3 = makeDarkSlide();

slide3.addText('PAGE 03', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide3.addText('工作流总图', {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide3.addText('我现在最笃定的一点：\n工具名不重要，交接链路更重要', {
  x: 0.8, y: 0.85, w: 8.5, h: 0.8,
  fontSize: 24,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.1
});

slide3.addText('如果只记工具名，最后很容易变成一页生态地图。但我最近最有启发的，其实是慢慢看清楚谁该先上、谁该补位、谁该做最后验收。', {
  x: 0.8, y: 1.75, w: 8.5, h: 0.4,
  fontSize: 11,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

// 工作流节点
const workflowNodes = ['人定目标', 'Claude Code\n拆解/复核', 'Codex\n执行/验证', 'Skills/MCP\n/Bridge', '人验收\n/追问', '沉淀成\nSOP/页面'];
const nodeWidth = 1.2;
const startX = 0.8;
const nodeY = 2.35;

workflowNodes.forEach((node, i) => {
  const x = startX + (i * 1.5);

  slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y: nodeY, w: nodeWidth, h: 0.7,
    fill: { color: COLORS.panelBg },
    line: { color: 'FFFFFF', width: 0.3, transparency: 85 },
    rectRadius: 0.15
  });
  slide3.addText(node, {
    x, y: nodeY, w: nodeWidth, h: 0.7,
    fontSize: 8,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    align: 'center',
    valign: 'middle'
  });

  // 添加箭头
  if (i < workflowNodes.length - 1) {
    slide3.addText('→', {
      x: x + nodeWidth, y: nodeY, w: 0.3, h: 0.7,
      fontSize: 16,
      color: COLORS.accent,
      fontFace: 'Arial',
      align: 'center',
      valign: 'middle'
    });
  }
});

// 分工理解
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 3.3, w: 4.3, h: 2.1,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide3.addText('我现在的分工理解', {
  x: 1.0, y: 3.45, w: 3.9, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const splitItems = [
  'Codex：更像进工程现场干活的人，读目录、改代码、验证结果。',
  'Claude Code：更像拆解、互评、补视角的人，适合帮我把混乱问题讲清楚。',
  'qclaw / WorkBuddy：不是主角，而是工作任务的交接通道和集成场。'
];

splitItems.forEach((item, i) => {
  slide3.addText('• ' + item, {
    x: 1.0, y: 3.85 + (i * 0.5), w: 3.9, h: 0.5,
    fontSize: 9,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

// 为什么重要
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 3.3, w: 4, h: 2.1,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide3.addText('这条链路为什么重要', {
  x: 5.5, y: 3.45, w: 3.6, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const whyItems = [
  '工具多不是问题，问题是它们如果并列摆着，就很难形成产出。',
  '我真正想讲的是：谁先拆，谁先干，谁来复核，谁来背交付。',
  '这也是后面讲 Skills、MCP 和 Bridge 时，我最看重的判断标准。'
];

whyItems.forEach((item, i) => {
  slide3.addText('• ' + item, {
    x: 5.5, y: 3.85 + (i * 0.5), w: 3.6, h: 0.5,
    fontSize: 9,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

slide3.addText('03 / Workflow', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide3.addText('工具要挂在流程上', {
  x: 6.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 4: Codex ====================
const slide4 = makeDarkSlide();

slide4.addText('PAGE 04', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.0, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide4.addText('Codex', {
  x: 2.5, y: 0.4, w: 1.0, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide4.addText('Codex：让它先试，你再审', {
  x: 0.8, y: 0.85, w: 8.5, h: 0.6,
  fontSize: 28,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true
});

// 左侧内容
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 1.6, w: 5.2, h: 3.5,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide4.addText('我现在最认可 Codex 的位置，不是"比谁更会解释"，而是它更像能进项目现场帮我做事的执行位。', {
  x: 1.0, y: 1.75, w: 4.8, h: 0.5,
  fontSize: 11,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

const codexPoints = [
  '先让它摸目录、看文档、归纳当前上下文，再进入具体任务。',
  '适合做"给我改、给我验、给我落文件"的动作，不只是聊天。',
  '当任务边界清楚时，它的推进速度比我自己手动切来切去快很多。'
];

codexPoints.forEach((item, i) => {
  slide4.addText('• ' + item, {
    x: 1.0, y: 2.35 + (i * 0.45), w: 4.8, h: 0.45,
    fontSize: 10,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

// 代码片段
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 3.8, w: 4.8, h: 1.15,
  fill: { color: '0d1528' },
  line: { color: COLORS.accent, width: 0.5, transparency: 70 },
  rectRadius: 0.15
});
slide4.addText('# 我现在更常见的用法\n1. 先让它总结代码库和限制\n2. 再让它只改一小块\n3. 最后看 diff / 结果 / 风险', {
  x: 1.15, y: 3.9, w: 4.5, h: 0.95,
  fontSize: 9,
  color: COLORS.accent,
  fontFace: 'Consolas',
  lineSpaceMult: 1.3
});

// 右侧坑
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.2, y: 1.6, w: 3.1, h: 3.5,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide4.addText('我踩过的坑', {
  x: 6.4, y: 1.75, w: 2.7, h: 0.4,
  fontSize: 14,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true
});

const pitfalls = [
  '上下文不够时，它会认真地往错方向干。',
  '如果我自己没定义验收标准，最后很容易"看起来改了，其实没收口"。',
  '把大而糊的任务直接丢进去，效率会明显下降。'
];

pitfalls.forEach((item, i) => {
  slide4.addText('• ' + item, {
    x: 6.4, y: 2.3 + (i * 0.8), w: 2.7, h: 0.75,
    fontSize: 10,
    color: COLORS.danger,
    fontFace: 'Arial',
    valign: 'top'
  });
});

slide4.addText('04 / Codex', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide4.addText('执行前先收边界', {
  x: 6.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 5: Claude Code ====================
const slide5 = makeDarkSlide();

slide5.addText('PAGE 05', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide5.addText('Claude Code', {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide5.addText('Claude Code：对我来说最强的\n不是写代码，而是把乱问题拆开', {
  x: 0.8, y: 0.85, w: 8.5, h: 0.8,
  fontSize: 22,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.1
});

slide5.addText('如果要我用一句话形容它现在在我工作流里的位置，我会说：它更像一个很会帮我拆题、对照、互评的搭子，但成本感要一直带着。', {
  x: 0.8, y: 1.75, w: 8.5, h: 0.4,
  fontSize: 11,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

// 左侧
slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 2.3, w: 4.3, h: 2.8,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide5.addText('我现在最常拿它干什么', {
  x: 1.0, y: 2.45, w: 3.9, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const ccUseCases = [
  '先帮我拆解混乱问题，再决定值不值得交给另一个 agent 去做。',
  '做方案比较、互评、battle、复盘，这类"需要补视角"的事它很适合。',
  '面对我这种脑子里有很多碎片、但还没收成结构的时候，它特别有用。'
];

ccUseCases.forEach((item, i) => {
  slide5.addText('• ' + item, {
    x: 1.0, y: 2.85 + (i * 0.55), w: 3.9, h: 0.55,
    fontSize: 10,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

// 右侧 - 数据卡片
slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 2.3, w: 4, h: 2.8,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide5.addText('我踩过的坑也很明确', {
  x: 5.5, y: 2.45, w: 3.6, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const metrics = [
  { value: '$404.42', desc: '长对话一旦失控，token 和费用会先把问题暴露出来。' },
  { value: '$83.87', desc: '不是只有最强模型贵，第二梯队长期跑下来也会有体感。' },
  { value: '225.2M', desc: '我现在会主动问自己：这个任务到底值不值得长上下文。' }
];

metrics.forEach((m, i) => {
  const y = 2.85 + (i * 0.7);
  slide5.addText(m.value, {
    x: 5.5, y, w: 1.2, h: 0.4,
    fontSize: 18,
    color: COLORS.accent,
    fontFace: 'Consolas',
    bold: true
  });
  slide5.addText(m.desc, {
    x: 6.7, y, w: 2.4, h: 0.65,
    fontSize: 8,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

slide5.addText('下一页我直接放真实截图，因为这件事靠口头讲不够有冲击力。', {
  x: 5.5, y: 4.8, w: 3.6, h: 0.25,
  fontSize: 8,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  italic: true
});

slide5.addText('05 / Claude Code', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide5.addText('拆解强，但要管上下文和成本', {
  x: 6.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 6: 成本数据 ====================
const slide6 = makeDarkSlide();

slide6.addText('PAGE 06', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.5, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide6.addText('成本 / 实战数据', {
  x: 2.5, y: 0.4, w: 1.5, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide6.addText('我开始高频用 AI 之后，\n最先被教育的其实是成本感', {
  x: 0.8, y: 0.85, w: 8.5, h: 0.8,
  fontSize: 26,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.1
});

slide6.addText('这张图我一定会留着，因为它逼着我从"这个工具好不好用"继续往下想，去看"这套工作流能不能长期跑"。', {
  x: 0.8, y: 1.8, w: 8.5, h: 0.4,
  fontSize: 11,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

// 截图占位
slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 2.35, w: 5.2, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: COLORS.textSecondary, width: 0.5, transparency: 70 },
  rectRadius: 0.2
});
slide6.addText('[Token Cost Screenshot]', {
  x: 0.8, y: 2.35, w: 5.2, h: 2.9,
  fontSize: 14,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});
slide6.addText('过去一段时间的模型调用与费用分布。这个图不是为了吓人，而是提醒我：方法不稳，token 会先告诉你。', {
  x: 1.0, y: 5.0, w: 4.8, h: 0.25,
  fontSize: 8,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  italic: true
});

// 右侧
slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.2, y: 2.35, w: 3.1, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide6.addText('这张图给我的 3 个提醒', {
  x: 6.4, y: 2.5, w: 2.7, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const costReminders = [
  '多 agent 协作不是"免费想象力"，它真的会回到成本上。',
  '我现在更在意的是：什么任务值得长对话，什么任务应该快拆快干。',
  '如果只拿工具好不好用来判断，很容易忽略"能不能长期用"的问题。'
];

costReminders.forEach((item, i) => {
  slide6.addText('• ' + item, {
    x: 6.4, y: 2.95 + (i * 0.7), w: 2.7, h: 0.65,
    fontSize: 10,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

slide6.addText('06 / Cost', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide6.addText('真实数据页', {
  x: 6.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 7: Bridge ====================
const slide7 = makeDarkSlide();

slide7.addText('PAGE 07', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.8, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide7.addText('qclaw / WorkBuddy', {
  x: 2.5, y: 0.4, w: 1.8, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide7.addText('这部分我不想吹大，\n我更想讲一件"已经跑通"的事', {
  x: 0.8, y: 0.85, w: 8.5, h: 0.8,
  fontSize: 24,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.1
});

slide7.addText('qclaw / OpenClaw 和 WorkBuddy 我还没有大规模带进长期真实项目，但我至少已经验证了一条可用的桥接路径：共享目录 + 定时轮询。这一页我只讲我已经确认过的东西。', {
  x: 0.8, y: 1.75, w: 8.5, h: 0.5,
  fontSize: 10,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

// 左侧
slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 2.4, w: 4.3, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide7.addText('我已经验证到哪一步了', {
  x: 1.0, y: 2.55, w: 3.9, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const bridgeItems = [
  'WorkBuddy 和 qclaw 的双向交接，至少在文件级别已经跑通。',
  '当前采用的方案是共享目录 + 定时轮询，不是最优雅，但足够稳、足够低依赖。',
  '这对我最大的启发不是"桥好酷"，而是跨 agent 协作必须先有一条最朴素的落地通道。',
  '我下一步才会把它往 bug 分析、任务派发、结果回传这些真实工作里带。'
];

bridgeItems.forEach((item, i) => {
  slide7.addText('• ' + item, {
    x: 1.0, y: 2.95 + (i * 0.5), w: 3.9, h: 0.5,
    fontSize: 9,
    color: COLORS.textSecondary,
    fontFace: 'Arial',
    valign: 'top'
  });
});

// 右侧截图
slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 2.4, w: 4, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: COLORS.textSecondary, width: 0.5, transparency: 70 },
  rectRadius: 0.2
});
slide7.addText('[API Details Screenshot]', {
  x: 5.3, y: 2.4, w: 4, h: 2.9,
  fontSize: 12,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide7.addText('07 / Bridge', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide7.addText('先验证一条稳通路，再谈规模化使用', {
  x: 5, y: 5.2, w: 4.5, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// ==================== SLIDE 8: 复盘 ====================
const slide8 = makeDarkSlide();

slide8.addText('PAGE 08', {
  x: 0.8, y: 0.4, w: 2, h: 0.3,
  fontSize: 10,
  color: COLORS.accent,
  fontFace: 'Arial',
  charSpacing: 2
});
slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fill: { color: COLORS.accent2, transparency: 80 },
  line: { color: COLORS.accent2, width: 0.5 },
  rectRadius: 0.1
});
slide8.addText('双 agent 复盘', {
  x: 2.5, y: 0.4, w: 1.2, h: 0.3,
  fontSize: 8,
  color: COLORS.accent2,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});

slide8.addText('这次做 PPT，\n双 agent 没有自动让我更稳', {
  x: 0.8, y: 0.85, w: 8.5, h: 0.8,
  fontSize: 28,
  color: COLORS.textPrimary,
  fontFace: 'Arial',
  bold: true,
  lineSpaceMult: 1.1
});

slide8.addText('这个结论我想明确讲出来：如果多一个 agent 只是多一层意见、多一轮建议、多一个分叉，那它不但不会更稳，反而会增加返工。这次做 PPT 的过程，我就明显感受到了这件事。', {
  x: 0.8, y: 1.75, w: 8.5, h: 0.5,
  fontSize: 10,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});

// 左侧截图
slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 2.4, w: 4.3, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: COLORS.textSecondary, width: 0.5, transparency: 70 },
  rectRadius: 0.2
});
slide8.addText('[Battle Consensus Screenshot]', {
  x: 0.8, y: 2.4, w: 4.3, h: 2.9,
  fontSize: 12,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'center',
  valign: 'middle'
});
slide8.addText('battle 本身并不等于高质量。真正有价值的是：它有没有明确承担批判、对照、校验的职责。', {
  x: 1.0, y: 5.05, w: 4, h: 0.25,
  fontSize: 8,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  italic: true
});

// 右侧问题
slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 2.4, w: 4, h: 2.9,
  fill: { color: COLORS.panelBg },
  line: { color: 'FFFFFF', width: 0.3, transparency: 90 },
  rectRadius: 0.2
});

slide8.addText('我这次踩到的 4 个问题', {
  x: 5.5, y: 2.55, w: 3.6, h: 0.3,
  fontSize: 10,
  color: COLORS.accent2,
  fontFace: 'Arial',
  bold: true
});

const retroProblems = [
  'brief 太宽，讨论可以发散，但执行需要收口，这两者没有被分开。',
  '第二个 agent 很多时候只是在提建议，没有扮演明确的验收或批判角色。',
  '做 PPT / 做前端这种任务，审美和内容判断如果不统一，就会不断返工。',
  '没有统一验收标准时，双 agent 只是把"可能错的地方"翻倍了。'
];

retroProblems.forEach((item, i) => {
  slide8.addText('• ' + item, {
    x: 5.5, y: 2.95 + (i * 0.5), w: 3.6, h: 0.5,
    fontSize: 9,
    color: COLORS.danger,
    fontFace: 'Arial',
    valign: 'top'
  });
});

slide8.addText('08 / Retro', {
  x: 0.5, y: 5.2, w: 3, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial'
});
slide8.addText('双 agent 不是神话，得有清晰角色和验收', {
  x: 5, y: 5.2, w: 4.5, h: 0.3,
  fontSize: 9,
  color: COLORS.textSecondary,
  fontFace: 'Arial',
  align: 'right'
});

// 保存文件
pptx.writeFile({ fileName: 'AI-Dev-Share-Presentation.pptx' })
  .then(fileName => {
    console.log('PPTX created: ' + fileName);
  })
  .catch(err => {
    console.error('Error: ' + err);
  });
