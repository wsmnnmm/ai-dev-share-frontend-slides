const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'WSM';
pptx.company = 'WSM';
pptx.subject = 'AI交流局分享 V2';
pptx.title = 'AI 编程不是魔法';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'zh-CN'
};

const C = {
  ink: '201A15',
  sub: '6F6559',
  bg: 'F7F2EA',
  paper: 'FFF9F1',
  orange: 'FF6B35',
  green: '2AA876',
  blue: '4361EE',
  red: 'C9483C',
  yellow: 'D89313',
  mint: '98D4BB',
  lavender: 'C7B8EA',
  pink: 'F4B8C5',
  sky: 'A8D8EA',
  cream: 'FFE6A7'
};

function slideBase() {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  s.addText('', {
    x: 0.28, y: 0.2, w: 12.55, h: 6.7,
    fill: { color: C.paper },
    line: { color: 'DDD3C5', width: 1 }
  });
  const tabs = [C.mint, C.lavender, C.pink, C.sky, C.cream];
  tabs.forEach((color, i) => {
    s.addText('', {
      x: 12.45, y: 0.8 + i * 1.1, w: 0.28, h: 0.8,
      fill: { color },
      line: { color, width: 0.5 }
    });
  });
  return s;
}

function addHeader(s, page, tag) {
  s.addText(`PAGE ${String(page).padStart(2, '0')}`, {
    x: 0.7, y: 0.45, w: 1.2, h: 0.25,
    fontFace: 'Courier New', fontSize: 9, color: '7A5B43', bold: true,
    charSpace: 1.2
  });
  s.addText(tag, {
    x: 1.95, y: 0.42, w: 2.8, h: 0.3,
    fontFace: 'Aptos', fontSize: 9, color: '256D52', bold: true,
    margin: 4,
    fill: { color: 'E7F6EF' },
    line: { color: 'BFE6D3', width: 1 },
    radius: 8
  });
}

function addTitle(s, title, subtitle='') {
  s.addText(title, {
    x: 0.75, y: 0.92, w: 10.5, h: 1.0,
    fontFace: 'Aptos Display', fontSize: 26, bold: true, color: C.ink,
    breakLine: false, fit: 'shrink'
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.78, y: 1.86, w: 10.4, h: 0.42,
      fontFace: 'Aptos', fontSize: 12, color: C.sub,
      breakLine: false, fit: 'shrink'
    });
  }
}

function addFooter(s, left, right) {
  s.addText(left, {
    x: 0.75, y: 6.45, w: 2.3, h: 0.2,
    fontFace: 'Courier New', fontSize: 8, color: '8D8173'
  });
  s.addText(right, {
    x: 8.4, y: 6.45, w: 3.8, h: 0.2,
    align: 'right', fontFace: 'Aptos', fontSize: 8, color: '8D8173'
  });
}

function box(s, x, y, w, h, title, body, color='FFFFFF') {
  s.addText('', {
    x, y, w, h,
    fill: { color },
    line: { color: 'E6DDCF', width: 1 }
  });
  s.addText(title, {
    x: x + 0.14, y: y + 0.12, w: w - 0.28, h: 0.26,
    fontFace: 'Aptos', fontSize: 13, bold: true, color: C.ink,
    fit: 'shrink'
  });
  s.addText(body, {
    x: x + 0.14, y: y + 0.42, w: w - 0.28, h: h - 0.52,
    fontFace: 'Aptos', fontSize: 10.5, color: C.sub,
    valign: 'top', breakLine: false, fit: 'shrink'
  });
}

function bullets(s, x, y, w, h, items, color=C.sub, size=14) {
  const runs = [];
  items.forEach((item) => {
    runs.push({ text: item, options: { bullet: { indent: 12 }, breakLine: true } });
  });
  s.addText(runs, {
    x, y, w, h,
    fontFace: 'Aptos', fontSize: size, color,
    paraSpaceAfterPt: 10, breakLine: false, fit: 'shrink', valign: 'top'
  });
}

// Slide 1
{
  const s = slideBase();
  s.addText('AI 交流局 / 分享 V2', {
    x: 0.8, y: 0.42, w: 2.5, h: 0.22,
    fontFace: 'Courier New', fontSize: 10, color: '7A5B43', bold: true
  });
  s.addText('面向老板 + 程序员 + 小白', {
    x: 9.65, y: 0.42, w: 2.2, h: 0.24,
    fontFace: 'Aptos', fontSize: 9, color: '256D52', bold: true,
    fill: { color: 'E7F6EF' }, line: { color: 'BFE6D3', width: 1 },
    margin: 4, align: 'center'
  });
  s.addText('AI 编程不是魔法，\n它更像一条从玩具到员工的升级路', {
    x: 0.8, y: 1.05, w: 8.8, h: 1.45,
    fontFace: 'Aptos Display', fontSize: 26, bold: true, color: C.ink,
    fit: 'shrink'
  });
  s.addText('从“随便玩玩”到“真的能交付”，中间差的不是模型，而是方法。', {
    x: 0.82, y: 2.45, w: 8.2, h: 0.3,
    fontFace: 'Courier New', fontSize: 11, color: '8C5B22'
  });
  s.addText('', {
    x: 0.8, y: 2.95, w: 8.8, h: 1.0,
    fill: { color: 'FFFDFC' },
    line: { color: 'E7DACA', width: 1.2 }
  });
  s.addText('这不是工具说明书，也不是模型横评。我想用更人话的方式，讲清楚 AI 编程为什么一边让人兴奋，一边让人崩溃；以及普通人到底该怎么一阶一阶用起来，而不是一上来就把自己交给“全自动”。', {
    x: 1.0, y: 3.15, w: 8.4, h: 0.62,
    fontFace: 'Aptos', fontSize: 13, color: C.sub,
    fit: 'shrink'
  });
  box(s, 0.8, 4.25, 2.7, 1.45, '六个阶段', '从一句话做小游戏，到像带 AI 员工一样管理系统。', 'FFFFFF');
  box(s, 3.75, 4.25, 2.7, 1.45, '我的判断', 'AI 不是神仙，更像一个很猛、但很会擅自发挥的实习生。', 'FFFFFF');
  box(s, 6.7, 4.25, 2.7, 1.45, '真实案例', '不讲空架构图，只讲我自己折腾过、跑通过的 qclaw / WorkBuddy。', 'FFFFFF');
  box(s, 9.65, 4.25, 2.35, 1.45, '带走什么', '明天就能上手的学习路径、沟通方法和协作套路。', 'FFFFFF');
  addFooter(s, '01 / Cover', '先别急着全自动，先搞懂你在哪一层');
}

// Slide 2
{
  const s = slideBase();
  addHeader(s, 2, '为什么大家一边兴奋一边失望');
  addTitle(s, '因为很多人拿到的是“答案生成器”，\n但真正要用的是“任务执行系统”');
  box(s, 0.75, 2.0, 5.55, 2.8, '常见幻觉', '• 我都用上最强模型了，为什么项目还是做不出来？\n• 它刚才还很聪明，为什么下一轮又开始胡说？\n• 别人都说 AI 能替代程序员了，怎么我还在修 prompt？', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 2.8, '我现在的理解', '• 争议往往不是 AI 不强，而是你给它的角色不清楚。\n• 你让它当搜索框，它就给答案；你让它当执行位，就得给任务、边界和验收。\n• 很多人不是不会用 AI，而是直接跳过了学习曲线。', 'FFFDFC');
  s.addText('一句话：AI 编程今天最大的问题，不是模型不够强，而是用户脑子里还没有“分阶段使用”的地图。', {
    x: 0.85, y: 5.25, w: 11.2, h: 0.55, fontFace: 'Aptos', fontSize: 14,
    color: C.ink, bold: true, fill: { color: 'FFF3EA' }, line: { color: 'FFD8C6', width: 1 }, margin: 10
  });
  addFooter(s, '02 / Why', '大家骂的往往不是 AI，本质上是在骂错误期待');
}

// Slide 3
{
  const s = slideBase();
  addHeader(s, 3, '六阶段总览');
  addTitle(s, '我把 AI 编程大致分成 6 个阶段', '这不是官方标准，更像我自己踩坑后总结出来的一张“驾驶证地图”。');
  const stages = [
    ['0️⃣ 玩具阶段', '一句话生成小游戏、网页 demo，最容易爽，最不容易落地。'],
    ['1️⃣ 工具阶段', 'Cursor 这类一体化 IDE，开始能做商业产品，但大项目会丢上下文。'],
    ['2️⃣ 副驾阶段', 'VS Code + 插件协作，像有个副驾驶一直帮你看地图。'],
    ['3️⃣ 搭档阶段', 'CLI 工具开始接管整段流程，但还是要你手动发车。'],
    ['4️⃣ 实习生阶段', '你开始用业务语言下任务，AI 直接产出 PPT、分析、代码等结果。'],
    ['5️⃣ 全职员工阶段', '自动触发、自动巡检、自动派活，像在带一支 AI 团队。']
  ];
  stages.forEach((it, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    box(s, 0.75 + col * 3.85, 2.15 + row * 1.6, 3.45, 1.25, it[0], it[1], 'FFFFFF');
  });
  addFooter(s, '03 / Stages', '不是谁先用到 5 阶最厉害，而是谁知道自己当前适合第几阶');
}

// Slide 4
{
  const s = slideBase();
  addHeader(s, 4, '0️⃣-1️⃣ 最容易上头的阶段');
  addTitle(s, '玩具阶段和工具阶段，\n为什么最容易让人误判 AI 能力？');
  box(s, 0.75, 2.0, 5.55, 3.0, '0️⃣ 玩具阶段', '• 输入一句话，它立刻给你一个能跑的小 demo，比如贪吃蛇、小网页、小工具。\n• 爽点很强，因为 30 秒就能看到结果。\n• 但复杂度一上去，你会发现“看起来像成了”，其实只是个玩具。', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 3.0, '1️⃣ 工具阶段', '• Cursor 这类工具就像手机摄影：直出就很好看，速度也快。\n• 你开始能做真正的产品，不只是 demo。\n• 但项目一大、文件一多、上下文一长，它就会像一个突然失忆的高手。', 'FFFDFC');
  s.addText('风趣一点说：前两阶段最像“AI 甜品区”——很好吃，但你不能拿甜品当主食。', {
    x: 0.85, y: 5.35, w: 11.15, h: 0.5, fontFace: 'Aptos', fontSize: 14,
    color: C.ink, bold: true, fill: { color: 'FFF3EA' }, line: { color: 'FFD8C6', width: 1 }, margin: 10
  });
  addFooter(s, '04 / Toy → Tool', '能做出东西，不等于能稳定交付东西');
}

// Slide 5
{
  const s = slideBase();
  addHeader(s, 5, '2️⃣-3️⃣ 开始真正像生产力');
  addTitle(s, '副驾阶段和搭档阶段，\nAI 才开始真正进入工作流');
  box(s, 0.75, 2.0, 5.55, 3.0, '2️⃣ 副驾阶段', '• 你还在开车，但旁边终于坐了个靠谱一点的副驾。\n• 它帮你看地图、查资料、补上下文。\n• 缺点是你还是得盯着屏幕，AI 不能自己出去跑腿。', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 3.0, '3️⃣ 搭档阶段', '• 到 CLI 这一步，AI 更像一个你能叫得动的搭档。\n• 不只会聊天，还能动文件、跑脚本、读目录、交付结果。\n• 但它还是不会自己醒来，所有任务都要你手动发起。', 'FFFDFC');
  addFooter(s, '05 / Copilot → Partner', '从“陪你干活”到“替你干一段活”，差别很大');
}

// Slide 6
{
  const s = slideBase();
  addHeader(s, 6, '4️⃣-5️⃣ 高级阶段');
  addTitle(s, '实习生阶段和全职员工阶段，\n为什么不适合小白直接跳？');
  box(s, 0.75, 2.0, 5.55, 2.75, '4️⃣ 实习生阶段', '• 你开始不用讲代码细节，而是讲业务目标。\n• 代码开始隐形，结果开始显形。\n• 这时候拉开差距的，不是会不会写 prompt，而是你懂不懂业务、懂不懂验收。', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 2.75, '5️⃣ 全职员工阶段', '• 自动触发、自动巡检、自动派活，已经不只是工具，而像在带团队。\n• 好处是手动少了，速度快了。\n• 坏处是：你如果不会管理，它会以极高速度把错误也自动化。', 'FFFDFC');
  s.addText('$1000+', {
    x: 0.9, y: 5.15, w: 2.2, h: 0.48, fontFace: 'Courier New', fontSize: 28,
    bold: true, color: C.red
  });
  s.addText('这是高级自动化可能对应的专业级日消耗量级。不是每个人都该一上来就开这档，就像新手不建议直接上 F1 赛道。', {
    x: 2.2, y: 5.1, w: 9.8, h: 0.6, fontFace: 'Aptos', fontSize: 13, color: C.sub,
    fit: 'shrink'
  });
  addFooter(s, '06 / Intern → Employee', '高级阶段拼的已经不是提示词，而是管理能力');
}

// Slide 7
{
  const s = slideBase();
  addHeader(s, 7, '我的核心类比');
  addTitle(s, 'AI 不是万能工具，\n而是一个能力很强但需要被管理的“实习生”');
  box(s, 0.75, 2.0, 5.55, 3.05, '它像实习生的地方', '• 做得快，学得快，态度还常常很好。\n• 但会误解需求、脑补背景、擅自发挥。\n• 你要是不盯，它不是偷懒，它是“热情地做错”。', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 3.05, '所以人类负责什么', '• 负责目标：到底要解决什么问题。\n• 负责边界：什么要做，什么不能做。\n• 负责验收：做到什么程度才算完成。\n• 负责后果：最终上线和签字的人，永远不是 AI。', 'FFFDFC');
  addFooter(s, '07 / Intern Metaphor', '你不是在“使用工具”，你是在“管理代理”');
}

// Slide 8
{
  const s = slideBase();
  addHeader(s, 8, '真正的分水岭');
  addTitle(s, '和 AI 协作的关键，不是聊天能力，\n而是把模糊想法变成清晰任务');
  s.addText('我要做个东西 → 解决什么问题 → 明确不做什么 → 怎么才算完成', {
    x: 0.85, y: 2.1, w: 11.0, h: 0.4, fontFace: 'Courier New', fontSize: 16,
    color: '8C5B22', bold: true, align: 'center'
  });
  box(s, 0.75, 2.7, 3.7, 1.8, '问题', '不是“帮我做个产品”，而是“帮我解决哪个具体麻烦”。', 'FFFFFF');
  box(s, 4.75, 2.7, 3.35, 1.8, '限制', '时间、成本、平台、权限、数据边界，不说清楚就等着返工。', 'FFFFFF');
  box(s, 8.4, 2.7, 3.7, 1.8, '验收', '没写完工标准，AI 就会默认“差不多就行”，而你通常不会满意。', 'FFFFFF');
  s.addText('坏需求：\n“帮我做一个 AI 产品。”\n\n好需求：\n“帮我做一个面向 SEO 团队的小工具，输入 URL 后追踪 AI 搜索是否引用了我的内容。先做 Web MVP，不接 CRM，不做团队协作，7 天内要能给 3 个种子用户试用。”', {
    x: 0.95, y: 4.95, w: 10.9, h: 1.2, fontFace: 'Courier New', fontSize: 12,
    color: 'F5EADC', fill: { color: '2B2521' }, line: { color: 'B55F1C', width: 1 }, margin: 12
  });
  addFooter(s, '08 / Requirement', '需求越具体，结果越可控；表达越含糊，返工越多');
}

// Slide 9
{
  const s = slideBase();
  addHeader(s, 9, '保留第一版：tabbit 学习方法');
  addTitle(s, '别一口气冲终局，\ntabbit 式学习更像“踩台阶”，不是“坐火箭”');
  box(s, 0.75, 2.0, 5.55, 3.0, '我理解的 tabbit 方法', '• 先从能立刻看见结果的小台阶开始，建立正反馈。\n• 每上一层，只多学一种新能力：先学生成，再学协作，再学自动化。\n• 不要因为别人晒“全自动 AI 公司”就觉得自己也得当天毕业。', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 3.0, '为什么这方法重要', '• 因为 AI 工具不是取消了驾照，而是换了车。\n• 你如果连方向盘都没摸熟，就去开自动驾驶，最后往往不是省事，是翻车。\n• 分阶段学习，才有机会把“惊艳一次”变成“稳定产出很多次”。', 'FFFDFC');
  s.addText('没人会因为你直接跳到第 5 阶，就多给你一张证书；但你很可能先多交几张账单。', {
    x: 0.85, y: 5.35, w: 11.1, h: 0.5, fontFace: 'Aptos', fontSize: 14,
    color: C.ink, bold: true, fill: { color: 'FFF3EA' }, line: { color: 'FFD8C6', width: 1 }, margin: 10
  });
  addFooter(s, '09 / Tabbit', '学习 AI，不是比谁跳得快，而是比谁每一阶都站稳');
}

// Slide 10
{
  const s = slideBase();
  addHeader(s, 10, '我的工作流');
  addTitle(s, '我现在怎么放工具位置：\nManager、Worker、qclaw、WorkBuddy 各自站哪');
  s.addText('人 → Manager → Worker\nqclaw ↔ WorkBuddy', {
    x: 1.1, y: 2.0, w: 10.5, h: 0.9, fontFace: 'Aptos Display', fontSize: 24,
    align: 'center', color: C.blue, bold: true
  });
  box(s, 0.75, 3.1, 2.7, 1.8, '人', '定目标 / 定边界 / 做判断 / 做验收', 'F6FBFF');
  box(s, 3.72, 3.1, 2.7, 1.8, 'Manager', '拆任务 / 派工 / 看进度 / 最后收口', 'F4F1FF');
  box(s, 6.69, 3.1, 2.7, 1.8, 'Worker', '执行 / 产出 / 回报 / 落文件', 'F2FFF7');
  box(s, 9.66, 3.1, 2.35, 1.8, '桥接层', 'qclaw 远程触发\nWorkBuddy 结果承接', 'FFF9F2');
  s.addText('核心思路不是“谁最强”，而是“谁该先上，谁该补位，谁该最后背锅”。', {
    x: 0.9, y: 5.35, w: 11.0, h: 0.45, fontFace: 'Aptos', fontSize: 14,
    color: C.ink, bold: true
  });
  addFooter(s, '10 / Workflow', 'AI 时代最像管理学，不像魔法课');
}

// Slide 11
{
  const s = slideBase();
  addHeader(s, 11, 'qclaw vs WorkBuddy');
  addTitle(s, '一句话区别：\nqclaw 像“微信遥控本地助理”，WorkBuddy 像“桌面 AI 工作台”');
  box(s, 0.75, 2.0, 5.55, 2.25, 'qclaw 更像什么', '• 你人在外面，也能把家里/办公室那台电脑叫起来干活。\n• 优势是“远程触发”和“本地执行”。\n• 适合文件、桌面任务、轻操作自动化。', 'FFFDFC');
  box(s, 6.55, 2.0, 5.55, 2.25, 'WorkBuddy 更像什么', '• 更像一个办公交付台，负责整理、展示、承接结果。\n• 优势是“桌面交付”和“结果组织”。\n• 适合办公流程、文档任务、产出查看。', 'FFFDFC');
  s.addImage({ path: 'assets/qclaw-workbuddy-prompt.png', x: 0.95, y: 4.45, w: 5.35, h: 1.55 });
  s.addImage({ path: 'assets/qclaw-workbuddy-result.png', x: 6.6, y: 4.45, w: 5.25, h: 1.55 });
  addFooter(s, '11 / qclaw × WorkBuddy', '先跑通一条土路，再谈高速公路');
}

// Slide 12
{
  const s = slideBase();
  addHeader(s, 12, '成本不是彩蛋');
  addTitle(s, '为什么我越来越觉得：\n高级 AI 编程不是炫技题，而是管理题');
  s.addText('$404', {
    x: 0.95, y: 2.0, w: 2.1, h: 0.6, fontFace: 'Courier New', fontSize: 30,
    bold: true, color: C.red
  });
  s.addText('一段失控长对话给我的提醒。不是花不起，而是它暴露了方法上的漏洞：边界不清、验收不明、上下文过长。', {
    x: 2.2, y: 2.05, w: 6.1, h: 0.6, fontFace: 'Aptos', fontSize: 13, color: C.sub,
    fit: 'shrink'
  });
  s.addImage({ path: 'assets/token-cost.png', x: 8.55, y: 1.8, w: 3.15, h: 2.15 });
  box(s, 0.75, 3.25, 11.35, 2.0, '我现在更在意的 3 件事', '• 低阶段拼的是“会不会用”；高阶段拼的是“会不会控成本”。\n• AI 员工越多、自动化越深，越需要明确角色、边界和巡检机制。\n• 真正的高级玩家，不一定最会写 prompt，但一定很会做管理。', 'FFFDFC');
  addFooter(s, '12 / Cost & Control', '别把账单看成惩罚，它其实是管理能力的镜子');
}

// Slide 13
{
  const s = slideBase();
  addHeader(s, 13, '明天就能用的三条建议');
  addTitle(s, '如果你听完只想带走 3 件事，我建议是这三件');
  box(s, 0.75, 2.0, 3.65, 2.7, '1. 别跳级', '先从你当前驾驭得住的一层开始。不会开车的人，别先研究自动驾驶公司怎么调度车队。', 'FFFDFC');
  box(s, 4.72, 2.0, 3.65, 2.7, '2. 学会写清楚任务', '说清问题、限制、不做什么、完成标准。你越清楚，AI 越像助手；你越模糊，AI 越像喜剧演员。', 'FFFDFC');
  box(s, 8.69, 2.0, 3.4, 2.7, '3. 把 AI 当团队成员', '会拆解、会派工、会验收、会复盘，才会越用越顺；否则就是一次次“它怎么又理解错了”。', 'FFFDFC');
  addFooter(s, '13 / Advice', '把神话拆掉，AI 才会真正开始帮你做事');
}

// Slide 14
{
  const s = slideBase();
  addHeader(s, 14, 'Closing');
  addTitle(s, 'AI 时代真正值钱的，\n不是更会按按钮，\n而是更会讲清问题、拆明需求、做出判断');
  s.addText('', {
    x: 0.8, y: 2.35, w: 10.8, h: 1.15,
    fill: { color: 'FFFDFC' }, line: { color: 'E7DACA', width: 1.1 }
  });
  s.addText('AI 可以提升执行效率，但它替代不了人的判断力。未来最重要的能力，不只是“会不会用工具”，而是“你能不能把问题讲清楚、把任务拆明白、把边界定准确”。', {
    x: 1.0, y: 2.62, w: 10.35, h: 0.58, fontFace: 'Aptos', fontSize: 14, color: C.sub,
    fit: 'shrink'
  });
  box(s, 0.8, 3.95, 2.8, 1.6, '给老板', '别只问“AI 能不能替人”，先问“我们有没有能力管理它”。', 'FFFFFF');
  box(s, 3.88, 3.95, 2.8, 1.6, '给程序员', '别只比模型，先把拆解、执行、验收这三段链路搭起来。', 'FFFFFF');
  box(s, 6.96, 3.95, 2.8, 1.6, '给小白', '先从一个能看见结果的小任务开始，不用一上来就想着“AI 创业”。', 'FFFFFF');
  box(s, 10.04, 3.95, 2.0, 1.6, '给我自己', '继续折腾，但别放弃判断力。', 'FFFFFF');
  addFooter(s, '14 / Closing', '谢谢，欢迎交流：别问哪家最强，先问自己现在在哪一层');
}

pptx.writeFile({ fileName: 'AI-Dev-Share-Presentation-v2.pptx' });
