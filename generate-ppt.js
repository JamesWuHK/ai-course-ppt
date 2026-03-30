const pptxgen = require("pptxgenjs");

// 创建演示文稿
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "小企业AI入门实战课";
pres.author = "AI培训";

// 配色方案
const colors = {
  primary: "1A365D",      // 深蓝色
  secondary: "2B6CB0",    // 中蓝色
  accent: "ED8936",       // 橙色强调
  light: "EBF8FF",        // 浅蓝色背景
  dark: "1A202C",         // 深灰色文字
  gray: "718096",         // 灰色文字
  white: "FFFFFF",
  lightGray: "F7FAFC"
};

// 工具函数：创建阴影
const makeShadow = () => ({
  type: "outer",
  color: "000000",
  blur: 8,
  offset: 3,
  angle: 135,
  opacity: 0.12
});

// ========== 第1页：封面 ==========
let slide1 = pres.addSlide();
slide1.background = { color: colors.primary };

// 装饰性圆形
slide1.addShape(pres.shapes.OVAL, {
  x: -1.5, y: -1.5, w: 4, h: 4,
  fill: { color: colors.secondary, transparency: 60 }
});
slide1.addShape(pres.shapes.OVAL, {
  x: 8, y: 3.5, w: 3.5, h: 3.5,
  fill: { color: colors.accent, transparency: 50 }
});

// 主标题
slide1.addText("小企业AI入门实战课", {
  x: 0.5, y: 1.8, w: 9, h: 1.2,
  fontSize: 44, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center"
});

// 副标题
slide1.addText("从0基础到AI Agent智能经营", {
  x: 0.5, y: 3.0, w: 9, h: 0.6,
  fontSize: 28, fontFace: "Microsoft YaHei",
  color: colors.accent, align: "center"
});

// 课程时长
slide1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 3.5, y: 3.9, w: 3, h: 0.6,
  fill: { color: colors.white, transparency: 20 },
  rectRadius: 0.1
});
slide1.addText("2小时快速上手", {
  x: 3.5, y: 3.9, w: 3, h: 0.6,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.white, align: "center", valign: "middle"
});

// 底部说明
slide1.addText("专为0基础小企业老板设计", {
  x: 0.5, y: 5.0, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Microsoft YaHei",
  color: colors.white, align: "center", italic: true
});

// ========== 第2页：开场痛点 ==========
let slide2 = pres.addSlide();
slide2.background = { color: colors.lightGray };

// 顶部强调条
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.15,
  fill: { color: colors.accent }
});

// 标题
slide2.addText("你是不是有这种感觉？", {
  x: 0.5, y: 0.5, w: 9, h: 0.8,
  fontSize: 36, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true, align: "center"
});

// 痛点卡片
const painPoints = [
  "每天忙到飞起，但好像什么都没完成",
  "重复的事情做不完：回客户、写文案、算账、发朋友圈",
  "想学新东西，但时间不够用"
];

painPoints.forEach((point, i) => {
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 1.6 + i * 1.1, w: 7, h: 0.9,
    fill: { color: colors.white },
    shadow: makeShadow()
  });
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 1.6 + i * 1.1, w: 0.08, h: 0.9,
    fill: { color: colors.accent }
  });
  slide2.addText(point, {
    x: 1.8, y: 1.6 + i * 1.1, w: 6.5, h: 0.9,
    fontSize: 20, fontFace: "Microsoft YaHei",
    color: colors.dark, valign: "middle"
  });
});

// 底部解决方案
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 1.5, y: 4.8, w: 7, h: 0.7,
  fill: { color: colors.primary }
});
slide2.addText("今天我告诉你一个秘密：这些事情，以后你只需要说一句话，AI帮你全干了", {
  x: 1.5, y: 4.8, w: 7, h: 0.7,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: colors.white, align: "center", valign: "middle"
});

// ========== 第3页：课程目标 ==========
let slide3 = pres.addSlide();
slide3.background = { color: colors.white };

// 左侧装饰
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.3, h: 5.625,
  fill: { color: colors.primary }
});

// 标题
slide3.addText("这堂课你会学到什么", {
  x: 0.8, y: 0.4, w: 8.5, h: 0.8,
  fontSize: 32, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// 两列布局
// 左列：课程目标
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 1.5, w: 4, h: 3.5,
  fill: { color: colors.light }
});
slide3.addText("课程目标", {
  x: 0.8, y: 1.5, w: 4, h: 0.6,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true, align: "center", valign: "middle"
});
slide3.addText([
  { text: "理解AI Agent时代的工作方式", options: { bullet: true, breakLine: true } },
  { text: "学会用自然语言驱动AI完成经营任务", options: { bullet: true, breakLine: true } },
  { text: "掌握与AI协作的核心技巧", options: { bullet: true } }
], {
  x: 1.0, y: 2.2, w: 3.6, h: 2.5,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: colors.dark, paraSpaceAfter: 12
});

// 右列：核心转变
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 1.5, w: 4, h: 3.5,
  fill: { color: colors.primary }
});
slide3.addText("核心转变", {
  x: 5.2, y: 1.5, w: 4, h: 0.6,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center", valign: "middle"
});
slide3.addText([
  { text: "旧思维", options: { bold: true, breakLine: true } },
  { text: "人操作软件 → 软件执行任务", options: { breakLine: true, breakLine: true } },
  { text: "", options: { breakLine: true } },
  { text: "新思维", options: { bold: true, breakLine: true } },
  { text: "人描述需求 → AI Agent自动调用工具 → 完成任务", options: {} }
], {
  x: 5.4, y: 2.2, w: 3.6, h: 2.5,
  fontSize: 15, fontFace: "Microsoft YaHei",
  color: colors.white, paraSpaceAfter: 6
});

// ========== 第4页：什么是AI ==========
let slide4 = pres.addSlide();
slide4.background = { color: colors.lightGray };

// 顶部条
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.12,
  fill: { color: colors.secondary }
});

slide4.addText("什么是AI？", {
  x: 0.5, y: 0.4, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// AI解释卡片
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.3, w: 9, h: 2.2,
  fill: { color: colors.white },
  shadow: makeShadow()
});

slide4.addText("一个特别聪明的虚拟大脑", {
  x: 0.5, y: 1.3, w: 9, h: 0.7,
  fontSize: 24, fontFace: "Microsoft YaHei",
  color: colors.secondary, bold: true, align: "center", valign: "middle"
});

slide4.addText([
  { text: "它读过全世界所有的书、看过所有的视频、听过所有的对话", options: { breakLine: true } },
  { text: "所以你问它什么，它都能回答你", options: { breakLine: true } },
  { text: "就像一个24小时在线、不要工资、最渊博的朋友", options: {} }
], {
  x: 0.8, y: 2.1, w: 8.4, h: 1.2,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.dark, align: "center", paraSpaceAfter: 8
});

// 常见AI产品
slide4.addText("你可能用过的AI产品", {
  x: 0.5, y: 3.7, w: 9, h: 0.5,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

const aiProducts = [
  { name: "豆包", desc: "智能小助手" },
  { name: "Kimi", desc: "文书助手" },
  { name: "通义千问", desc: "阿里版ChatGPT" },
  { name: "文心一言", desc: "百度版ChatGPT" }
];

aiProducts.forEach((product, i) => {
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 0.5 + i * 2.35, y: 4.3, w: 2.15, h: 1.1,
    fill: { color: colors.white },
    shadow: makeShadow()
  });
  slide4.addText(product.name, {
    x: 0.5 + i * 2.35, y: 4.35, w: 2.15, h: 0.5,
    fontSize: 16, fontFace: "Microsoft YaHei",
    color: colors.secondary, bold: true, align: "center"
  });
  slide4.addText(product.desc, {
    x: 0.5 + i * 2.35, y: 4.85, w: 2.15, h: 0.5,
    fontSize: 12, fontFace: "Microsoft YaHei",
    color: colors.gray, align: "center"
  });
});

// ========== 第5页：AI vs Agent ==========
let slide5 = pres.addSlide();
slide5.background = { color: colors.white };

// 左侧装饰
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.3, h: 5.625,
  fill: { color: colors.accent }
});

slide5.addText("AI 和 Agent 有什么区别？", {
  x: 0.8, y: 0.3, w: 8.5, h: 0.7,
  fontSize: 30, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// 对比表格
const tableData = [
  [
    { text: "", options: { fill: { color: colors.lightGray } } },
    { text: "AI（人工智能）", options: { fill: { color: colors.secondary }, color: colors.white, bold: true, align: "center" } },
    { text: "Agent（智能代理）", options: { fill: { color: colors.primary }, color: colors.white, bold: true, align: "center" } }
  ],
  [
    { text: "像什么", options: { fill: { color: colors.light }, bold: true, align: "center" } },
    { text: "聪明的 人", options: { fill: { color: colors.light }, align: "center" } },
    { text: "能干的 员工", options: { fill: { color: colors.light }, align: "center" } }
  ],
  [
    { text: "会做什么", options: { fill: { color: colors.white }, bold: true, align: "center" } },
    { text: "回答问题、写东西", options: { fill: { color: colors.white }, align: "center" } },
    { text: "帮你 执行任务", options: { fill: { color: colors.white }, align: "center" } }
  ],
  [
    { text: "例子", options: { fill: { color: colors.light }, bold: true, align: "center" } },
    { text: '你问"怎么做红烧肉"，它告诉你', options: { fill: { color: colors.light }, align: "center" } },
    { text: '你说"帮我做红烧肉"，它帮你做', options: { fill: { color: colors.light }, align: "center" } }
  ]
];

slide5.addTable(tableData, {
  x: 0.8, y: 1.2, w: 8.4,
  colW: [1.8, 3.3, 3.3],
  border: { pt: 1, color: colors.gray },
  fontFace: "Microsoft YaHei",
  fontSize: 14
});

// 核心总结
slide5.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 4.0, w: 8.4, h: 1.2,
  fill: { color: colors.primary }
});
slide5.addText([
  { text: '简单说：AI是"能说"，Agent是"能做"', options: { breakLine: true, fontSize: 22, bold: true } },
  { text: "", options: { breakLine: true } },
  { text: "你就把Agent理解成一个——你给它下命令，它帮你干活的那种员工", options: { fontSize: 16 } }
], {
  x: 0.8, y: 4.0, w: 8.4, h: 1.2,
  fontFace: "Microsoft YaHei",
  color: colors.white, align: "center", valign: "middle"
});

// ========== 第6页：什么是养龙虾 ==========
let slide6 = pres.addSlide();
slide6.background = { color: colors.lightGray };

// 顶部强调
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.12,
  fill: { color: colors.accent }
});

slide6.addText("什么是\"养龙虾\"？", {
  x: 0.5, y: 0.4, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// 错误认知
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.2, w: 4.3, h: 1.8,
  fill: { color: "FEE2E2" }
});
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.2, w: 0.08, h: 1.8,
  fill: { color: "EF4444" }
});
slide6.addText("❌ 错误理解", {
  x: 0.7, y: 1.25, w: 3.9, h: 0.4,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: "DC2626", bold: true
});
slide6.addText([
  { text: "养龙虾是某个公司的AI游戏", options: { breakLine: true } },
  { text: "养龙虾是OpenClaw推出的产品", options: {} }
], {
  x: 0.7, y: 1.7, w: 3.9, h: 1.2,
  fontSize: 14, fontFace: "Microsoft YaHei",
  color: colors.dark, paraSpaceAfter: 6
});

// 正确理解
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 1.2, w: 4.3, h: 1.8,
  fill: { color: "D1FAE5" }
});
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 1.2, w: 0.08, h: 1.8,
  fill: { color: "10B981" }
});
slide6.addText("✅ 正确理解", {
  x: 5.4, y: 1.25, w: 3.9, h: 0.4,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: "059669", bold: true
});
slide6.addText([
  { text: "养龙虾 = 学会使用OpenClaw这个工具", options: { breakLine: true } },
  { text: "OpenClaw = 开源AI Agent框架", options: { breakLine: true } },
  { text: "Logo是红色小龙虾，所以叫养龙虾", options: {} }
], {
  x: 5.4, y: 1.7, w: 3.9, h: 1.2,
  fontSize: 14, fontFace: "Microsoft YaHei",
  color: colors.dark, paraSpaceAfter: 6
});

// 核心信息
slide6.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 3.2, w: 9, h: 2.1,
  fill: { color: colors.primary }
});
slide6.addText("养龙虾的本质", {
  x: 0.5, y: 3.3, w: 9, h: 0.5,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.accent, bold: true, align: "center"
});
slide6.addText([
  { text: "学会用AI工具帮你干活", options: { breakLine: true, fontSize: 20, bold: true } },
  { text: "", options: { breakLine: true } },
  { text: "学会了，你就有一个24小时不睡觉、", options: { breakLine: true } },
  { text: "不要工资、不会辞职的员工", options: {} }
], {
  x: 0.5, y: 3.8, w: 9, h: 1.4,
  fontFace: "Microsoft YaHei",
  color: colors.white, align: "center", valign: "middle"
});

// ========== 第7页：从软件到Agent ==========
let slide7 = pres.addSlide();
slide7.background = { color: colors.white };

// 左侧装饰
slide7.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.3, h: 5.625,
  fill: { color: colors.secondary }
});

slide7.addText("从\"用软件\"到\"用Agent\"", {
  x: 0.8, y: 0.3, w: 8.5, h: 0.7,
  fontSize: 30, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// 左侧：软件时代
slide7.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 1.2, w: 4, h: 3.8,
  fill: { color: "FEE2E2" }
});
slide7.addText("软件时代的痛点", {
  x: 0.8, y: 1.2, w: 4, h: 0.6,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: "DC2626", bold: true, align: "center", valign: "middle"
});
slide7.addText([
  { text: "要学N个软件的操作", options: { bullet: true, breakLine: true } },
  { text: "每个软件功能只用10%", options: { bullet: true, breakLine: true } },
  { text: "人在软件之间来回切换", options: { bullet: true, breakLine: true } },
  { text: "重复劳动浪费大量时间", options: { bullet: true } }
], {
  x: 1.0, y: 1.9, w: 3.6, h: 2.8,
  fontSize: 15, fontFace: "Microsoft YaHei",
  color: colors.dark, paraSpaceAfter: 12
});

// 箭头
slide7.addText("→", {
  x: 4.5, y: 2.8, w: 1, h: 1,
  fontSize: 48, fontFace: "Arial",
  color: colors.accent, bold: true, align: "center", valign: "middle"
});

// 右侧：Agent时代
slide7.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 1.2, w: 4, h: 3.8,
  fill: { color: "D1FAE5" }
});
slide7.addText("Agent时代的到来", {
  x: 5.2, y: 1.2, w: 4, h: 0.6,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: "059669", bold: true, align: "center", valign: "middle"
});
slide7.addText([
  { text: "你描述需求", options: { bullet: true, breakLine: true } },
  { text: "Agent自动选择工具", options: { bullet: true, breakLine: true } },
  { text: "自动执行任务", options: { bullet: true, breakLine: true } },
  { text: "不需要学操作，只需要说话", options: { bullet: true, bold: true } }
], {
  x: 5.4, y: 1.9, w: 3.6, h: 2.8,
  fontSize: 15, fontFace: "Microsoft YaHei",
  color: colors.dark, paraSpaceAfter: 12
});

// ========== 第8页：为什么小企业需要Agent ==========
let slide8 = pres.addSlide();
slide8.background = { color: colors.lightGray };

// 顶部条
slide8.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.12,
  fill: { color: colors.primary }
});

slide8.addText("为什么小企业更需要Agent？", {
  x: 0.5, y: 0.4, w: 9, h: 0.7,
  fontSize: 30, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true, align: "center"
});

// 小企业困境
const challenges = [
  { icon: "👥", text: "人手有限", desc: "一个人当三个人用" },
  { icon: "💰", text: "预算有限", desc: "买不起企业级软件" },
  { icon: "⏰", text: "时间有限", desc: "重复工作占满时间" }
];

challenges.forEach((item, i) => {
  slide8.addShape(pres.shapes.RECTANGLE, {
    x: 0.5 + i * 3.1, y: 1.3, w: 2.9, h: 1.5,
    fill: { color: colors.white },
    shadow: makeShadow()
  });
  slide8.addText(item.icon, {
    x: 0.5 + i * 3.1, y: 1.35, w: 2.9, h: 0.6,
    fontSize: 28, align: "center"
  });
  slide8.addText(item.text, {
    x: 0.5 + i * 3.1, y: 1.9, w: 2.9, h: 0.4,
    fontSize: 16, fontFace: "Microsoft YaHei",
    color: colors.primary, bold: true, align: "center"
  });
  slide8.addText(item.desc, {
    x: 0.5 + i * 3.1, y: 2.3, w: 2.9, h: 0.4,
    fontSize: 12, fontFace: "Microsoft YaHei",
    color: colors.gray, align: "center"
  });
});

// 成本对比
slide8.addText("Agent的解决方案", {
  x: 0.5, y: 3.0, w: 9, h: 0.5,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

const costTable = [
  [
    { text: "你需要", options: { fill: { color: colors.secondary }, color: colors.white, bold: true, align: "center" } },
    { text: "传统方案", options: { fill: { color: colors.secondary }, color: colors.white, bold: true, align: "center" } },
    { text: "Agent方案", options: { fill: { color: colors.secondary }, color: colors.white, bold: true, align: "center" } }
  ],
  [
    { text: "写文案", options: { align: "center" } },
    { text: "5000/月", options: { align: "center" } },
    { text: "说需求", options: { align: "center", color: "059669", bold: true } }
  ],
  [
    { text: "做海报", options: { align: "center" } },
    { text: "6000/月", options: { align: "center" } },
    { text: "描述画面", options: { align: "center", color: "059669", bold: true } }
  ],
  [
    { text: "回客户", options: { align: "center" } },
    { text: "4000/月", options: { align: "center" } },
    { text: "Agent自动", options: { align: "center", color: "059669", bold: true } }
  ],
  [
    { text: "合计", options: { fill: { color: colors.light }, bold: true, align: "center" } },
    { text: "19000/月", options: { fill: { color: colors.light }, align: "center" } },
    { text: "0-100/月", options: { fill: { color: colors.light }, color: "059669", bold: true, align: "center" } }
  ]
];

slide8.addTable(costTable, {
  x: 0.5, y: 3.5, w: 9,
  colW: [3, 3, 3],
  border: { pt: 1, color: colors.gray },
  fontFace: "Microsoft YaHei",
  fontSize: 14
});

// ========== 第9页：Skill生态 ==========
let slide9 = pres.addSlide();
slide9.background = { color: colors.white };

// 左侧装饰
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.3, h: 5.625,
  fill: { color: colors.accent }
});

slide9.addText("Skill生态：Agent的超能力", {
  x: 0.8, y: 0.3, w: 8.5, h: 0.7,
  fontSize: 30, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// 类比说明
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 1.1, w: 8.4, h: 1.3,
  fill: { color: colors.light }
});
slide9.addText([
  { text: "Agent = 智能手机    ", options: { bold: true } },
  { text: "Skill = App", options: { bold: true, breakLine: true } },
  { text: "", options: { breakLine: true } },
  { text: '你说"我想导航" → Agent自动打开地图App', options: { breakLine: true } },
  { text: '你说"做张海报" → Agent自动调用海报Skill', options: {} }
], {
  x: 1.0, y: 1.2, w: 8, h: 1.1,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: colors.dark, align: "center"
});

// 常用Skill分类
slide9.addText("常用Skill一览", {
  x: 0.8, y: 2.6, w: 8.4, h: 0.5,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

// 企业办公Skill
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 3.2, w: 4, h: 2.1,
  fill: { color: colors.primary }
});
slide9.addText("企业办公Skill", {
  x: 0.8, y: 3.25, w: 4, h: 0.5,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center"
});
slide9.addText([
  { text: "飞书Skill：文档、多维表、日历", options: { bullet: true, breakLine: true } },
  { text: "钉钉Skill：AI表格、日历、待办", options: { bullet: true, breakLine: true } },
  { text: "企微Skill：消息、文档、智能表", options: { bullet: true } }
], {
  x: 1.0, y: 3.8, w: 3.6, h: 1.4,
  fontSize: 13, fontFace: "Microsoft YaHei",
  color: colors.white, paraSpaceAfter: 8
});

// 设计Skill
slide9.addShape(pres.shapes.RECTANGLE, {
  x: 5.2, y: 3.2, w: 4, h: 2.1,
  fill: { color: colors.accent }
});
slide9.addText("设计Skill", {
  x: 5.2, y: 3.25, w: 4, h: 0.5,
  fontSize: 16, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center"
});
slide9.addText([
  { text: "海报生成：输入需求，直接出图", options: { bullet: true, breakLine: true } },
  { text: "文案生成：写各种文案", options: { bullet: true, breakLine: true } },
  { text: "图像生成：生成产品图", options: { bullet: true } }
], {
  x: 5.4, y: 3.8, w: 3.6, h: 1.4,
  fontSize: 13, fontFace: "Microsoft YaHei",
  color: colors.white, paraSpaceAfter: 8
});

// ========== 第10页：RTGF指令公式 ==========
let slide10 = pres.addSlide();
slide10.background = { color: colors.lightGray };

// 顶部条
slide10.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.12,
  fill: { color: colors.secondary }
});

slide10.addText("RTGF指令公式", {
  x: 0.5, y: 0.4, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true, align: "center"
});

slide10.addText("让Agent听懂你的需求", {
  x: 0.5, y: 1.0, w: 9, h: 0.4,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.gray, align: "center"
});

// RTGF卡片
const rtgfItems = [
  { letter: "R", name: "Role", desc: "你希望Agent扮演什么？", example: "餐饮营销专家" },
  { letter: "T", name: "Task", desc: "具体要做什么？", example: "写3条促销文案" },
  { letter: "G", name: "Goal", desc: "要达到什么效果？", example: "吸引年轻上班族" },
  { letter: "F", name: "Format", desc: "以什么形式输出？", example: "每条配3个标签" }
];

rtgfItems.forEach((item, i) => {
  const yPos = 1.6 + i * 0.95;

  // 字母圆圈
  slide10.addShape(pres.shapes.OVAL, {
    x: 0.8, y: yPos, w: 0.7, h: 0.7,
    fill: { color: colors.secondary }
  });
  slide10.addText(item.letter, {
    x: 0.8, y: yPos, w: 0.7, h: 0.7,
    fontSize: 24, fontFace: "Arial",
    color: colors.white, bold: true, align: "center", valign: "middle"
  });

  // 内容
  slide10.addText(item.name, {
    x: 1.7, y: yPos, w: 1.5, h: 0.7,
    fontSize: 18, fontFace: "Arial",
    color: colors.primary, bold: true, valign: "middle"
  });
  slide10.addText(item.desc, {
    x: 3.2, y: yPos, w: 3, h: 0.7,
    fontSize: 15, fontFace: "Microsoft YaHei",
    color: colors.dark, valign: "middle"
  });
  slide10.addText(item.example, {
    x: 6.2, y: yPos, w: 3, h: 0.7,
    fontSize: 14, fontFace: "Microsoft YaHei",
    color: colors.accent, italic: true, valign: "middle"
  });
});

// ========== 第11页：各行业应用 ==========
let slide11 = pres.addSlide();
slide11.background = { color: colors.white };

// 左侧装饰
slide11.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.3, h: 5.625,
  fill: { color: colors.primary }
});

slide11.addText("各行业AI应用思路", {
  x: 0.8, y: 0.3, w: 8.5, h: 0.7,
  fontSize: 30, fontFace: "Microsoft YaHei",
  color: colors.primary, bold: true
});

const industries = [
  { name: "餐饮", app: "客服 + 内容生成" },
  { name: "零售", app: "客服 + 数据分析" },
  { name: "美业", app: "预约 + 案例展示" },
  { name: "培训", app: "课件 + 家长沟通" },
  { name: "服务", app: "客服 + 日程管理" }
];

industries.forEach((ind, i) => {
  const yPos = 1.2 + i * 0.85;

  slide11.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: yPos, w: 8.4, h: 0.7,
    fill: { color: i % 2 === 0 ? colors.light : colors.white }
  });
  slide11.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: yPos, w: 0.08, h: 0.7,
    fill: { color: colors.accent }
  });
  slide11.addText(ind.name, {
    x: 1.1, y: yPos, w: 1.5, h: 0.7,
    fontSize: 18, fontFace: "Microsoft YaHei",
    color: colors.primary, bold: true, valign: "middle"
  });
  slide11.addText(ind.app, {
    x: 2.8, y: yPos, w: 6, h: 0.7,
    fontSize: 16, fontFace: "Microsoft YaHei",
    color: colors.dark, valign: "middle"
  });
});

// ========== 第12页：课程总结 ==========
let slide12 = pres.addSlide();
slide12.background = { color: colors.primary };

// 装饰
slide12.addShape(pres.shapes.OVAL, {
  x: -2, y: -2, w: 5, h: 5,
  fill: { color: colors.secondary, transparency: 50 }
});
slide12.addShape(pres.shapes.OVAL, {
  x: 7.5, y: 3, w: 4, h: 4,
  fill: { color: colors.accent, transparency: 40 }
});

slide12.addText("课程总结", {
  x: 0.5, y: 0.5, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center"
});

// 一句话总结
slide12.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1.4, w: 8, h: 1.0,
  fill: { color: colors.white, transparency: 20 }
});
slide12.addText("Agent时代，你只需要说话，AI帮你搞定一切", {
  x: 1, y: 1.4, w: 8, h: 1.0,
  fontSize: 22, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center", valign: "middle"
});

// 三个关键带走
slide12.addText("3个关键带走", {
  x: 0.5, y: 2.6, w: 9, h: 0.5,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.accent, bold: true, align: "center"
});

const takeaways = [
  { num: "1", title: "RTGF指令公式", desc: "让Agent听懂你的需求" },
  { num: "2", title: "Skill生态", desc: "Agent背后的超能力库" },
  { num: "3", title: "迭代优化", desc: "和Agent协作，越用越顺手" }
];

takeaways.forEach((item, i) => {
  slide12.addShape(pres.shapes.OVAL, {
    x: 1.2 + i * 2.8, y: 3.2, w: 0.6, h: 0.6,
    fill: { color: colors.accent }
  });
  slide12.addText(item.num, {
    x: 1.2 + i * 2.8, y: 3.2, w: 0.6, h: 0.6,
    fontSize: 20, fontFace: "Arial",
    color: colors.white, bold: true, align: "center", valign: "middle"
  });
  slide12.addText(item.title, {
    x: 0.8 + i * 2.8, y: 3.9, w: 2.4, h: 0.5,
    fontSize: 16, fontFace: "Microsoft YaHei",
    color: colors.white, bold: true, align: "center"
  });
  slide12.addText(item.desc, {
    x: 0.8 + i * 2.8, y: 4.4, w: 2.4, h: 0.4,
    fontSize: 12, fontFace: "Microsoft YaHei",
    color: colors.white, align: "center"
  });
});

// ========== 第13页：结束页 ==========
let slide13 = pres.addSlide();
slide13.background = { color: colors.primary };

// 大装饰圆
slide13.addShape(pres.shapes.OVAL, {
  x: 3, y: 0.5, w: 4, h: 4,
  fill: { color: colors.secondary, transparency: 40 }
});

slide13.addText("🎁", {
  x: 0.5, y: 1.5, w: 9, h: 1,
  fontSize: 60, align: "center"
});

slide13.addText("今日行动", {
  x: 0.5, y: 2.6, w: 9, h: 0.7,
  fontSize: 36, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center"
});

slide13.addText([
  { text: "本周内，每天用Agent完成至少1个任务", options: { breakLine: true } },
  { text: "在学员群分享你的指令和成果", options: { breakLine: true } },
  { text: "遇到问题随时提问", options: {} }
], {
  x: 1.5, y: 3.4, w: 7, h: 1.2,
  fontSize: 18, fontFace: "Microsoft YaHei",
  color: colors.white, align: "center", paraSpaceAfter: 10
});

// 结束语
slide13.addShape(pres.shapes.RECTANGLE, {
  x: 1.5, y: 4.7, w: 7, h: 0.7,
  fill: { color: colors.accent }
});
slide13.addText("Agent不是替代你，是让你更强大", {
  x: 1.5, y: 4.7, w: 7, h: 0.7,
  fontSize: 20, fontFace: "Microsoft YaHei",
  color: colors.white, bold: true, align: "center", valign: "middle"
});

// 保存文件
pres.writeFile({ fileName: "小企业AI入门实战课.pptx" })
  .then(() => console.log("PPT已生成：小企业AI入门实战课.pptx"))
  .catch(err => console.error("生成失败：", err));