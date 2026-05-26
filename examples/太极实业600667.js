const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
        ShadingType, PageNumber, LevelFormat, PageBreak } = require("docx");
const fs = require("fs");

// Helpers
const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 }, children: [new TextRun({ text, bold: true, size: 32, font: "Microsoft YaHei" })] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 }, children: [new TextRun({ text, bold: true, size: 28, font: "Microsoft YaHei" })] });
const h3 = (text) => new Paragraph({ spacing: { before: 200, after: 120 }, children: [new TextRun({ text, bold: true, size: 24, font: "Microsoft YaHei" })] });
const p = (text, opts = {}) => new Paragraph({ spacing: { before: 80, after: 80, line: 360 }, children: [new TextRun({ text, size: 22, font: "Microsoft YaHei", ...opts })] });
const pBold = (label, value) => new Paragraph({ spacing: { before: 60, after: 60, line: 360 }, children: [new TextRun({ text: label, bold: true, size: 22, font: "Microsoft YaHei" }), new TextRun({ text: value, size: 22, font: "Microsoft YaHei" })] });
const pHighlight = (text) => new Paragraph({ spacing: { before: 120, after: 120, line: 360 }, shading: { fill: "FFF3CD", type: ShadingType.CLEAR }, children: [new TextRun({ text, size: 22, font: "Microsoft YaHei", bold: true })] });

// Table helper
const borderDef = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const allBorders = { top: borderDef, bottom: borderDef, left: borderDef, right: borderDef };

function makeTable(headers, rows, colWidths) {
  const tableWidth = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: allBorders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: "2B5797", type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 20, font: "Microsoft YaHei", color: "FFFFFF" })] })]
    }))
  });
  const dataRows = rows.map(row => new TableRow({
    children: row.map((cell, i) => new TableCell({
      borders: allBorders,
      width: { size: colWidths[i], type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ children: [new TextRun({ text: String(cell), size: 20, font: "Microsoft YaHei" })] })]
    }))
  }));
  return new Table({ width: { size: tableWidth, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
}

const children = [
  // Cover
  new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "个股调研报告", size: 44, bold: true, font: "Microsoft YaHei", color: "2B5797" })] }),
  new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "太极实业（600667.SH）", size: 36, bold: true, font: "Microsoft YaHei" })] }),
  new Paragraph({ spacing: { before: 200 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "半导体工程龙头 + 存储封测双主业", size: 26, font: "Microsoft YaHei", color: "666666" })] }),
  new Paragraph({ spacing: { before: 600 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "报告日期：2026年5月26日", size: 22, font: "Microsoft YaHei", color: "888888" })] }),
  new Paragraph({ spacing: { before: 100 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "数据截止：2026年一季报 + 5月26日实时行情", size: 22, font: "Microsoft YaHei", color: "888888" })] }),

  new Paragraph({ children: [new PageBreak()] }),

  // 一、公司概况
  h1("一、公司概况"),
  pBold("全称：", "无锡市太极实业股份有限公司"),
  pBold("股票代码：", "600667.SH（上海证券交易所）"),
  pBold("上市时间：", "1993年（A股老牌上市公司，上市超30年）"),
  pBold("实控人：", "中国电子信息产业集团（央企控股）"),
  pBold("核心子公司：", "十一科技（工程设计总承包）、海太半导体（DRAM封测，与SK海力士合资）"),
  pBold("总部：", "江苏无锡"),
  pBold("当前市值：", "约258-277亿元"),
  pBold("最新股价：", "约12.6-13.7元（2026年5月下旬）"),
  p(""),
  p("公司经过50余年发展和转型升级，已从传统化纤企业转型为国内领先的半导体（集成电路）制造与服务厂商。当前核心战略围绕两大业务板块展开：电子高科技工程技术服务（通过子公司十一科技运营）和半导体制造与封测服务（通过海太半导体和太极半导体运营）。"),

  // 二、产业布局
  h1("二、产业布局与业务结构"),

  h2("2.1 电子高科技工程技术服务（营收占比约76.5%）"),
  p("通过全资子公司十一科技运营，是公司营收的绝对主力。"),
  h3("核心能力："),
  p("从咨询、规划、设计到工程总承包（EPC）的全链条服务能力，聚焦电子高科技产业厂房建设，尤其擅长半导体晶圆厂洁净室工程设计与施工。"),
  h3("服务领域："),
  p("集成电路/芯片制造厂房（晶圆厂FAB）、液晶显示面板厂房、新能源（光伏/锂电）生产线、生物医药洁净车间等高科技产业工厂。"),
  h3("核心客户："),
  p("中芯国际、长江存储、长鑫存储（合肥睿力）、华虹半导体、SK海力士等国内外主流芯片制造企业。"),
  h3("业务特征——一次性收入模型："),
  p("项目制、建成交付后不再产生后续收入。单个合同金额大（数亿至数十亿），但公司持续增长完全依赖获取新建/扩产项目订单。维护类收入极低（洁净室运维一般由专业FM公司承接）。"),

  h2("2.2 半导体制造与封测服务（营收占比约15.15%）"),
  p("2025年度半导体业务完成营业收入46.50亿元。"),
  h3("海太半导体（控股子公司，与SK海力士合资）："),
  p("主要为SK海力士DRAM产品提供后工序封装测试服务。技术覆盖TSOP、QFP、BOC、BGA等封装形式。高度依赖SK海力士单一大客户。"),
  h3("太极半导体："),
  p("自主开拓市场的封测平台，曾中标合肥睿力（长鑫存储）封测业务。"),
  h3("业务特征——持续性收入但占比小："),
  p("相比工程服务，封测业务可产生持续性收入（客户持续出货=持续提供封测服务），但当前体量仅占15%，尚不足以对冲工程业务的周期波动。"),

  // 三、行业地位
  h1("三、行业地位与竞争格局"),

  h2("3.1 半导体工程设计——国内绝对龙头"),
  p("十一科技在半导体洁净室工程领域是国内的绝对龙头，在半导体晶圆厂设计领域占据国内大部分市场份额。"),

  makeTable(
    ["公司", "优势领域", "市场地位"],
    [
      ["十一科技（太极实业）", "半导体晶圆厂洁净室工程", "国内第一，绝对龙头"],
      ["柏诚股份（601133）", "半导体/面板洁净室工程", "第二梯队，中高端竞争力强"],
      ["中电二/中电四", "显示面板洁净室工程", "面板领域强，半导体偏弱"],
      ["世源科技", "洁净室分包", "细分领域参与者"],
    ],
    [2200, 3500, 3660]
  ),

  p(""),
  p("核心壁垒：半导体晶圆厂对洁净度要求极高（Class 1-10级），设计和施工经验积累需数十年，客户切换成本高。十一科技凭借央企背景+数十年项目积累+全链条能力，护城河深厚。中国几乎所有主流芯片工厂的厂房和产线设计，基本都绑定十一科技。"),

  h2("3.2 半导体封测——第二梯队，非主流玩家"),

  makeTable(
    ["公司", "2025年封测营收", "PE（2026.5月）", "行业地位"],
    [
      ["长电科技（600584）", "388.71亿（全公司）", "约87.7倍", "国内第一、全球第三，先进封装270亿"],
      ["通富微电（002156）", "约250亿+", "约94倍", "国内第二，AMD代工绑定"],
      ["华天科技（002185）", "约130亿+", "--", "国内第三，市值504亿"],
      ["太极实业（海太）", "46.5亿（半导体板块）", "约58-62倍", "第二梯队，依赖SK海力士"],
    ],
    [2200, 2400, 2000, 2760]
  ),

  p(""),
  p("在先进封装（CoWoS、Chiplet、Fan-out等）领域，太极实业竞争力较弱。长电科技先进封装营收已达270亿元，通富微电拟定增加码先进封装，而太极实业在该领域缺乏布局。"),

  // 四、财务分析
  h1("四、核心财务数据"),

  h2("4.1 近期业绩概览"),

  makeTable(
    ["指标", "2025年全年", "2026年Q1", "说明"],
    [
      ["营业收入", "306.82亿元", "66.58亿元（同比-0.91%）", "营收规模大但增速停滞"],
      ["归母净利润", "4.48亿元", "1.29亿元（同比+9.48%）", "利润端微增但绝对值低"],
      ["综合毛利率", "12.79%", "--", "工程+制造双低毛利"],
      ["加权平均ROE", "5.22%", "--", "资本回报率偏低"],
      ["每股净资产", "4.18元", "--", ""],
      ["每股收益EPS", "约0.21元/年", "0.06元/季", "年化约0.24元"],
    ],
    [2000, 2600, 2800, 1960]
  ),

  p(""),
  pHighlight("重要提示：2023年券商研报曾预测2025年归母净利润13.63亿元，实际仅完成4.48亿元，仅为预期的33%。业绩大幅低于预期。"),

  h2("4.2 收入结构"),

  makeTable(
    ["业务板块", "2025年收入（亿元）", "占比", "收入特征"],
    [
      ["电子高科技工程服务", "约234.6", "76.5%", "一次性项目收入，建完即止"],
      ["半导体制造与封测", "46.5", "15.15%", "持续性收入，强依赖SK海力士"],
      ["其他", "约25.7", "8.35%", "化纤遗留+贸易等"],
    ],
    [2600, 2400, 1200, 3160]
  ),

  // 五、估值分析
  h1("五、估值分析"),

  h2("5.1 当前估值水平（2026年5月）"),

  makeTable(
    ["指标", "太极实业", "说明"],
    [
      ["最新股价", "12.6-13.7元", "5月12日涨停后高位震荡"],
      ["总市值", "约258-277亿元", ""],
      ["PE（TTM）", "约58-62倍", "以2025年4.48亿净利润计"],
      ["PE（中位）", "约25.9倍", "中财网数据"],
      ["PB", "约2.8倍", "每股净资产4.18元"],
      ["机构目标价", "10.97元（平均）", "低于当前股价，券商认为高估"],
    ],
    [2200, 2800, 4360]
  ),

  h2("5.2 同行业估值对比（2026年5月最新）"),
  pHighlight("当前整个半导体板块处于高估值状态，并非太极实业独有现象。"),
  p(""),

  makeTable(
    ["公司", "市值（亿）", "PE（TTM）", "PB", "2025归母净利润"],
    [
      ["太极实业（600667）", "~270", "~58-62倍", "~2.8", "4.48亿"],
      ["柏诚股份（601133）", "~180", "~85-91倍", "~5.4-5.8", "~2亿"],
      ["长电科技（600584）", "~1304", "~87.7倍", "~4.63", "15.65亿"],
      ["通富微电（002156）", "~929", "~94倍", "~6.12", "~10亿"],
      ["华天科技（002185）", "~504", "--", "~2.81", "--"],
    ],
    [2200, 1600, 1800, 1400, 2360]
  ),

  p(""),
  p("关键发现：太极实业PE 58-62倍，在当前半导体板块中反而是相对最低的。长电PE 88倍、通富PE 94倍、柏诚PE 85-91倍。整个板块都在博弈AI算力扩产带来的业绩爆发预期。"),

  h2("5.3 估值合理性判断"),
  h3("结论：板块整体高估值环境下，太极实业估值相对同行不算最贵，但绝对估值仍然偏高。"),
  p(""),
  p("具体分析："),
  p("1. 绝对估值角度：PE 58-62倍，对于一家ROE仅5.22%、毛利率12.79%、2025年业绩远不及预期的工程+制造公司而言，估值偏高。传统建筑工程公司通常PE在10-20倍。"),
  p("2. 相对估值角度：在半导体封测板块中（长电88x、通富94x、柏诚85-91x），太极的58-62倍反而是最低的。这说明市场给太极的定价较为克制，隐含了其工程属性（低毛利、周期性）的折价。"),
  p("3. 机构观点：平均目标价10.97元低于当前12-14元股价，表明券商整体认为短期估值偏贵。但需注意部分研报可能未更新（半导体板块近两月普涨）。"),
  p("4. 历史中枢：太极实业历史PE中位数约25.9倍（中财网数据），当前58-62倍已处于历史高位区域。"),

  // 六、情绪面与资金面分析（新增）
  new Paragraph({ children: [new PageBreak()] }),
  h1("六、情绪面与资金面分析"),

  h2("6.1 近期异动事件"),
  p("2026年5月12日：太极实业涨停（+10.04%），收盘价12.27元，全天换手率15.62%，成交额38.01亿元。"),
  p("2026年5月13日：连续3个交易日涨幅偏离值累计达20%，触发股票交易异常波动公告。"),
  p("涨停原因：存储涨价预期+SK海力士长单利好+半导体板块整体情绪带动。"),

  h2("6.2 资金面信号"),
  p("5日主力资金净流出约5亿元（东方财富数据），说明涨停后主力有兑现行为。"),
  p("融资融券余额9.90亿元（5月11日数据），其中融资余额9.82亿元，融资盘占比高，杠杆资金参与度大。"),
  p("龙虎榜显示5月12日涨停日有机构和游资共同参与，但5月13日异动后主力净卖出4265万元。"),

  h2("6.3 市场情绪驱动逻辑"),
  p("当前太极实业被市场定价的核心逻辑不是基本面（4.48亿利润撑不起270亿市值），而是以下情绪/主题驱动："),
  p(""),
  p("1. 半导体国产替代主线：2026年AI算力大爆发+美国制裁升级=国产芯片扩产加速，太极作为建厂必经环节的卖铲人，确定性受益。"),
  p("2. AI算力基建概念：英伟达816亿营收验证AI算力需求暴增，国内算力建设追赶=更多芯片厂需要建设=十一科技订单预期。"),
  p("3. 存储涨价周期：SK海力士业绩爆炸（2026Q1净利280亿美元），HBM供不应求，存储扩产=封测需求增加=海太半导体受益。"),
  p("4. 央企改革预期：中国电子集团旗下，存在资产整合和注入的想象空间。"),
  p("5. 板块beta效应：整个半导体封测板块估值集体抬升（长电88x、通富94x），太极作为板块成员被动提估。"),

  h2("6.4 情绪面风险提示"),
  p("当前股价更多是情绪和预期驱动，而非业绩支撑。一旦出现以下情况，估值可能快速回落："),
  p("- 半导体板块情绪退潮（如大盘回调、资金从科技转向其他方向）"),
  p("- 芯片扩产项目延期或取消（如制裁加码导致设备无法到位）"),
  p("- 公司后续季度业绩持续平淡（Q2/Q3不及预期将加速估值回归）"),
  p("- SK海力士订单调整或合资关系变化"),
  p(""),
  pHighlight("情绪结论：当前是典型的\"主题驱动+板块beta\"行情。适合趋势交易者参与，但基本面投资者需警惕估值回归风险。"),

  // 七、核心看点与风险
  h1("七、核心看点与风险"),

  h2("7.1 看多逻辑"),
  p("1. 半导体洁净室工程绝对龙头：国内芯片建厂几乎必经十一科技，护城河深厚。"),
  p("2. 中国半导体扩产大周期：自主可控+大基金三期+地方投资，芯片建厂潮远未结束。"),
  p("3. AI算力建设间接拉动：GPU/HBM/存储芯片扩产=更多晶圆厂建设需求。"),
  p("4. 存储涨价周期利好封测：SK海力士/三星扩产HBM，海太封测订单有望增长。"),
  p("5. 央企整合预期：中国电子集团旗下资产整合的想象空间。"),
  p("6. 板块相对估值不贵：PE 58-62倍在半导体封测板块中反而最低（对比长电88x、通富94x）。"),

  h2("7.2 看空/风险逻辑"),
  p("1. 工程业务一次性特征：建完即止，天然有天花板，增长依赖新项目。"),
  p("2. 2025年业绩远不及预期：实际利润仅为券商预期的33%，证伪力度大。"),
  p("3. 绝对估值偏高：PE 58-62倍 vs 5.22%的ROE和12.79%的毛利率，不匹配。"),
  p("4. 主力资金净流出：涨停后5日主力净卖出约5亿，短期获利盘压力大。"),
  p("5. 周期性风险：扩产潮退潮则订单断崖，工程公司没有\"存量收入\"兜底。"),
  p("6. 先进封装缺位：在CoWoS等先进封装领域缺乏布局，缺少长期增长第二曲线。"),
  p("7. SK海力士依赖：封测业务高度依赖单一客户，合资关系变动风险。"),

  // 八、综合评价
  h1("八、综合评价"),
  p("太极实业的本质定位是\"半导体基建公司\"——帮芯片厂建厂房的卖铲人，兼做一些存储封测。在工程设计领域有绝对龙头地位，但业务模式（一次性项目制+低毛利+低ROE）限制了估值天花板。"),
  p(""),
  p("当前270亿市值、PE 58-62倍的定价，核心是半导体国产替代+AI算力基建的情绪溢价，而非业绩支撑（4.48亿利润）。在整个封测板块估值集体抬升（长电88x、通富94x、柏诚85-91x）的背景下，太极的相对估值不算最贵，但绝对估值仍处于历史高位。"),
  p(""),
  p("投资判断取决于交易风格：趋势交易者可关注板块情绪和资金动向，基本面投资者则需等待业绩拐点确认或估值回落至合理区间。机构平均目标价10.97元（对应PE约52倍）暗示当前存在10-20%的高估。"),

  // 免责声明
  new Paragraph({ children: [new PageBreak()] }),
  h2("免责声明"),
  p("本报告仅为个人研究记录，不构成任何投资建议。所有数据来源于公开信息，已尽力核实但不保证完全准确。投资有风险，入市需谨慎。"),
  p(""),
  p("数据来源：新浪财经、东方财富、同花顺、中财网、搜狐证券、财联社、亿牛网、理杏仁、证券时报、21世纪经济报道、凤凰财经、公司2025年年报及2026年一季报、券商研报（国金证券、粤开证券等）。"),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Microsoft YaHei", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "太极实业（600667.SH）调研报告 | 2026.05.26", size: 18, color: "999999", font: "Microsoft YaHei" })] })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 18, color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "999999" })] })] })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/Users/dfq/.qoderwork/workspace/mn4da9d3aqvrxwng/outputs/太极实业600667调研报告_v2.docx", buf);
  console.log("Done! Size:", buf.length, "bytes");
});
