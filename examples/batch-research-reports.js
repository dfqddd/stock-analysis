const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, Header, Footer,
  PageNumber, NumberFormat, ShadingType, HeadingLevel,
  TableLayoutType, VerticalAlign, PageBreak, Tab
} = require("docx");
const fs = require("fs");
const path = require("path");

const OUTPUT_DIR = path.join(__dirname, "..", "outputs");

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const TODAY = "2026-05-28";

// Common styles
const FONT = "Microsoft YaHei";
const BODY_SIZE = 22;   // 11pt in half-points
const H1_SIZE = 32;     // 16pt
const H2_SIZE = 28;     // 14pt
const TABLE_HEADER_BG = "2B5797";
const TABLE_HEADER_TEXT = "FFFFFF";
const BORDER_COLOR = "BFBFBF";
const HIGHLIGHT_BG = "FFFF00";

// Helper: create a text run
function text(content, options = {}) {
  return new TextRun({
    text: content,
    font: FONT,
    size: options.size || BODY_SIZE,
    bold: options.bold || false,
    color: options.color || "000000",
    ...options
  });
}

// Helper: create a paragraph
function para(content, options = {}) {
  const runs = Array.isArray(content) ? content : [text(content, options)];
  return new Paragraph({
    children: runs,
    spacing: { after: 120, line: 360 },
    alignment: options.alignment || AlignmentType.LEFT,
    ...( options.shading ? { shading: { type: ShadingType.CLEAR, fill: HIGHLIGHT_BG } } : {}),
    ...(options.paraOptions || {})
  });
}

// Helper: H1 heading
function h1(content) {
  return new Paragraph({
    children: [text(content, { size: H1_SIZE, bold: true })],
    spacing: { before: 240, after: 160, line: 360 },
    alignment: AlignmentType.LEFT
  });
}

// Helper: H2 heading
function h2(content) {
  return new Paragraph({
    children: [text(content, { size: H2_SIZE, bold: true })],
    spacing: { before: 200, after: 120, line: 360 },
    alignment: AlignmentType.LEFT
  });
}

// Helper: bullet point
function bullet(content) {
  return para(`  \u2022  ${content}`);
}

// Helper: table cell border
const cellBorder = {
  top: { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR },
  left: { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR },
  right: { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR }
};

// Helper: create table header cell
function headerCell(content, width) {
  return new TableCell({
    children: [new Paragraph({
      children: [text(content, { bold: true, color: TABLE_HEADER_TEXT, size: BODY_SIZE })],
      alignment: AlignmentType.CENTER
    })],
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: TABLE_HEADER_BG },
    borders: cellBorder,
    verticalAlign: VerticalAlign.CENTER
  });
}

// Helper: create table body cell
function bodyCell(content, width) {
  return new TableCell({
    children: [new Paragraph({
      children: [text(content, { size: BODY_SIZE })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 40, after: 40 }
    })],
    width: { size: width, type: WidthType.DXA },
    borders: cellBorder,
    verticalAlign: VerticalAlign.CENTER
  });
}

// Helper: create a 2-column table (key-value)
function kvTable(rows, col1Width = 2800, col2Width = 6400) {
  const tableRows = [];
  tableRows.push(new TableRow({
    children: [
      headerCell("项目", col1Width),
      headerCell("内容", col2Width)
    ]
  }));
  for (const [key, value] of rows) {
    tableRows.push(new TableRow({
      children: [
        bodyCell(key, col1Width),
        bodyCell(value, col2Width)
      ]
    }));
  }
  return new Table({
    rows: tableRows,
    width: { size: col1Width + col2Width, type: WidthType.DXA },
    columnWidths: [col1Width, col2Width],
    layout: TableLayoutType.FIXED
  });
}

// Helper: create multi-column table
function multiColTable(headers, rows, widths) {
  const totalWidth = widths.reduce((a, b) => a + b, 0);
  const tableRows = [];
  tableRows.push(new TableRow({
    children: headers.map((h, i) => headerCell(h, widths[i]))
  }));
  for (const row of rows) {
    tableRows.push(new TableRow({
      children: row.map((cell, i) => bodyCell(cell, widths[i]))
    }));
  }
  return new Table({
    rows: tableRows,
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED
  });
}

// Helper: highlighted paragraph
function highlight(content) {
  return new Paragraph({
    children: [text(content, { size: BODY_SIZE, bold: true })],
    spacing: { before: 120, after: 120, line: 360 },
    shading: { type: ShadingType.CLEAR, fill: HIGHLIGHT_BG }
  });
}

// Helper: create cover page
function coverPage(title, company, oneLiner, date) {
  return [
    new Paragraph({ spacing: { before: 3000 } }),
    new Paragraph({
      children: [text(title, { size: 52, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 }
    }),
    new Paragraph({
      children: [text(company, { size: 36, bold: true, color: "2B5797" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 }
    }),
    new Paragraph({
      children: [text(oneLiner, { size: 24 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 }
    }),
    new Paragraph({
      children: [text(date, { size: 24, color: "666666" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 }
    }),
    new Paragraph({
      children: [new PageBreak()]
    })
  ];
}

// Helper: disclaimer
function disclaimer() {
  return [
    h1("免责声明"),
    para("本报告仅供参考，不构成任何投资建议。报告中的信息来源于公开资料，作者不对其准确性、完整性做出任何保证。"),
    para("投资有风险，入市需谨慎。过往业绩不代表未来表现。读者应根据自身情况独立判断，并承担相应的投资风险。"),
    para("本报告版权归作者所有，未经许可不得转载、复制或用于商业用途。")
  ];
}

// Helper: create document
function createDoc(companyName, sections) {
  return new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: BODY_SIZE }
        }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } // 1 inch = 1440 DXA
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [
              text(`${companyName}`, { size: 18, color: "666666" }),
              text(`    ${TODAY}`, { size: 18, color: "666666" })
            ],
            alignment: AlignmentType.RIGHT
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              text("第 ", { size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "666666" }),
              text(" 页", { size: 18, color: "666666" })
            ],
            alignment: AlignmentType.CENTER
          })]
        })
      },
      children: sections
    }]
  });
}

// ==================== Report 1: 三花智控 ====================
function report1() {
  const sections = [
    ...coverPage("个股调研报告", "三花智控（002050.SZ）", "全球制冷阀件+汽车热管理龙头，机器人第三曲线", TODAY),
    
    h1("一、公司概况"),
    kvTable([
      ["全称", "浙江三花智能控制股份有限公司"],
      ["代码", "002050.SZ"],
      ["上市时间", "2005年"],
      ["实控人", "张道才、张亚波、张少波（父子）"],
      ["子公司", "三花汽车零部件、三花商用制冷、三花国际（新加坡/美国）"],
      ["总部", "浙江绍兴新昌"],
      ["市值", "约2200-2270亿元"],
      ["股价", "约52-54元"]
    ]),

    h1("二、产业布局"),
    h2("1. 制冷空调零部件"),
    para("营收185.85亿，占比60%，同比+12.22%"),
    para("产品：电子膨胀阀、截止阀、四通换向阀等，全球制冷阀件龙头，与盾安环境形成\"双寡头\"格局。"),
    h2("2. 汽车热管理"),
    para("营收124.27亿，占比40%，同比+9.14%"),
    para("汽车电子膨胀阀全球市占率52-53%排名第一，客户含特斯拉/比亚迪/大众/宝马。"),
    h2("3. 人形机器人执行器（新兴）"),
    para("旋转/线性执行器，绑定特斯拉Optimus，已获样品订单。"),
    h2("4. 数据中心液冷（新方向）"),
    para("AI算力液冷温控方案。"),

    h1("三、行业地位"),
    multiColTable(
      ["指标", "详情"],
      [
        ["全球汽车电子膨胀阀", "市占率52-53%，全球第一"],
        ["全球制冷阀件", "与盾安环境双寡头，海外对手不二工机（19%）"],
        ["竞争对手", "盾安环境（~200亿市值，PE 25-30x）"],
        ["竞争对手", "拓普集团（~1180亿，35-40x）"],
        ["竞争对手", "银轮股份（~300亿，30-35x）"]
      ],
      [3200, 6000]
    ),

    h1("四、财务数据"),
    highlight("2025全年业绩"),
    multiColTable(
      ["指标", "数值", "增速"],
      [
        ["营收", "310.12亿", "+10.97%"],
        ["归母净利", "40.63亿", "+31.10%"],
        ["毛利率", "28.78%", "-"],
        ["ROE", "约15.80%", "-"]
      ],
      [2600, 3200, 3400]
    ),
    para(""),
    highlight("2026Q1业绩"),
    multiColTable(
      ["指标", "数值", "增速"],
      [
        ["营收", "77.74亿", "+1.36%"],
        ["归母净利", "9.28亿", "+2.68%"]
      ],
      [2600, 3200, 3400]
    ),
    para("Q1增速放缓受汽车热管理短期承压。"),

    h1("五、估值分析"),
    kvTable([
      ["PE(TTM)", "约44-50倍"],
      ["PB", "约5.7倍"],
      ["东方证券目标价", "55.55元"],
      ["华泰目标价", "57元"],
      ["国元目标价", "59.17元"],
      ["同花顺一致预期", "52.65元"],
      ["高盛目标价", "40.9-43.1元（反复调整：中性→买入→中性→买入）"]
    ]),

    h1("六、情绪面"),
    bullet("2025年9月涨停创新高46元，10月\"50亿订单\"传闻涨停（后证实不实）"),
    bullet("2026年5月52-54元区间，较3月低点43元反弹20-25%"),
    bullet("核心概念：人形机器人（特斯拉Optimus供应商）+液冷+热管理"),
    bullet("主力净流入曾达22亿元创板块纪录"),

    h1("七、看多/看空"),
    h2("看多逻辑"),
    para("全球制冷龙头稳健+汽车热管理单车价值提升+机器人第三曲线+液冷新方向"),
    h2("看空逻辑"),
    para("PE 44-50x高于同行盾安25x+机器人尚未量产贡献收入+Q1增速放缓+高盛看空"),

    h1("八、综合评价"),
    highlight("核心结论"),
    para("三花的核心价值在于\"传统业务全球龙头提供业绩底部+机器人打开成长天花板\"。约50%的估值溢价来自机器人预期。若Optimus量产不及预期，估值面临回落。当前位于高盛目标价（40.9元）上方约30%，需机器人业务兑现来消化估值。"),

    ...disclaimer()
  ];
  return createDoc("三花智控（002050.SZ）", sections);
}

// ==================== Report 2: 巨轮智能 ====================
function report2() {
  const sections = [
    ...coverPage("个股调研报告", "巨轮智能（002031.SZ）", "轮胎模具老二+央企入主+人形机器人XT减速器概念", TODAY),
    
    h1("一、公司概况"),
    kvTable([
      ["全称", "巨轮智能装备股份有限公司"],
      ["代码", "002031.SZ"],
      ["上市时间", "2004年"],
      ["实控人", "国务院国资委（通过中国通用技术集团，2025年2月完成变更）"],
      ["总部", "广东揭阳"],
      ["市值", "约145亿元"],
      ["股价", "约6.59元（5月27日跌停收盘）"]
    ]),

    h1("二、产业布局"),
    h2("1. 轮胎模具（50-60%）"),
    para("子午线轮胎活络模具，国内第二（仅次于豪迈科技）。"),
    h2("2. 液压式硫化机（20-30%）"),
    para("轮胎硫化成型装备。"),
    h2("3. 工业机器人/减速器（10-15%）"),
    para("RV减速器+XT减速器（人形机器人专用，自主首创）。"),

    h1("三、行业地位"),
    multiColTable(
      ["指标", "详情"],
      [
        ["轮胎模具", "国内第二（豪迈第一，市占30-40%，巨轮约10-15%）"],
        ["XT减速器", "中国首款为人形机器人设计的创新型减速器"],
        ["竞争对手", "豪迈科技（700亿市值，PE~40x）"],
        ["竞争对手", "绿的谐波（200亿，PE>100x）"]
      ],
      [3200, 6000]
    ),

    h1("四、财务数据"),
    highlight("警告：持续亏损"),
    multiColTable(
      ["指标", "数值", "增速/备注"],
      [
        ["2025营收", "7.88亿", "下降"],
        ["2025归母净利", "-2.32亿", "亏损扩大"],
        ["2025毛利率", "10.37%", "-"],
        ["2025 ROE", "-11.72%", "-"],
        ["2026Q1营收", "1.18亿", "-38.95%"],
        ["2026Q1归母净利", "-3584万", "继续亏损"]
      ],
      [2600, 3200, 3400]
    ),
    para("亏损原因：硫化机交付大幅减少+费用刚性。"),

    h1("五、估值分析"),
    kvTable([
      ["PE(TTM)", "亏损无有效PE"],
      ["PB", "约5.9倍"],
      ["券商覆盖", "无主流券商覆盖和目标价"],
      ["定性", "属于纯概念炒作标的"]
    ]),

    h1("六、情绪面"),
    bullet("5月14-15日连续涨停（特斯拉第三代机器人官宣催化）"),
    bullet("5月20日跌停（换手率19%，分歧加剧）"),
    bullet("5月27日再次跌停（6.59元）"),
    bullet("5日主力净流出6.08亿元"),
    bullet("概念：人形机器人+XT减速器+央企改革+国产替代"),

    h1("七、看多/看空"),
    h2("看多逻辑"),
    para("央企背景+XT减速器国产替代+人形机器人风口+技术突破预期"),
    h2("看空逻辑"),
    para("持续亏损（2025年-2.32亿）+营收暴跌39%+PB 5.9x无业绩支撑+纯游资博弈+ST风险"),

    h1("八、综合评价"),
    highlight("核心结论"),
    para("典型的\"概念股\"——没有业绩，全靠故事。央企背景+XT减速器概念吸引游资炒作，但连续亏损、营收暴跌、无券商覆盖。属于高风险投机标的，适合短线博弈，基本面投资者应远离。当前145亿市值完全是情绪定价。"),

    ...disclaimer()
  ];
  return createDoc("巨轮智能（002031.SZ）", sections);
}

// ==================== Report 3: 卧龙电驱 ====================
function report3() {
  const sections = [
    ...coverPage("个股调研报告", "卧龙电驱（600580.SH）", "国内电机市占率第一+人形机器人电机电控概念", TODAY),
    
    h1("一、公司概况"),
    kvTable([
      ["全称", "卧龙电气驱动集团股份有限公司"],
      ["代码", "600580.SH"],
      ["上市时间", "2002年"],
      ["实控人", "陈建成（持卧龙控股48.93%）"],
      ["子公司", "希尔机器人（SIR，全球工业机器人品牌）、上海卧龙电驱等"],
      ["总部", "浙江绍兴上虞"],
      ["市值", "约660-690亿元"],
      ["股价", "约38.95-42元"]
    ]),

    h1("二、产业布局"),
    h2("1. 暖通电驱动（49.58亿，32%）"),
    para("受益全球暖通空调能效升级。"),
    h2("2. 防爆电驱动（46.33亿，30%）"),
    para("石化矿业等高危行业，技术壁垒高。"),
    h2("3. 工业电驱动（~20%+）"),
    para("通用工业电机及驱动。"),
    h2("4. 电动交通（~10%）"),
    para("新能源车用电机电控。"),
    h2("5. 人形机器人"),
    para("子公司希尔（SIR）获智元机器人战略投资。"),

    h1("三、行业地位"),
    multiColTable(
      ["指标", "详情"],
      [
        ["国内电机市场份额", "第一（约6.2%，行业集中度低CR5<20%）"],
        ["全球定位", "全球电机前列，与ABB/西门子/Nidec竞争"],
        ["竞争对手", "汇川技术（2158亿，PE 45-50x）"],
        ["竞争对手", "大洋电机（180-200亿，40-50x）"]
      ],
      [3200, 6000]
    ),

    h1("四、财务数据"),
    highlight("2025全年业绩"),
    multiColTable(
      ["指标", "数值", "增速/备注"],
      [
        ["营收", "154.54亿", "-4.88%（出售子公司导致）"],
        ["归母净利", "11.26亿", "+42.04%"],
        ["毛利率", "25.37%", "-"],
        ["ROE", "10.73%", "-"]
      ],
      [2600, 3200, 3400]
    ),
    para(""),
    highlight("2026Q1业绩"),
    multiColTable(
      ["指标", "数值", "增速"],
      [
        ["营收", "39.86亿", "-1.31%"],
        ["归母净利", "2.76亿", "+2.94%"]
      ],
      [2600, 3200, 3400]
    ),
    para("营收降是报表范围调整，利润端持续修复。"),

    h1("五、估值分析"),
    kvTable([
      ["PE(TTM)", "约56-58倍"],
      ["PB", "约3.41倍"],
      ["机构目标价", "分析师平均约27元（未更新，远低于现价42元）"],
      ["评级", "国元证券维持增持"]
    ]),

    h1("六、情绪面"),
    bullet("5月多次盘中涨停（人形机器人概念带动）"),
    bullet("近3月日均成交16.75亿，交投活跃"),
    bullet("5月27日主力净流入-1.01亿"),
    bullet("核心概念：人形机器人（希尔+智元）+特斯拉+A+H上市预期"),

    h1("七、看多/看空"),
    h2("看多逻辑"),
    para("电机市占率第一+利润修复+42%增长+希尔机器人+智元投资+港股上市"),
    h2("看空逻辑"),
    para("PE 57x远超机构目标价27元+营收下滑+机器人营收贡献仅~3%+概念炒作成分大"),

    h1("八、综合评价"),
    highlight("核心结论"),
    para("卧龙电驱是国内电机龙头，基本面稳健（利润+42%），但当前660亿市值中机器人概念溢价占大头。机器人相关收入贡献目前极小（约3%），PE 57倍远超分析师目标价对应估值。属于\"好公司但贵了\"的状态，需要机器人业务实质兑现才能消化估值。"),

    ...disclaimer()
  ];
  return createDoc("卧龙电驱（600580.SH）", sections);
}

// ==================== Report 4: 长电科技 ====================
function report4() {
  const sections = [
    ...coverPage("个股调研报告", "长电科技（600584.SH）", "国内封测绝对龙头/全球第三，AI先进封装核心标的", TODAY),
    
    h1("一、公司概况"),
    kvTable([
      ["全称", "江苏长电科技股份有限公司"],
      ["代码", "600584.SH"],
      ["上市时间", "2003年"],
      ["实控人", "华润集团（2023年入主）"],
      ["子公司", "星科金朋（新加坡）、长电韩国、长电先进封装"],
      ["总部", "江苏江阴"],
      ["市值", "约2400亿元（5月26日涨停时）"],
      ["股价", "约84.81元"]
    ]),

    h1("二、产业布局"),
    para("全球领先半导体封装测试企业。"),
    kvTable([
      ["先进封装营收", "2025年270亿（占比约69.5%，创历史新高）"],
      ["核心技术", "XDFOI（极高密度扇出）、Chiplet、2.5D/3D封装、FCBGA、SiP"],
      ["应用领域", "通讯36.4%、消费23.6%、HPC/AI高增、汽车电子"],
      ["境外收入", "占比70%+"],
      ["研发费用", "2025年20.86亿（+21.37%）"]
    ]),

    h1("三、行业地位"),
    multiColTable(
      ["指标", "详情"],
      [
        ["全球排名", "全球封测第三（日月光第一、安靠第二）"],
        ["国内排名", "国内封测绝对第一（市占率36.94%）"],
        ["先进封装", "占比最高（约70%）"],
        ["竞争对手", "通富微电（营收279亿，PE 50-60x，AMD绑定）"],
        ["竞争对手", "华天科技（~150亿，PE 60-80x）"]
      ],
      [3200, 6000]
    ),

    h1("四、财务数据"),
    highlight("2025全年业绩"),
    multiColTable(
      ["指标", "数值", "增速/备注"],
      [
        ["营收", "388.71亿", "+8.09%（创新高）"],
        ["归母净利", "15.65亿", "-2.75%"],
        ["ROE", "5.56%", "-"],
        ["增收不增利原因", "-", "先进封装前期投入大+研发费用高增"]
      ],
      [2600, 3200, 3400]
    ),
    para(""),
    highlight("2026Q1业绩"),
    multiColTable(
      ["指标", "数值", "增速"],
      [
        ["营收", "91.71亿", "显著增长"],
        ["归母净利", "2.90亿", "+约50%"]
      ],
      [2600, 3200, 3400]
    ),
    para("Q1业绩大幅改善验证先进封装进入收获期。"),

    h1("五、估值分析"),
    kvTable([
      ["PE(TTM)", "约90-95倍（历史百分位99.86%，极高）"],
      ["PB", "约5.0-5.5倍"],
      ["高盛目标价", "50.9元（中性）"],
      ["瑞银目标价", "79.5元"],
      ["10位分析师平均", "56.84元"],
      ["当前状态", "股价85元已显著超过大多数目标价"],
      ["对比参考", "日月光PE仅15-20x"]
    ]),

    h1("六、情绪面"),
    bullet("5月22/25/26日连续3日涨停（10%+10%+10%），触发异动公告"),
    bullet("公司声明无应披露而未披露重大事项"),
    bullet("核心催化：AI算力爆发 -> 先进封装是算力瓶颈 -> 国产替代 -> XDFOI对标台积电CoWoS -> 华润央企+固投加码100亿"),
    bullet("主力资金持续大幅流入"),

    h1("七、看多/看空"),
    h2("看多逻辑"),
    para("全球第三国内第一+先进封装270亿领先+AI算力确定性受益+XDFOI技术突破+Q1业绩拐点+华润央企背景"),
    h2("看空逻辑"),
    para("PE 93x历史极值+股价远超所有目标价+日月光仅15-20x+先进封装前期仍压毛利+全球贸易风险"),

    h1("八、综合评价"),
    highlight("核心结论"),
    para("长电是AI先进封装的\"正统龙头\"——全球第三、国内第一、先进封装占比70%，基本面过硬。但PE 93倍处于历史极端水平，股价85元远超券商平均目标价57元（+49%）。当前定价已包含极其乐观的AI先进封装爆发预期。如果2026年利润能达到券商预期的24亿，则对应远期PE约100x，需要持续高增长才能消化。短期存在追高风险，中长期取决于先进封装利润释放节奏。"),

    ...disclaimer()
  ];
  return createDoc("长电科技（600584.SH）", sections);
}

// ==================== Report 5: 博云新材 ====================
function report5() {
  const sections = [
    ...coverPage("个股调研报告", "博云新材（002297.SZ）", "C919唯一国产刹车供应商+硬质合金业绩爆发", TODAY),
    
    h1("一、公司概况"),
    kvTable([
      ["全称", "湖南博云新材料股份有限公司"],
      ["代码", "002297.SZ"],
      ["上市时间", "2009年"],
      ["实控人", "湖南省国资委（通过中南大学粉末冶金工程研究中心）"],
      ["子公司", "博云东方（硬质合金）、鑫航机轮刹车、霍尼韦尔博云（合资，C919刹车）"],
      ["总部", "湖南长沙"],
      ["市值", "约135-144亿元"],
      ["股价", "约23.69-25.85元"]
    ]),

    h1("二、产业布局"),
    h2("1. 硬质合金（营收占比>50%）"),
    para("特粗晶/超细纳米硬质合金，博云东方为专精特新小巨人。"),
    h2("2. 航空制动（战略性业务）"),
    para("C919唯一国产刹车系统供应商（霍尼韦尔博云合资），军机粉末冶金刹车副。"),
    h2("3. 航天材料"),
    para("碳/碳复合材料，固体火箭发动机喷管。"),
    h2("4. 稀有金属粉体"),
    para("钨/钼粉末原材料。"),

    h1("三、行业地位"),
    multiColTable(
      ["指标", "详情"],
      [
        ["C919国产刹车系统", "唯一供应商（通过霍尼韦尔博云合资）"],
        ["粉末冶金刹车副", "国内主要供应商，打破国外垄断"],
        ["硬质合金", "博云东方为专精特新小巨人，特粗晶领域技术领先"],
        ["竞争对手", "航材股份（234-267亿，PE 40-56x）"],
        ["竞争对手", "中航高科（240亿+，PE 25-28x）"],
        ["竞争对手", "北摩高科（100亿+，PE 50x+）"]
      ],
      [3200, 6000]
    ),

    h1("四、财务数据"),
    highlight("业绩爆发"),
    multiColTable(
      ["指标", "数值", "增速/备注"],
      [
        ["2025营收", "9.1亿", "+27.77%"],
        ["2025归母净利", "6185万", "扭亏为盈"],
        ["2026Q1营收", "3.80亿", "+125.94%"],
        ["2026Q1归母净利", "1.32亿", "+13362%（133倍！）"],
        ["ROE(TTM)", "约8.7%", "-"]
      ],
      [2600, 3200, 3400]
    ),
    para("爆发原因：硬质合金毛利率大幅提升（钨价上涨+低价库存重估）+航空业务增长。"),

    h1("五、估值分析"),
    kvTable([
      ["PE(TTM)", "约77倍（基于2025全年6185万净利）"],
      ["PE(动态)", "约27倍（基于Q1利润1.32亿年化5亿+估算）"],
      ["PB", "约6.01倍"],
      ["券商覆盖", "无主流券商近期覆盖和目标价"]
    ]),

    h1("六、情绪面"),
    bullet("9天6板：5月期间6次涨停"),
    bullet("5月22日换手率29.18%，成交额42.71亿元"),
    bullet("5月26日跌停（公司发布异动公告后）"),
    bullet("近半年累计上龙虎榜22次"),
    bullet("核心催化：C919量产加速（每10-15天一架）+钨涨价+Q1利润133倍+军工板块轮动+商业航天"),

    h1("七、看多/看空"),
    h2("看多逻辑"),
    para("C919量产受益+硬质合金业绩爆发+Q1利润133倍+军工+商业航天+扭亏拐点"),
    h2("看空逻辑"),
    para("硬质合金高毛利难持续（钨价波动）+静态PE 77x极高+航空营收占比低+游资主导+异动公告后跌停"),

    h1("八、综合评价"),
    highlight("核心结论"),
    para("博云新材处于\"业绩拐点+主题催化\"双重驱动中。Q1利润1.32亿远超2025全年（6185万），若能持续则动态PE约27倍并不贵。核心不确定性在于硬质合金高毛利能否持续（取决于钨价走势），以及C919刹车系统何时贡献实质收入。当前9天6板后情绪已high，短期追高风险大，需等回调或业绩持续验证。"),

    ...disclaimer()
  ];
  return createDoc("博云新材（002297.SZ）", sections);
}

// ==================== Main: Generate all reports ====================
async function main() {
  const reports = [
    { fn: report1, filename: "三花智控002050调研报告.docx" },
    { fn: report2, filename: "巨轮智能002031调研报告.docx" },
    { fn: report3, filename: "卧龙电驱600580调研报告.docx" },
    { fn: report4, filename: "长电科技600584调研报告.docx" },
    { fn: report5, filename: "博云新材002297调研报告.docx" }
  ];

  for (const report of reports) {
    const doc = report.fn();
    const buffer = await Packer.toBuffer(doc);
    const outputPath = path.join(OUTPUT_DIR, report.filename);
    fs.writeFileSync(outputPath, buffer);
    console.log(`Generated: ${outputPath}`);
  }

  console.log("\nAll 5 reports generated successfully!");
}

main().catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
