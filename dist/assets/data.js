window.SITE_DATA = {
  "site": {
    "name": "Skill 展厅",
    "title": "「一句话」把专业活儿做完",
    "desc": "13 个经过业务验证的智能体技能，覆盖公文、会议、经营、数据与办公交付。找得到、抄得走、用得上——把「这事得干两天」变成「这句话我发你」。",
    "badge": "内部智能体平台 · Skill 资产库",
    "team": "智能体平台 · 能力建设组",
    "theme": "#2563eb",
    "version": "v1.0",
    "contact": "skill-platform@example.com",
    "footer": "数据维护在《技能展厅.xlsx》，双击「一键生成.cmd」即可更新本页。",
    "about": ""
  },
  "skills": [
    {
      "id": "gov-doc-writing",
      "name": "公文写作",
      "en": "Official Document Writer",
      "desc": "按单位、文种、角色和现有材料，生成通知、报告、请示等规范公文，打开就能改。",
      "biz": "公文行政",
      "func": "写作与提炼",
      "ownerBiz": "张琳",
      "ownerIT": "王磊",
      "scenario": "接到「发个通知」「写份请示」的临时任务，手上只有零散材料；需要套用单位固定版式。",
      "prompts": [
        "帮我写一份《关于开展2026年度安全生产大检查的通知》，单位是XX集团，发文对象为各子公司，重点检查消防、用电、危化品三项，10月30日前完成自查并上报。",
        "根据下面这段会议记录，整理成向上级请示的公文，请示事项是追加专项预算80万元，语言正式，结尾用「妥否，请批示」。"
      ],
      "value": "从「找模板—改抬头—调格式—核口径」的半日工，变成一次对话拿到可编辑成稿。",
      "saved": 235,
      "weekly": 2.0,
      "tags": [
        "公文",
        "通知",
        "请示",
        "红头"
      ],
      "status": "已上线",
      "featured": true,
      "weight": 100,
      "row": 2,
      "gain": "单次省 3.9 小时",
      "weeklySaved": 470,
      "monthly": 2021,
      "index": 1
    },
    {
      "id": "meeting-recap",
      "name": "会议纪要",
      "en": "Meeting Recap",
      "desc": "把零散发言、录音转写整理成结构化纪要：议题、结论、行动项、负责人、截止日期。",
      "biz": "会议协作",
      "func": "写作与提炼",
      "ownerBiz": "陈晓",
      "ownerIT": "王磊",
      "scenario": "周会、评审会长纪要；会上说法乱、结论散，会后还要挨个追行动项。",
      "prompts": [
        "把下面的会议记录整理成纪要，按「议题—讨论要点—决议—行动项（负责人/截止时间）」输出，没提到的信息不要编。",
        "这是1小时的评审会转写稿，帮我提炼5条关键决议和3个待办，每条标注责任人。"
      ],
      "value": "纪要从「回听录音两小时」变成「贴进去一分钟」，行动项再也不会漏在人名下。",
      "saved": 119,
      "weekly": 3.0,
      "tags": [
        "会议",
        "纪要",
        "行动项"
      ],
      "status": "已上线",
      "featured": true,
      "weight": 95,
      "row": 3,
      "gain": "单次省 2 小时",
      "weeklySaved": 357,
      "monthly": 1535,
      "index": 2
    },
    {
      "id": "data-viz-gen",
      "name": "数据图表",
      "en": "Chart Builder",
      "desc": "按数据类型和业务问题选图、配色、标注，直接给出能进汇报的可编辑图表。",
      "biz": "数据洞察",
      "func": "数据分析",
      "ownerBiz": "赵敏",
      "ownerIT": "周涛",
      "scenario": "汇报要放图，纠结柱状还是折线；配色土、标注乱、重点不突出。",
      "prompts": [
        "这是各区域季度销售额，帮我选最合适的图型，生成带同比标注的图表，配色以公司蓝为主色，重点标出下滑的两个区。"
      ],
      "value": "图表从「配色试半天」变成「说清问题就出图」，一版过会。",
      "saved": 88,
      "weekly": 2.0,
      "tags": [
        "图表",
        "可视化",
        "汇报"
      ],
      "status": "已上线",
      "featured": true,
      "weight": 90,
      "row": 6,
      "gain": "单次省 1.5 小时",
      "weeklySaved": 176,
      "monthly": 757,
      "index": 3
    },
    {
      "id": "saas-metrics-coach",
      "name": "经营指标",
      "en": "Business Metrics",
      "desc": "核对收入、MRR、客户留存、增长假设与计算口径，指出数字背后的真问题。",
      "biz": "经营管理",
      "func": "数据分析",
      "ownerBiz": "赵敏",
      "ownerIT": "周涛",
      "scenario": "月度经营会前，三张表口径对不上，同比环比算法各说各话。",
      "prompts": [
        "这是我们本月的收入、MRR、客户数与流失明细，帮我核对口径，算清MRR增速、净留存和获客成本，指出三个最该关注的异常。"
      ],
      "value": "把「三张表对不上、会上被问住」变成「一套口径、一页结论」。",
      "saved": 177,
      "weekly": 0.25,
      "tags": [
        "经营分析",
        "MRR",
        "留存"
      ],
      "status": "已上线",
      "featured": true,
      "weight": 88,
      "row": 7,
      "gain": "单次省 3 小时",
      "weeklySaved": 44,
      "monthly": 190,
      "index": 4
    },
    {
      "id": "office-pptx",
      "name": "PPT 演示",
      "en": "Slide Deck",
      "desc": "按模板生成结构清晰的汇报 PPT，版式统一、每页一个观点、可直接改。",
      "biz": "办公交付",
      "func": "文档生成",
      "ownerBiz": "陈晓",
      "ownerIT": "李静",
      "scenario": "临时被通知明天要汇报，晚上才开始做 PPT。",
      "prompts": [
        "用公司模板做一份10页的季度复盘PPT：结论先行、每页一个观点、数据用图表讲，别写大段文字。"
      ],
      "value": "从「熬夜排版」到「半小时出可讲的稿」。",
      "saved": 240,
      "weekly": 0.5,
      "tags": [
        "PPT",
        "汇报",
        "模板"
      ],
      "status": "已上线",
      "featured": true,
      "weight": 85,
      "row": 11,
      "gain": "单次省 4 小时",
      "weeklySaved": 120,
      "monthly": 516,
      "index": 5
    },
    {
      "id": "work-report-writer",
      "name": "工作汇报",
      "en": "Work Report",
      "desc": "按手头记录整理日报、周报、月报，自动核对交付、数字、风险与下期计划。",
      "biz": "经营管理",
      "func": "写作与提炼",
      "ownerBiz": "陈晓",
      "ownerIT": "李静",
      "scenario": "周五下班前拼周报；月度述职要同时讲清数据、风险和下一步。",
      "prompts": [
        "根据我这周的5条工作记录，写一份周报，分「本周交付/关键数字/风险与求助/下周计划」四块，控制在400字内，别堆形容词。"
      ],
      "value": "周报从「翻聊天记录凑字数」变成「喂记录出结构」，口径统一、领导好读。",
      "saved": 88,
      "weekly": 1.0,
      "tags": [
        "周报",
        "月报",
        "述职"
      ],
      "status": "已上线",
      "featured": false,
      "weight": 80,
      "row": 4,
      "gain": "单次省 1.5 小时",
      "weeklySaved": 88,
      "monthly": 378,
      "index": 6
    },
    {
      "id": "office-docx",
      "name": "Word 文档",
      "en": "Word Document",
      "desc": "生成、修改、校验 Word：保留样式、页眉页脚、编号与图表，交付即可用。",
      "biz": "办公交付",
      "func": "文档生成",
      "ownerBiz": "陈晓",
      "ownerIT": "王磊",
      "scenario": "要交的是可编辑的正式文档，不是一段聊天文字。",
      "prompts": [
        "把这份内容做成Word：A4、仿宋三号、一级标题黑体、带页码页脚，输出可编辑文件。"
      ],
      "value": "交付物直接能发，不用再手工排版半小时。",
      "saved": 40,
      "weekly": 1.0,
      "tags": [
        "Word",
        "排版",
        "交付"
      ],
      "status": "已上线",
      "featured": false,
      "weight": 75,
      "row": 9,
      "gain": "单次省 40 分钟",
      "weeklySaved": 40,
      "monthly": 172,
      "index": 7
    },
    {
      "id": "office-xlsx",
      "name": "Excel 表格",
      "en": "Excel Workbook",
      "desc": "建表、写公式、做透视与图表，保留原生公式，后续还能自己改。",
      "biz": "办公交付",
      "func": "文档生成",
      "ownerBiz": "赵敏",
      "ownerIT": "李静",
      "scenario": "手工拉公式、做透视表、一格一格调格式。",
      "prompts": [
        "把这份流水做成月度汇总表，用SUMIFS按部门和月份汇总，加同比增长列和柱状图。"
      ],
      "value": "表格从「拉公式拉到眼花」变成「说需求出成品」。",
      "saved": 60,
      "weekly": 1.0,
      "tags": [
        "Excel",
        "公式",
        "透视表"
      ],
      "status": "已上线",
      "featured": false,
      "weight": 72,
      "row": 10,
      "gain": "单次省 1 小时",
      "weeklySaved": 60,
      "monthly": 258,
      "index": 8
    },
    {
      "id": "research-writer",
      "name": "研究报告",
      "en": "Research Report",
      "desc": "从研究问题到证据、来源、判断与风险，产出每句结论都带出处的研究报告。",
      "biz": "市场研究",
      "func": "写作与提炼",
      "ownerBiz": "李静",
      "ownerIT": "周涛",
      "scenario": "新市场、新赛道摸底，需要一份能直接进决策会讨论的材料。",
      "prompts": [
        "围绕「县域咖啡市场是否值得进入」写一份研究报告：市场规模、竞争格局、进入方式三选一、主要风险，每个结论后面标注数据来源。"
      ],
      "value": "把「搜三天资料还写不出结论」变成「两小时出带出处的初稿」。",
      "saved": 360,
      "weekly": 0.25,
      "tags": [
        "研究",
        "调研",
        "决策"
      ],
      "status": "试点",
      "featured": false,
      "weight": 70,
      "row": 5,
      "gain": "单次省 6 小时",
      "weeklySaved": 90,
      "monthly": 387,
      "index": 9
    },
    {
      "id": "dataset-health-audit",
      "name": "数据体检",
      "en": "Data Health Check",
      "desc": "检查缺失、异常、单位、口径、重复与可复现计算，先体检再分析。",
      "biz": "数据洞察",
      "func": "数据分析",
      "ownerBiz": "王磊",
      "ownerIT": "赵敏",
      "scenario": "拿到的业务数据不敢直接用，又说不清到底哪里有问题。",
      "prompts": [
        "帮我体检这份销售明细：缺失率、异常值、单位是否统一、有无重复订单，输出问题清单和修复建议。"
      ],
      "value": "避免「分析三天，结论因脏数据作废」这种最贵的返工。",
      "saved": 120,
      "weekly": 0.5,
      "tags": [
        "数据质量",
        "清洗",
        "校验"
      ],
      "status": "已上线",
      "featured": false,
      "weight": 60,
      "row": 8,
      "gain": "单次省 2 小时",
      "weeklySaved": 60,
      "monthly": 258,
      "index": 10
    },
    {
      "id": "office-export-pdf",
      "name": "文档转 PDF",
      "en": "PDF Export",
      "desc": "Word / Excel / PPT 转 PDF，版式不乱、字体不丢、页码不断。",
      "biz": "办公交付",
      "func": "格式转换",
      "ownerBiz": "李静",
      "ownerIT": "王磊",
      "scenario": "对外发送必须 PDF，转完格式全乱、表格跨页断行。",
      "prompts": [
        "把这个Word转成PDF，检查页码和表格有没有跨页断裂，有问题先告诉我。"
      ],
      "value": "对外交付不再「转一次修一次」。",
      "saved": 20,
      "weekly": 2.0,
      "tags": [
        "PDF",
        "转换",
        "对外"
      ],
      "status": "已上线",
      "featured": false,
      "weight": 50,
      "row": 12,
      "gain": "单次省 20 分钟",
      "weeklySaved": 40,
      "monthly": 172,
      "index": 11
    },
    {
      "id": "office-render",
      "name": "文档预览",
      "en": "Doc Preview",
      "desc": "把文档渲染成图片，发出去之前先看效果，版式问题一眼发现。",
      "biz": "办公交付",
      "func": "质量校验",
      "ownerBiz": "李静",
      "ownerIT": "周涛",
      "scenario": "文件发出去才发现文字溢出、图片错位，只能道歉重发。",
      "prompts": [
        "把这个PPT每页渲染成图片，我检查一下有没有文字溢出和错位。"
      ],
      "value": "交付前自查一分钟，省掉一轮返工和一次尴尬。",
      "saved": 15,
      "weekly": 1.0,
      "tags": [
        "预览",
        "渲染",
        "质检"
      ],
      "status": "试点",
      "featured": false,
      "weight": 45,
      "row": 13,
      "gain": "单次省 15 分钟",
      "weeklySaved": 15,
      "monthly": 64,
      "index": 12
    },
    {
      "id": "office-recalculate",
      "name": "表格重算",
      "en": "Sheet Recalculate",
      "desc": "刷新公式缓存并核对计算值，避免「一打开就是错的数」。",
      "biz": "办公交付",
      "func": "质量校验",
      "ownerBiz": "王磊",
      "ownerIT": "赵敏",
      "scenario": "表格公式结果和手算对不上，或者打开一片 0。",
      "prompts": [
        "重算这个工作簿的所有公式，把结果和手工核算的差异逐项列出来。"
      ],
      "value": "数字可信，汇报时不被追问第二遍。",
      "saved": 30,
      "weekly": 0.5,
      "tags": [
        "公式",
        "重算",
        "校验"
      ],
      "status": "规划",
      "featured": false,
      "weight": 40,
      "row": 14,
      "gain": "单次省 30 分钟",
      "weeklySaved": 15,
      "monthly": 64,
      "index": 13
    }
  ],
  "cases": [
    {
      "id": "C01",
      "title": "两小时录音，变成一分钟纪要",
      "skill": "会议纪要",
      "role": "项目助理",
      "beforeText": "回听录音3遍→手打讨论要点→群里逐个确认口径→整理成行动项表格",
      "beforeSteps": [
        "回听录音3遍",
        "手打讨论要点",
        "群里逐个确认口径",
        "整理成行动项表格"
      ],
      "nowPrompt": "把1小时评审会转写稿整理成纪要：议题、决议、行动项（负责人/截止时间），没提到的别编。",
      "before": 120,
      "now": 1,
      "value": "会后10分钟发出纪要，5个行动项直接落到人头，下一次会不用再对上次的账。",
      "home": true,
      "row": 2
    },
    {
      "id": "C02",
      "title": "三张表对不上，变成一页结论",
      "skill": "经营指标",
      "role": "经营分析",
      "beforeText": "三个系统各导一张表→手工对齐口径→Excel透视→算同比→会上被问住现场重算",
      "beforeSteps": [
        "三个系统各导一张表",
        "手工对齐口径",
        "Excel透视",
        "算同比",
        "会上被问住现场重算"
      ],
      "nowPrompt": "核对本月收入、MRR、客户与流失明细的口径，算清增速、净留存、获客成本，指出三个异常。",
      "before": 180,
      "now": 3,
      "value": "经营会从「解释数字」变成「讨论动作」，口径争议当场消掉。",
      "home": true,
      "row": 3
    },
    {
      "id": "C03",
      "title": "半天的通知，变成一句话",
      "skill": "公文写作",
      "role": "行政 / 文秘",
      "beforeText": "翻去年同类文→改抬头改日期→调字体行距→找人核口径→打印校对三遍",
      "beforeSteps": [
        "翻去年同类文",
        "改抬头改日期",
        "调字体行距",
        "找人核口径",
        "打印校对三遍"
      ],
      "nowPrompt": "写一份《安全生产大检查通知》，单位XX集团，对象各子公司，查消防/用电/危化品，10月30日前上报。",
      "before": 240,
      "now": 5,
      "value": "当天下午就能发文，格式一次过，文秘不用再做「排版工」。",
      "home": true,
      "row": 4
    },
    {
      "id": "C04",
      "title": "配色试五版，变成说清问题就出图",
      "skill": "数据图表",
      "role": "业务汇报人",
      "beforeText": "挑图型→配色试5版→手工加标注→调坐标轴→导出贴进PPT",
      "beforeSteps": [
        "挑图型",
        "配色试5版",
        "手工加标注",
        "调坐标轴",
        "导出贴进PPT"
      ],
      "nowPrompt": "用各区域季度销售额选最合适的图，带同比标注，公司蓝为主色，突出下滑的两个区。",
      "before": 90,
      "now": 2,
      "value": "汇报图一眼看出重点，领导不再问「所以呢」。",
      "home": false,
      "row": 5
    },
    {
      "id": "C05",
      "title": "周五拼周报，变成喂记录出结构",
      "skill": "工作汇报",
      "role": "一线主管",
      "beforeText": "翻聊天记录和邮件→拼凑本周条目→补数字→统一措辞→反复改语气",
      "beforeSteps": [
        "翻聊天记录和邮件",
        "拼凑本周条目",
        "补数字",
        "统一措辞",
        "反复改语气"
      ],
      "nowPrompt": "根据这周5条工作记录写周报，分交付/数字/风险/下周计划，400字内。",
      "before": 90,
      "now": 2,
      "value": "周五下午不再耗在措辞上，周报还能看出趋势。",
      "home": false,
      "row": 6
    }
  ],
  "guide": {
    "steps": [
      {
        "title": "找到你要的能力",
        "body": "在技能库按业务分类、功能分类或关键词筛选；卡片上的「复制提示词」可以直接用。"
      },
      {
        "title": "把示例改成你的场景",
        "body": "复制后替换括号里的业务信息（单位、时间、数据口径）。写得越具体，产出越接近成稿。"
      },
      {
        "title": "交给智能体执行",
        "body": "粘贴发送即可。需要交付文件时，直接说「输出 Word / Excel / PPT」。"
      },
      {
        "title": "有问题找谁",
        "body": "卡片上标了业务负责人（提需求、验收效果）和 IT 负责人（维护提示词、修问题）。"
      }
    ],
    "faq": [
      {
        "title": "网页需要联网吗？",
        "body": "不需要。整个站点是本地文件，双击 index.html 就能看，断网、内网、发给别人都一样。"
      },
      {
        "title": "提示词能直接改吗？",
        "body": "能。示例是起点不是标准答案，把占位信息换成真实业务，效果会明显更好。"
      },
      {
        "title": "数据在哪里维护？",
        "body": "全部在《技能展厅.xlsx》。改完双击「一键生成.cmd」，dist 目录里的网页就更新了。"
      },
      {
        "title": "业务分类和功能分类有什么区别？",
        "body": "业务分类是「哪个业务域在用」，功能分类是「这个能力干什么活」。两者交叉，能快速定位。"
      },
      {
        "title": "能发给外部或客户看吗？",
        "body": "可以，但请先清掉内部联系方式、未公开的业务数据和试点中的技能。"
      }
    ]
  },
  "meta": {
    "generated": "2026-10-26 23:30",
    "skillCount": 13,
    "caseCount": 5,
    "bizCount": 6,
    "funcCount": 5,
    "savedMinutes": 707,
    "bizCategories": [
      "办公交付",
      "数据洞察",
      "经营管理",
      "会议协作",
      "公文行政",
      "市场研究"
    ],
    "funcCategories": [
      "写作与提炼",
      "数据分析",
      "文档生成",
      "质量校验",
      "格式转换"
    ],
    "weeklyCalls": 15.0,
    "weeklySavedMin": 1575,
    "monthlySavedMin": 6774,
    "manDays": 14.1,
    "domains": [
      {
        "name": "公文行政",
        "count": 1,
        "weekly": 470,
        "skills": [
          "公文写作"
        ],
        "share": 30
      },
      {
        "name": "会议协作",
        "count": 1,
        "weekly": 357,
        "skills": [
          "会议纪要"
        ],
        "share": 23
      },
      {
        "name": "办公交付",
        "count": 6,
        "weekly": 290,
        "skills": [
          "PPT 演示",
          "Word 文档",
          "Excel 表格",
          "文档转 PDF",
          "文档预览",
          "表格重算"
        ],
        "share": 18
      },
      {
        "name": "数据洞察",
        "count": 2,
        "weekly": 236,
        "skills": [
          "数据图表",
          "数据体检"
        ],
        "share": 15
      },
      {
        "name": "经营管理",
        "count": 2,
        "weekly": 132,
        "skills": [
          "经营指标",
          "工作汇报"
        ],
        "share": 8
      },
      {
        "name": "市场研究",
        "count": 1,
        "weekly": 90,
        "skills": [
          "研究报告"
        ],
        "share": 6
      }
    ]
  }
};
