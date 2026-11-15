using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Outline.Addin
{
    internal sealed class GkdAiHelpForm : Form
    {
        internal GkdAiHelpForm()
        {
            Ui.Base(this, new Size(920, 740), "★ GKD 使用帮助与实用示例");
            MinimumSize = new Size(640, 480);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var document = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true, WordWrap = true,
                DetectUrls = false, ScrollBars = RichTextBoxScrollBars.Vertical,
                BorderStyle = BorderStyle.None, BackColor = SystemColors.Window,
                Font = Font, AccessibleName = "GKD 使用帮助与实用示例"
            };
            document.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.A) { document.SelectAll(); e.SuppressKeyPress = true; }
            };
            layout.Controls.Add(document, 0, 0);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            var close = Ui.Button("关闭");
            close.Click += (s, e) => Close();
            footer.Controls.Add(close);
            layout.Controls.Add(footer, 0, 1);
            Controls.Add(layout);
            CancelButton = close;

            var text = new StringBuilder();
            AddText("先配置，再写公式", true);
            AddText("在 AI 模型设置中填写并保存配置，设为默认；先做连接 / 鉴权测试，需要表格时再做表格能力测试。函数实际注册名为 GKD，“搞快点”是产品名称，不是公式别名。下面公式可在本只读帮助文档中选中后按 Ctrl+C 复制。示例数据是演示输入，不代表已核实的工程、价格或银行数据。");
            AddText("一、单值输出：省略第三参数，或填 0", true);
            AddText("语法：=GKD(提示词[,数据[,输出类型]])。=GKD(\"问题\") 可直接纯文本问答，不会隐式读取当前选区、其他单元格或工作簿；提示词写 A2 不等于读取 A2。需要资料时必须用第二参数明确引用或传值，提示词本身也可引用单元格。默认返回单值，即使要求列表也不会自动变成二维表格。=GKD() 显示欢迎语“感谢使用！”，不请求模型。");
            AddCode("=GKD(\"问题\")\r\n=GKD(\"总结\",A2)\r\n=GKD(\"总结\",A2,0)\r\n=GKD($H$1,A2)\r\n=GKD()", 5);

            AddText("1. 中文在前，俄语在后：工程清单逐行翻译", true);
            AddText("A2 放中文清单名称，例如“无缝钢管安装”。在 B2 输入下式：先保留 A2 中文，再用 CHAR(10) 换行，最后由 GKD 翻译 A2。开启 Excel“自动换行”以显示中俄双语。向下拖动 B2 的填充柄，A2 会依次变为 A3、A4；不要把数据引用锁成 $A$2。每一行会形成独立请求，请先小批量试算，并检查术语与单位。");
            AddCode("=A2&CHAR(10)&GKD(\"将这条工程清单名称翻译成俄语，只返回俄语译文，保留规格、型号和单位，不加解释\",A2)", 3);
            AddText("也可直接让 AI 返回中俄双语，用下面的 GKD 公式；若必须逐字保留原文，优先用上面的 A2&CHAR(10) 拼接写法。");
            AddCode("=GKD(\"翻译成俄语。第一行原样保留中文，第二行输出俄语译文，中间换行；只返回这两行，保留规格型号和单位\",A2)", 3);
            AddText("只要俄语译文时，用普通翻译指令即可。");
            AddCode("=GKD(\"翻译成俄语，只返回译文\",A2)", 2);

            AddText("2. 按给定规格计算钢管理论质量", true);
            AddText("A1:D2 示例：表头“外径(mm)、壁厚(mm)、长度(m)、密度(kg/m³)”；第二行为 219、6、12、7850。外径 D、壁厚 t、长度 L 对应理论质量 m = π × (D − t) × t × L × ρ / 1000000，结果单位 kg。下式要求列出代入过程，并与普通 Excel 数值公式交叉校核。");
            AddCode("=GKD(\"仅按给定外径D(mm)、壁厚t(mm)、长度L(m)、密度ρ(kg/m³)计算钢管理论质量。先核验D>2t、t>0、L>0、ρ>0；用m=π*(D-t)*t*L*ρ/1000000，单位kg。列出参数、单位、代入公式和结果；缺项则说明，不假设规格、密度或损耗\",A1:D2)\r\n=PI()*(A2-B2)*B2*C2*D2/1000000", 5);
            AddText("延米重（每米理论质量）：将长度取 1m，w=π×(D−t)×t×ρ/1000000，单位 kg/m。下面同时给出 AI 说明和确定性 Excel 计算写法。");
            AddCode("=GKD(\"按给定外径D(mm)、壁厚t(mm)、密度ρ(kg/m³)计算钢管延米重。使用w=π*(D-t)*t*ρ/1000000，单位kg/m，列出代入过程；核验尺寸有效，不假设缺失参数\",(A1:B2,D1:D2))\r\n=PI()*(A2-B2)*B2*D2/1000000", 5);
            AddText("这是均匀圆环截面的理论值，未含焊缝、涂层、附件、允许偏差及损耗。材料牌号、产品标准及版本、壁厚偏差和计量规则须以合同、图纸及适用标准核对；不得把理论计算当作过磅重量或合规证明。实际规格不符时先修正输入。");

            AddText("3. 依据给定定额参考推荐子目", true);
            AddText("A1:D20 放项目描述、施工工艺与条件，以及适用地区、专业、定额名称和版本、候选子目编码、名称、单位、工作内容、章说明等参考原文。第二参数必须包含这些参考资料；不要只给一个工程名称就要求模型编造定额编号。");
            AddCode("=GKD(\"仅依据给定项目说明及定额参考，推荐最匹配的候选子目。逐项写出参考中已有的编码、名称、单位、适用条件、匹配理由、出处及待核对项；禁止编造编码或引用未提供的定额。参考不足则回答资料不足，并列出缺失的地区、版本、工艺或换算依据。最终由造价人员核对\",A1:D20)", 4);

            AddText("4. 从给定询价数据比较市场价格", true);
            AddText("A1:H10 放名称、规格、供应商、报价日期、地区、币种/计价单位、单价及含税含运条件（条件可合并在单元格说明）。只比较相同口径的数据；不同税率、运距、规格或单位须先换算并写明依据。没有真实来源与日期时，输出不能称为当前市场价。");
            AddCode("=GKD(\"仅依据给定报价数据比较同规格材料价格。标明来源、日期、地区、币种、单位、含税含运条件；能统一口径才计算最低、最高与中位价，列出纳入的样本。无法换算或条件不明时单独列示，不补造实时价格，并提示复核报价有效期\",A1:H10)", 4);

            AddText("5. 汇率：已给银行数据与仅给网址是两种场景", true);
            AddText("已有银行牌价：A1:F8 放银行名称、日期时间、币种、每 100 外币对应人民币的现汇买入价及现汇卖出价等原始数据；H1 放交易方向与金额。明确每 100 单位口径，区分现汇/现钞、买入/卖出；客户购汇通常看银行卖出价，结汇通常看银行买入价，实际以银行交易规则及成交汇率为准。数据时点不等于实时成交时点。");
            AddCode("=GKD(\"仅依据给定银行牌价和交易要求，说明选用的银行、时间、币种、现汇或现钞及买入或卖出方向，列出每100外币兑人民币的单位和换算过程；不假设实时汇率，不计未提供的手续费，资料不足则指出\",(A1:F8,H1))", 4);
            AddText("仅给网址：A2 可填银行官方牌价网页完整 URL，B2 填日期、币种和购汇/结汇方向。只有当前后端工作流已配置联网浏览且成功读取网页时，下面任务才可能完成。GKD 加载项自身不会抓取网址；提供 URL 不等于已把网页内容提交给模型。未启用联网、网页不可访问或数据无法核实时，应粘贴银行原始数据后使用上一示例，不接受模型凭记忆填数。");
            AddCode("=GKD(\"请使用后端联网浏览访问给定银行官方URL，按给定日期、币种和交易方向查找牌价。返回银行名称、牌价时间、数值、每100单位口径、方向、来源URL与访问时间；无法联网或未读到对应数据请明确说明，禁止猜测\",A2:B2)", 4);

            AddText("6. 按水电工程概算定额识别项目工程类别", true);
            AddText("A2 放项目名称，可在 B2 输入下式，再向下拖填。只给名称时属于初步识别；同名项目可能因施工内容、部位或工艺不同而分属不同类别，信息不足应标为待核实。");
            AddCode("=GKD(\"根据水电工程概算定额的工程分类，识别给定项目名称所属的工程类别。仅返回类别名称；信息不足或类别不唯一时返回待核实，不编造定额编号\",A2)", 3);
            AddText("需要按指定版本准确分类时，将定额名称、年份版本和分类目录或相关条文放入 H1:K30，并在第二参数中明确引用。下面将参考区域锁定，向下拖填时仅项目名称 A2 随行变化；类别只允许取自所给目录，最终仍须核对定额适用范围。");
            AddCode("=GKD(\"依据所给水电工程概算定额名称、版本及分类目录，识别当前项目名称所属的工程类别。仅返回目录已有的类别名称；有歧义或缺少施工内容时返回待核实，不编造目录或定额编号\",(A2,$H$1:$K$30))", 4);

            AddText("二、二维表格输出：第三参数必须填 1", true);
            AddText("第三参数填 1 返回二维表格（标题、空行、表头及数据），省略或填 0 都是单值。第二参数可留空：=GKD(\"要求\",,1) 允许无数据生成表格；这也不会读取工作簿。表格需要支持动态数组的 Excel（Microsoft 365 / 2021 / 2024），公式右侧和下方须留足空白，不要放在 Excel 的结构化表格内部。遇到 #SPILL! 先清理阻挡区域或合并单元格。AI 返回的内容仍需人工核对。");
            AddCode("=GKD(\"要求\",,1)\r\n=GKD(\"生成一张空白材料登记表，列为名称、数量、单位\",,1)\r\n=GKD(\"整理为名称、数量两列\",A2,1)", 4);

            AddText("7. 根据给定项目来源整理投资估算表", true);
            AddText("A1:F30 放已取得的项目来源资料：来源文件/链接、页码或条款、项目名称、投资项目、金额及币种/单位、价格基期与范围。若有多个项目，应把各项目的来源和数据一并放入区域。网址未读取时请先粘贴相关原文；不要让模型凭项目名称补造总投资。");
            AddCode("=GKD(\"仅依据给定项目来源整理投资估算对比表。列为项目名称、投资构成、金额、币种及单位、价格基期、来源文件及页码或URL、待核实事项。保留原始数据；仅对同币种同口径且无重复的项目求和并说明公式。缺项写未提供，不虚构工程量、单价、汇率、来源或总投资\",A1:F30,1)", 4);
            AddText("如需分析不同地区、年份或币种的投资差异，应补充明确的价格指数、汇率、范围调整和计算依据；表格只是资料整理与初步分析，不能替代正式概算、投资决策或审定结论。");

            AddText("三、本地自定义公式", true);
            AddText("在“AI 助手 → 自定义公式”打开本地列表，选择一项查看含义说明、参数顺序、输出单位和算法定义。支持新建、编辑、删除、JSON 导入 / 导出。手写定义或点击“AI 设计”生成草案后，核对参数单位、适用范围和边界条件，再预览并确认保存。AI 草案不会自动保存；生成成功或格式检查通过不代表算法已验证，须用独立算例复核。");
            AddText("确定性表达式在本地运行，仅支持声明的输入名、数字、+ - * / ^、括号及受限 CONVERT 单位换算，例如 CONVERT(X,\"mm\",\"m\")。换算限内置单位的同维度转换，单位名区分大小写，不支持任意 Excel 函数。AI 提示词模板把要求与显式参数值发送到当前默认后端，仍须复核模型输出。两种类型都不会执行任意生成代码。");
            AddText("先保存并启用名为 PIPE_MASS 的定义，再用动态名称别名或稳定入口调用同一公式；下面不表示内置了已验证的钢管算法。按声明顺序提供全部参数，最多声明 32 个参数、传入 32 个值；确定性表达式的每个参数须为单个有限数值。Excel 标识及输入名使用大写字母开头的 A-Z、数字或下划线，最长 31 字符；公式标识不能与保留名或单元格地址冲突。");
            AddCode("=PIPE_MASS(A2,B2)\r\n=GKD_CUSTOM(\"PIPE_MASS\",A2,B2)", 2);
            AddText("辅助公式存于当前 Windows 用户的 %LOCALAPPDATA%\\GKD\\custom-formulas.json，不会自动随工作簿携带。换电脑或分享前导出 JSON，在目标电脑导入并预览确认；遇到同名或同标识冲突先处理旧定义。公式 JSON 不含 API Key 等密钥，请勿把凭据写进说明或模板；AI 类型需另行配置默认后端。删除、停用或改名会影响引用它的公式。");

            AddText("四、引用、刷新与设置排查", true);
            AddText("Ctrl 选择分散单元格时，用括号包住第二参数中的引用，如 (A1,C3:E3)。一次最多提交 10000 个单元格；不要包含不应上传的敏感资料。提示词提到的其他单元格不会被自动读取，必须加入第二参数引用。一次调用是单轮请求，不会继承上一单元格对话。");
            AddCode("=GKD(\"按所给资料总结，并标明缺失信息\",(A1,C3:E3))", 2);
            AddText("保存或切换默认配置后，已完成的结果不会自动更新；选中结果，点击“刷新 AI 结果”重新运行。批量拖填可能触发较多请求及服务费用；请先检查少量样本。模型输出可能有误，工程计算、定额、价格、汇率和投资结论都应回查来源。");
            AddText("完整请求 URL 不自动追加路径，例如 https://服务地址/v1/chat/completions 或 https://服务地址/v1/workflows/run。Dify 使用阻塞工作流，输入/输出字段必须与实际工作流一致；模型 / 路由为空时由工作流决定。API Key 只保存到当前 Windows 用户的加密配置，勿放进工作表或截图。");
            AddText("连接 / 鉴权测试和表格能力测试都使用当前设置（含未保存编辑），只发送内置合成数据，不读取工作簿，也不会自动保存配置。连接通过不代表表格协议一定通过。失败时查看安全错误详情中的请求编号、阶段与类别，点击设置页“打开诊断日志”定位同一次请求；共享日志前仍应检查敏感内容。");

            document.Text = text.ToString();
            Shown += (s, e) => BeginInvoke((Action)(() =>
            {
                if (IsDisposed || !document.IsHandleCreated) return;
                document.SelectionStart = 0;
                document.SelectionLength = 0;
                document.ScrollToCaret();
            }));

            void AddText(string value, bool heading = false)
            {
                if (text.Length > 0) text.AppendLine().AppendLine();
                text.Append(value);
            }

            void AddCode(string formula, int lines)
            {
                if (text.Length > 0) text.AppendLine();
                text.Append(formula);
            }
        }
    }
}
