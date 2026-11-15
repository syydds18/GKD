using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Outline.Addin
{
    internal sealed class GeneralHelpForm : Form
    {
        internal GeneralHelpForm()
        {
            Text = "搞快点 1.6.0 · 使用说明";
            Font = new Font("Microsoft YaHei UI", 9F);
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            StartPosition = FormStartPosition.CenterParent;
            ShowIcon = false;
            MinimizeBox = false;
            ClientSize = new Size(820, 720);
            MinimumSize = new Size(600, 420);
            BackColor = Color.White;

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(20) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "按“搞快点”功能区从左到右排列；操作前请先保存工作簿。", Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 12) }, 0, 0);

            // One native scroll owner avoids nested AutoSize/TableLayout reflows and
            // child controls intercepting the mouse wheel.
            var content = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None,
                DetectUrls = true, ScrollBars = RichTextBoxScrollBars.Vertical, WordWrap = true,
                BackColor = Color.White, TabStop = true, Font = Font, AccessibleName = "搞快点使用说明" };
            AddSection(content, "1. 智能分级",
                "智能分级：按编号重建所选工作表的原生分级，并展开原有隐藏行列。\r\n清除分级：清除所选表的分组，并展开隐藏行列。\r\n恢复上次分级：恢复本次 Excel 会话最近一批分级及隐藏状态；插入、删除或移动行列后不要使用。");
            AddSection(content, "2. 智能求和",
                "智能求和：按直接子项预览父项 SUM 公式，仅写入勾选项。\r\n撤销上次求和：撤销本次会话最近一批求和；快照变化时拒绝覆盖。");
            AddSection(content, "3. 求和校验",
                "求和校验：只读检查父子项金额；不改公式和金额，双击结果可定位原单元格。");
            AddSection(content, "4. 分级显示",
                "显示一级、显示二级、显示三级：折叠到对应层级。\r\n全部展开：展开全部分级。");
            AddSection(content, "5. 项目提取",
                "水电工程提项目：选择工作表与打印范围，在副本中清理并导出，不修改源工作簿。");
            AddSection(content, "6. AI 助手",
                "AI 设置：配置接口和密钥，测试后保存并设为默认；★ 表示默认配置。\r\n自定义公式：查看本地列表、含义和参数；新建、编辑、删除或 JSON 导入导出，预览确认后保存，算法须复核。\r\n刷新 AI 结果：重新请求所选 GKD 公式。\r\n取消 AI 任务：取消全部排队和运行中的 AI 任务。\r\n结果转为值：转换完整溢出区域，建议先保存副本。\r\n公式：=GKD(\"问题\") 纯文本问答，不隐式读取工作簿；=GKD(\"要求\",,1) 无数据表格；=GKD() 显示欢迎语。表格需动态数组及空白输出区。参数引用与自定义公式调用见 ★ 使用帮助与示例。");
            AddSection(content, "7. 说明",
                "使用说明：打开本窗口，查看全部功能及项目链接。\r\n运行要求：Windows 桌面 Excel 2010+（x86/x64）、.NET Framework 4.8；动态表格需 Microsoft 365 或 Excel 2021/2024 等支持动态数组的版本。");
            AddLink(content, "下载与更新：https://github.com/syydds18/GKD/releases", "https://github.com/syydds18/GKD/releases");
            AddLink(content, "项目源码：https://github.com/syydds18/GKD", "https://github.com/syydds18/GKD");
            layout.Controls.Add(content, 0, 1);

            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 12, 0, 0) };
            var close = new Button { Text = "关闭", AutoSize = true, MinimumSize = new Size(90, 32), DialogResult = DialogResult.Cancel };
            footer.Controls.Add(close);
            layout.Controls.Add(footer, 0, 2);
            CancelButton = close;
            Controls.Add(layout);
            Shown += (sender, args) => BeginInvoke((Action)(() =>
            {
                if (IsDisposed || !content.IsHandleCreated) return;
                content.SelectionStart = 0;
                content.SelectionLength = 0;
                content.ScrollToCaret();
            }));
        }

        private static void AddSection(RichTextBox panel, string title, string text)
        {
            if (panel.TextLength > 0) panel.AppendText(Environment.NewLine + Environment.NewLine);
            int start = panel.TextLength;
            panel.AppendText(title + Environment.NewLine);
            panel.Select(start, title.Length);
            panel.SelectionFont = new Font(panel.Font, FontStyle.Bold);
            panel.Select(panel.TextLength, 0);
            panel.AppendText(text);
        }

        private void AddLink(RichTextBox panel, string text, string url)
        {
            if (panel.TextLength > 0) panel.AppendText(Environment.NewLine + Environment.NewLine);
            int start = panel.TextLength;
            panel.AppendText(text);
            panel.Select(start, text.Length);
            panel.SelectionColor = Color.Blue;
            panel.Select(panel.TextLength, 0);
            panel.MouseUp += (sender, args) =>
            {
                if (args.Button != MouseButtons.Left) return;
                int index = panel.GetCharIndexFromPosition(args.Location);
                if (index < start || index >= start + text.Length) return;
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex) { MessageBox.Show(panel.FindForm(), "无法打开浏览器，请复制以下地址：\r\n" + url + "\r\n\r\n" + ex.Message, "打开链接失败", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            };
        }
    }
}
