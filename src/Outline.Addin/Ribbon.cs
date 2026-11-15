using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace Outline.Addin
{
    [ComVisible(true)]
    public sealed class OutlineRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='tabOutlineTool' label='搞快点'>
    <group id='grpOutline' label='智能分级'>
      <button id='auto' label='智能分级' size='large' imageMso='OutlineShowDetail' onAction='Run' screentip='按编号重建分级' supertip='选择工作表后，按中文、数字、点号等编号推导层级。先清除原分级及隐藏状态，不修改数值、公式或样式。'/>
      <button id='clear' label='清除分级' imageMso='ClearAll' onAction='Run'/>
      <button id='restore' label='恢复上次分级' imageMso='Undo' onAction='Run' supertip='恢复本次Excel会话中最近一批分级及隐藏状态；不恢复金额公式。插入、删除、移动行列后请勿使用。'/>
    </group>
    <group id='grpSum' label='智能求和'>
      <button id='smartSum' label='智能求和' size='large' imageMso='AutoSum' onAction='Run' screentip='指定列生成父项SUM公式' supertip='选择列及父子关系来源，逐项预览勾选确认；允许替换现有父项公式，不改未勾选单元格。'/>
      <button id='undoSum' label='撤销上次求和' imageMso='Undo' onAction='Run' supertip='仅撤销本次Excel会话中的最近一批智能求和；后续编辑导致快照变化时拒绝覆盖。'/>
    </group>
    <group id='grpCheck' label='求和校验'>
      <button id='check' label='求和校验' size='large' imageMso='FormulaEvaluate' onAction='Run' screentip='校验父子项金额' supertip='只读检查金额列；不填充SUM公式、不修改金额。结果窗口双击可定位原单元格。'/>
    </group>
    <group id='grpDisplay' label='分级显示'>
      <button id='level1' label='显示一级' onAction='Run'/>
      <button id='level2' label='显示二级' onAction='Run'/>
      <button id='level3' label='显示三级' onAction='Run'/>
      <button id='expand' label='全部展开' imageMso='OutlineShowAllLevels' onAction='Run'/>
    </group>
    <group id='grpProject' label='项目提取'>
      <button id='projectExtract' label='水电工程提项目' size='large' imageMso='CopyToFolder' onAction='Run' screentip='水电工程提项目' supertip='选择要导出的工作表和打印范围，在副本中清理；不修改源工作簿。'/>
    </group>
    <group id='grpAi' label='AI 助手'>
      <button id='aiSettings' label='AI 设置' size='large' getImage='GetAiImage' onAction='Run' supertip='保存模型配置，检测连接、响应耗时和二维表格输出。使用 GKD(提示词,数据区域)；表格模式需要支持动态数组的 Excel。'/>
      <button id='customFormula' label='自定义公式' size='large' getImage='GetCustomFormulaImage' onAction='Run' supertip='管理本地公式的含义、参数和算法定义；新建、编辑、导入导出或由 AI 设计，预览并确认后保存。'/>
      <button id='aiRefresh' label='刷新 AI 结果' imageMso='RefreshAll' onAction='Run' supertip='只重新请求所选 GKD 公式；不会全表刷新。'/>
      <button id='aiCancel' label='取消 AI 任务' imageMso='CancelRequest' onAction='Run'/>
      <button id='aiValues' label='结果转为值' imageMso='PasteValues' onAction='Run' supertip='将所选已完成 GKD 结果及其完整溢出区域转为值，便于分享。'/>
    </group>
    <group id='grpRecovery' label='说明'>
      <button id='help' label='使用说明' imageMso='Help' onAction='Run'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
        }
        private System.Drawing.Image aiImage;
        public System.Drawing.Image GetAiImage(IRibbonControl control)
        {
            if (aiImage == null)
            {
                using (var stream = typeof(OutlineRibbon).Assembly.GetManifestResourceStream("Outline.Addin.Resources.ai-assistant.png"))
                using (var source = System.Drawing.Image.FromStream(stream))
                    aiImage = new System.Drawing.Bitmap(source);
            }
            return aiImage;
        }
        private System.Drawing.Image customImage;
        public System.Drawing.Image GetCustomFormulaImage(IRibbonControl control)
        {
            if (customImage == null)
                using (var stream = typeof(OutlineRibbon).Assembly.GetManifestResourceStream("Outline.Addin.Resources.custom-formula.png"))
                using (var source = System.Drawing.Image.FromStream(stream))
                    customImage = new System.Drawing.Bitmap(source);
            return customImage;
        }
        public void Run(IRibbonControl control)
        {
            Controller.Execute(control.Id);
        }
    }
}
