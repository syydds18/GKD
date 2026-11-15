from pathlib import Path
import unittest
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]

class CustomFormulaIntegration(unittest.TestCase):
    def test_ribbon_and_dispatch(self):
        source = (ROOT / 'src/Outline.Addin/Ribbon.cs').read_text()
        xml = source.split('return @"', 1)[1].split('";', 1)[0]
        root = ET.fromstring(xml)
        group = next(e for e in root.iter() if e.attrib.get('id') == 'grpAi')
        button = next((e for e in group if e.attrib.get('id') == 'customFormula'), None)
        self.assertIsNotNone(button, 'custom formula manager button is in AI group')
        assert button is not None
        self.assertEqual(button.attrib.get('label'), '自定义公式')
        self.assertTrue(button.attrib.get('imageMso') or button.attrib.get('getImage'))
        self.assertEqual(button.attrib.get('onAction'), 'Run')
        bridge = (ROOT / 'src/Outline.Addin/ExcelBridge.cs').read_text()
        action = bridge.index('action=="customFormula"')
        self.assertLess(action, bridge.index('if(app.ActiveWorkbook==null)'))
        self.assertIn('new GkdCustomFormulaForm()', bridge)
        self.assertIn('customFormulaForm.CloseForShutdown()', bridge)

    def test_native_manager_safety_boundary(self):
        path = ROOT / 'src/Outline.Addin/GkdCustomFormulaForm.cs'
        self.assertTrue(path.is_file(), 'native custom formula manager exists')
        source = path.read_text()
        self.assertIn('GkdCustomFormulaForm : Form', source)
        for label in ['新建', '编辑', '删除', '导入', '导出', 'AI 设计', '预览并确认保存']:
            self.assertIn('Ui.Button("' + label + '"', source)
        self.assertIn('GkdCustomFormulaStore', source)
        self.assertIn('TextArea(true)', source)
        self.assertIn('ReadOnly = readOnly', source)
        self.assertIn('Description', source)
        self.assertIn('FormulaExpression', source)
        for unsafe in ['CSharpCodeProvider', 'CompileAssemblyFromSource', 'Assembly.Load', 'Process.Start', 'Application.Run', 'ExcelBridge.App', '.Formula =']:
            self.assertNotIn(unsafe, source)
        self.assertIn('确认保存', source)
        self.assertIn('MessageBoxButtons.YesNo', source)

if __name__ == '__main__':
    unittest.main()
