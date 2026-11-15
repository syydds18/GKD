from pathlib import Path
import unittest,xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
class Integration(unittest.TestCase):
 def test_ribbon(self):
  text=(ROOT/'src/Outline.Addin/Ribbon.cs').read_text()
  xml=text.split('return @"',1)[1].split('";',1)[0]
  tree=ET.fromstring(xml)
  ids={e.attrib.get('id') for e in tree.iter()}
  self.assertTrue({'aiSettings','aiRefresh','aiCancel','aiValues'}<=ids)
  self.assertNotIn('aiWorkflowTest',ids)
 def test_ribbon_order_custom_icon_and_publisher(self):
  text=(ROOT/'src/Outline.Addin/Ribbon.cs').read_text()
  xml=text.split('return @"',1)[1].split('";',1)[0]
  tree=ET.fromstring(xml)
  ids=[e.attrib.get('id') for e in tree.iter() if e.tag.endswith('group')]
  self.assertEqual(ids[-3:],['grpProject','grpAi','grpRecovery'])
  ai=next(e for e in tree.iter() if e.attrib.get('id')=='aiSettings')
  self.assertEqual(ai.attrib.get('getImage'),'GetAiImage')
  self.assertNotIn('imageMso',ai.attrib)
  self.assertTrue((ROOT/'src/Outline.Addin/Resources/ai-assistant.png').is_file())
  for p in ['src/Outline.Addin/Outline.Addin.csproj','src/Outline.Setup/Outline.Setup.csproj']:
   self.assertEqual(ET.parse(ROOT/p).find('.//Company').text,'不会开花的草')
  self.assertIn('SetValue("Publisher", "不会开花的草",',(ROOT/'src/Outline.Setup/Installer.cs').read_text())
  self.assertNotIn('恢复上次分级仅恢复最近一批',(ROOT/'src/Outline.Addin/ExcelBridge.cs').read_text())
 def test_dispatch(self):
  text=(ROOT/'src/Outline.Addin/ExcelBridge.cs').read_text()
  self.assertIn('new GkdAiSettingsForm()',text)
  self.assertIn('GkdAiActions.Run(action)',text)
  self.assertNotIn('new AiWorkflowForm()',text)
 def test_release_dependency(self):
  tree=ET.parse(ROOT/'src/Outline.Addin/Outline.Addin.csproj')
  self.assertEqual(tree.find('.//TargetFramework').text,'net48')
  self.assertTrue(any(e.attrib.get('Include')=='Newtonsoft.Json' for e in tree.iter('PackageReference')))
  self.assertTrue(any(e.attrib.get('Remove')=='AiWorkflow*.cs' for e in tree.iter('Compile')))
  self.assertEqual(ET.parse(ROOT/'src/Outline.Addin/Outline.Addin.csproj').find('.//Version').text,'1.6.0')
  self.assertIn('Newtonsoft.Json.dll',(ROOT/'tools/build.py').read_text())
 def test_installed_dependency_and_lifecycle(self):
  text=(ROOT/'src/Outline.Setup/Installer.cs').read_text()
  required=text.split('RequiredFiles = {',1)[1].split('};',1)[0]
  self.assertIn('"Newtonsoft.Json.dll"',required)
  bridge=(ROOT/'src/Outline.Addin/ExcelBridge.cs').read_text()
  self.assertIn('GkdAiWelcome.Start()',bridge)
  self.assertIn('GkdAiWelcome.Stop()',bridge)
  self.assertIn('GkdAiTableFormatting.Start()',bridge)
  self.assertIn('GkdAiTableFormatting.Stop()',bridge)
 def test_settings_notify_and_manual_refresh(self):
  text=(ROOT/'src/Outline.Addin/GkdAiSettingsForm.cs').read_text()
  self.assertTrue(text.count('GkdAiFunctions.ConfigurationChanged();') == 3)
  self.assertIn('new GkdAiHelpForm()',text)
  self.assertIn('刷新 AI 结果',(ROOT/'src/Outline.Addin/GkdAiHelpForm.cs').read_text())
if __name__=='__main__':unittest.main()
