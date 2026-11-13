# 原 HTML 算法 → C# 对照测试

## 运行

在本目录执行（需要 Node.js 和 .NET 8 SDK）：

```sh
node generate-fixtures.js
dotnet run --project CoreParity.csproj
```

`generate-fixtures.js` 直接 `require()` 随附、未修改的 `ExcelOutlineCore.original.js`，真实执行 `codeFamily`、`toNumberValue`、`detectSheetValues`、`validateSheetSums`，写出 `fixtures.json`。也可传入另一份原源 JS 的绝对路径。JSON 包含输入、全部原始输出及源文件 SHA-256。测试程序校验同目录源 JS 哈希，再逐字段对比 C# 输出。不是手写期望值，也没有通过 UI/浏览器运行算法。

预置覆盖：九种编号、浮点编号还原、同级点号不逐行加深、段内点号父级、无编号低置信度、无表头文本编号兜底、多行表头、金额列筛选、合计/小计跳过、fee/summary候选、最小差候选、0.05精确边界/超界、相对容差、大数字/科学计数法文本、负数/人民币/千分位、Excel错误值/公式、outline优先及fallback、空行/尾空行、500条错误截断，以及固定种子的160张混合随机表。

## Core API

- 项目：`../../src/Outline.Core/Outline.Core.csproj`，netstandard2.0 / C# 7.3，无第三方包。
- `Engine.Detect(SheetData)`：等同 `detectSheetValues`（先裁剪末尾空行）；`Engine.DetectLevels(SheetData)` 是不裁剪的原始变体。
- `Engine.FindHeader(SheetData)`、`Engine.CodeFamily(object)`、`Engine.ToNumberValue(object)` 可独立调用。
- `Engine.Validate(SheetData, Detection)`：等同 `validateSheetSums`。
- `Values` / `Formulas` 的行列键均为 **1-based**；`Detection.Levels` 与 `OutlineLevels` 的值均为 **0..7**。外部 Excel 的 `Row.OutlineLevel` 传入前必须减1，写回时加1。`OutlineMax` 同为减1后的最大层级。
- 公式仅携带展示信息；核心使用 Excel 提供的已计算值，不自行求公式。
- 原始返回字段均有 PascalCase 属性，包括 `ErrorGroups`、`Fallback`。原 JS 未定义的可选字段在 C# 中保留类型默认值；测试检查 JS 实际输出的所有字段，不要求 C# 模拟属性缺失。

## 忠实保留的规则

父子关系从上向下堆栈构建；不会改成“查找下方总计”。子项候选为编号子项、编号子项+费用行、全部非汇总子项，按与父值的绝对差最小者选择，同差保留先出现者。它是原有启发式，不代表财务意义上的唯一正确关系。

容差保留原式 `Math.Max(0.05, Math.Abs(parentValue) * 1e-9)`，仅严格大于容差时错误；未改为 decimal、四舍五入或固定0.05。低置信度早退出也忠实保留：仅有原表outline并不能绕过“未识别到编号层级”。原 JS 的 `#NAME?` 并不被当前错误正则识别，此行为同样保留。

## 已执行结果

Release 编译：0 warnings、0 errors。实际对照输出：

```text
PASS: 188 sheets, 30 code cases, 26 number cases, 48438 assertions; source SHA-256 6c724cbfa750160847925ceb68013cc27fa9ca98fafff855b013bd976a9adf53
```

逐个数值使用 `double == double` 精确比较（不是另设测试容差），字符串、枚举顺序、错误分组、截断计数和所有原 JS 返回字段均对照。测试开发中已实际捕获并修复 C# 与 JS 科学计数法/大整数文本的5处差异。

## 验证范围

已在 macOS .NET 8 运行 C# 核心（编译目标 netstandard2.0、强制 C# 7.3）与源 JS 对照。Windows .NET Framework 4.8、Excel-DNA、实际 Excel 的 COM 读取/写入不在此子任务验证范围。数字转字符串有专门的 JS 兼容最短往返格式处理；极端 IEEE-754 格式化在 Windows Framework 上尚需宿主回归。
