# 搞快点：Windows 安装与卸载

## 发行目录

完整解压发行包后运行 `安装与卸载.exe`，不要在 ZIP 预览中运行，也不要只移动 EXE。发行包中 EXE 和配置文件已成对改名；安装后保留为 `Outline.Setup.exe`。

```text
安装与卸载.exe
安装与卸载.exe.config
payload/
  x86/
    Outline.xll
    Outline.dna
    Outline.Addin.dll
    Outline.Core.dll
    ExcelDna.Integration.dll
  x64/
    Outline.xll
    Outline.dna
    Outline.Addin.dll
    Outline.Core.dll
    ExcelDna.Integration.dll
```

Windows 桌面版 Excel 2010 / 2013 / 2016 及后续 16.0 系列，.NET Framework 4.8。仅 x86/x64；不支持原生 ARM64 Excel。安装程序为 AnyCPU、`asInvoker`，不要求管理员权限，不使用 HTML/WebView。

## 使用

1. 安装前保存并关闭 Excel。卸载时若检测到当前用户、当前会话仍有 Excel，会先请求正常关闭并等待；若正常关闭超时，安装器会单独提供“强制关闭 Excel 并卸载…”按钮，必须再次确认，未保存内容可能永久丢失。
2. 选择检测到的 Excel。多个版本会列在下拉框中，显示完整路径、Office 主版本和 x86/x64；没有自动发现时点击“选择 EXCEL.EXE…”定位真实程序。
3. 点击“一键安装”。将 x86 和 x64 两份负载复制到 `%LOCALAPPDATA%\ExcelOutlineAddin\1.0\`，只为所选 Excel 的 Office 主版本登记匹配位数的 XLL。不会启动 Excel。
4. 自行打开 Excel，查看“搞快点”功能区。
5. 卸载可使用发行包中的安装程序，或 Windows“应用/程序和功能”中的“搞快点”。固定目录保存了 `Outline.Setup.exe`，因此删除原发行包后仍可卸载。
6. 切换 Excel 版本/位数或重新安装时，先卸载再安装。发现未完成的安装时保留日志，点击卸载清理后重试。

## 注册策略与安全边界

- 探测 HKCU / HKLM 两个注册表视图中的 `App Paths\EXCEL.EXE`，并补充各 Office 版本的 Excel InstallRoot 与 Click-to-Run 安装路径。读取真实 `EXCEL.EXE` 的 PE machine (`0x014c` / `0x8664`) 识别 Excel 位数；从文件主版本识别 12.0 / 14.0 / 15.0 / 16.0。**不会把 Windows 位数当成 Excel 位数。**安装前重新检查 PE。
- 使用 HKCU `Software\Microsoft\Office\<所选版本>\Excel\Options` 中首个空闲的 `OPEN`、`OPEN1`、`OPEN2`…，写入 `/R "完整安装路径\x86或x64\Outline.xll"`。不覆盖其他槽位，不调整其他加载项顺序。
- 这是 Excel-DNA 原生 XLL，并非 COM Add-in；使用 Excel 的 OPEN 启动加载机制，不通过启动 Excel 的 COM 自动化执行 `AddIns.Add`，避免打开用户工作簿、隐式创建 Excel 实例和额外版本/权限副作用。
- 每次写注册表前，先把 Office 版本、注册表视图、槽位名称与完整预期值持久化到 `install-state.xml`，写入后读回核对。复制文件也使用写前归属记录和 SHA-256 核对。状态通过 XmlSerializer 保存，禁用 DTD/外部 XML 实体。
- 卸载仅在值仍与持久化预期字符串**逐字完全匹配**时删除对应槽位。被用户/其他软件改写的值保留；不重排其他槽位。卸载不根据模糊路径搜索全注册表。
- 控制面板卸载信息仅写 HKCU；卸载时同样核对自己的 `UninstallString`。不写 HKLM。操作系统位数仅用于选取控制面板卸载键视图，不用于推断 Excel 位数。
- 安装目录中缺少状态记录但存在未知内容时拒绝覆盖。卸载只删除白名单内并记录归属的文件，未知文件保留；拒绝目录联接/符号链接路径。失败时保留状态以便重试，不声称已经完成。
- Windows 无法删除正在运行的 EXE。固定目录的安装程序会将自身和配置复制到独立 `%TEMP%\ExcelOutline-Uninstall-<随机标识>`，等待原进程退出，再显示卸载维护窗口；这样能够删除固定目录内的 EXE。临时帮助程序及配置可能保留于用户 TEMP，不影响加载项卸载，可由 Windows 磁盘清理删除。
- **不改变宏设置、信任中心、受信任位置或全盘信任规则；不移除文件来源标记，不绕过企业策略。**未签名 XLL、网络下载标记、组织策略可能阻止加载，请咨询 IT。安装成功仅说明文件与用户级登记就绪，不等同于企业策略允许执行或 Excel 内功能已验证。

## 构建与独立测试

源码：`src/Outline.Setup/`。SDK 风格项目，net48 WinForms，引用 `Microsoft.NETFramework.ReferenceAssemblies.net48` NuGet 1.0.3，可在 macOS/Linux 使用 .NET SDK 交叉编译。

```sh
dotnet build src/Outline.Setup/Outline.Setup.csproj -c Release --nologo
dotnet run --project src/Outline.Setup/Tests/Setup.Tests.csproj -c Release
```

输出：`src/Outline.Setup/bin/Release/net48/Outline.Setup.exe` 与 `.exe.config`。发布时将这两个文件与上述完整 `payload` 放在同一级。

独立测试为 net8 控制台，直接编译生产 `CoreUtil.cs`、`RegistrationLogic.cs`、`InstallState.cs`；不访问 Windows 注册表、不启动 Excel。覆盖 OPEN 首空槽分配、其他加载项保留、精确归属删除、被改写值保护、写前状态落盘、状态落盘失败时不写注册表、带空格/中文路径、PE x86/x64/不支持架构、相对路径越界、XML 状态保存/替换与 DTD 拒绝。

**验证限制：当前开发机为 macOS；已完成 net48 交叉编译及可跨平台纯逻辑测试，但未在 Windows/Excel 中运行 WinForms、真实注册表安装卸载、控制面板入口或高 DPI 多显示器测试。正式分发前必须在干净 Windows 用户下分别用 x86/x64 Excel 验证安装、启动、已有 OPEN 值保护、重装、文件锁定、卸载入口和企业策略提示。**
