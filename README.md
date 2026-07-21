# OneFinder — OneNote 无遗漏搜索工具

轻量级 OneNote 插件，遍历所有笔记本页面进行全文搜索，不依赖屎一样的 WSearch 索引，防止因内容未索引而造成的搜索遗漏。

A lightweight OneNote add-in that performs full-text search by traversing all pages across all notebooks, without relying on the Windows Search index, to avoid search omissions caused by unindexed content.

## 界面预览

<img src="UI.png" width="600" alt="OneFinder 界面预览">

## 前提条件

**安装（用户）**

- Windows 10/11 x64
- 已安装 Microsoft OneNote 或 Microsoft 365 OneNote 桌面版（OneNote COM 服务器必须存在）
- .NET 8 Desktop Runtime（x64） — 若未安装需从 Microsoft 下载
- .NET Framework 4.8 (通常系统已经预装）
- 运行 OneFinderSetup.exe

**开发 / 构建（开发者）**

- Visual Studio 2022+ 或 MSBuild 17+（用于从源码编译和发布）
- .NET SDK 8.x（用于 `dotnet build` / `dotnet publish`）

## 用户手册

1. 工具栏”开始”选项卡中找到OneFinder工具栏，点击”全文搜索”<br>
<img src=”UI-2.png” width=”400” alt=”OneFinder 界面预览”>

2. 在搜索框输入关键词，按 Enter 或点击”搜索”
3. 等待扫描完成（底部状态栏显示当前扫描进度）
4. 双击结果列表中的条目，OneNote 会自动跳转到对应页面

### 更新：**最近修改记录**

首次打开 OneFinder，或搜索框为空时点击搜索，会列出最近修改的页面。预览中“Def.”占位文本会被忽视。

### 注意事项

- 回收站中的页面、受密码保护的页面会被自动跳过
- 同一页最多显示5条匹配结果 [5/5]
- 笔记本越多、页面越多搜索越慢，关键词仅支持完全匹配
- 单个笔记本页面过多时，搜索期间OneNote可能会短暂未响应（由于 OneNote COM API 的架构限制，OneFinder 必须逐页调用 `GetPageContent()` 由 OneNote 主进程同步处理）

## 项目结构

```
<repo-root>/
├── README.md
├── build.ps1                     # 一键构建脚本（必须用）
├── nuget.config
├── OneFinder.sln
├── installer/
│   ├── Package.wxs               # WiX MSI 定义
│   └── Setup/                    # 引导程序项目
│       ├── OneFinder.Setup.csproj
│       ├── Program.cs            # 语言选择 → 启动 msiexec
│       └── LanguageDialog.cs     # 语言选择对话框
├── OneFinder/                    # 主程序 net8.0-windows x64
│   ├── OneFinder.csproj
│   ├── Program.cs
│   ├── Loc.cs                    # 本地化帮助类
│   ├── Strings.resx              # 中文资源（中性/默认）
│   ├── Strings.en-US.resx        # 英文资源
│   ├── MainForm.cs
│   ├── OneNoteService.cs
│   └── OneNoteScheduler.cs
└── OneFinder.AddIn/              # .NET Framework 4.8 COM AddIn
    ├── OneFinder.AddIn.csproj
    ├── AddIn.cs
    ├── Ribbon.xml                # 中文 Ribbon 按钮
    └── Ribbon.en-US.xml          # 英文 Ribbon 按钮
```

## 构建

优先使用仓库根目录下的一键脚本 `build.ps1`：

```powershell
.\build.ps1
```

按顺序执行 4 步：

| 步骤 | 命令 | 产出 |
|------|------|------|
| 1 | MSBuild `OneFinder.AddIn\OneFinder.AddIn.csproj` | `OneFinder.AddIn.dll`（含嵌入式 Ribbon XML 资源） |
| 2 | dotnet publish `OneFinder\OneFinder.csproj` | 发布到 `publish\`（含 `en-US\` 卫星程序集） |
| 3 | wix build `installer\Package.wxs` | `installer\OneFinderSetup.msi` |
| 4 | MSBuild `installer\Setup\OneFinder.Setup.csproj` | `OneFinderSetup.exe`（嵌入 MSI 的引导程序） |

最终产出：**`OneFinderSetup.exe`** — 用户下载运行的唯一文件。

### 构建注意事项

**不要手动逐步构建**。手动构建极易遗漏步骤，导致以下问题：

| 陷阱 | 症状 |
|------|------|
| 未重新编译 AddIn DLL | OneNote 中 OneFinder 图标消失（Ribbon XML 资源版本不匹配） |
| 未重新生成卫星程序集 | 选择英文后界面仍显示中文（`en-US\OneFinder.resources.dll` 未打包进 MSI） |
| 修改 `.resx` 后只 build 不 publish | 同上（`publish\en-US\` 目录不会自动更新） |
| 使用 `dotnet build AddIn` 代替 MSBuild | 编译失败——COM 引用项目必须用 .NET Framework MSBuild |
| 修改 .resx key 名称后只改一处 | `Loc.Get("Key")` 的 key 必须与 `Strings.resx` 和 `Strings.en-US.resx` 三处一致 |

**修改版本号：** 版本号由仓库根目录的 `Directory.Build.props` 统一管理，所有 `.csproj` 项目自动继承。修改时只需编辑该文件中的 `<Version>`、`<FileVersion>`。但 WiX 安装包 `installer\Package.wxs` **不读 MSBuild 属性**，以下两处需手动同步：
- 第 8 行：`Version="x.x.x"` — MSI 包版本
- 第 158、164 行：`Version=x.x.x.x` — COM 注册表中的 AddIn 程序集版本

## 开发者注意事项

### 可调参数

以下常量分散在各源文件中，调整后重新编译即可生效，无需改动业务逻辑：

| 参数 | 位置 | 说明 |
|------|------|------|
| 最近修改页面数量 | `MainForm.cs` → `LoadRecentPages()` 中 `maxCount: 10` | 空搜索或首次打开时显示的最近页面数 |
| 预览文本最大长度 | `OneNoteService.cs` → `ExtractPagePreview()` 中 `maxLength: 120` | 最近页面行 2 预览的字符数上限 |
| 预览占位文本过滤 | `OneNoteService.cs` → `ExtractPagePreview()` 中的 `text.Equals("Def.", ...)` 判断 | 跳过无意义的占位段落（如仅含 `Def.`），可追加其他忽略词 |
| 预览刷新限流间隔 | `MainForm.cs` → `_previewRefreshTimer` 的 `Interval = 200`（毫秒） | 渐进加载预览时列表重绘的最短间隔，避免闪烁 |
| 搜索结果每页最多匹配数 | `OneNoteService.cs` → `snippets.Count >= 5` | 同一页面在搜索结果中最多显示多少条命中片段 |
| 搜索结果行高 | `MainForm.cs` → `ItemHeight = 88` | 每条搜索结果的高度（像素），影响单页可见行数 |
| 窗口默认尺寸 | `MainForm.cs` → `Size = new Size(950, 990)` | 首次启动或无已保存尺寸时的窗口大小 |
| 窗口尺寸持久化路径 | `MainForm.cs` → `WindowSizeStore.FilePath` | `%LocalAppData%\OneFinder\window.json` |

### i18n
**新增i18n词条的步骤：**
1. 在 `Strings.resx` 中添加中文 `<data name="NewKey" ...>`
2. 在 `Strings.en-US.resx` 中添加同名英文条目
3. 在代码中用 `Loc.Get("NewKey")` 引用（不要硬编码字符串）
4. 运行完整 `build.ps1`

**语言本地化架构：**
```
引导程序 (OneFinderSetup.exe)
  └→ 语言选择对话框 → 写 HKCU\Software\OneFinder\Language
      └→ msiexec 安装 MSI
          └→ OneFinder.exe 启动
              ├→ Loc.Initialize() 读 HKCU → 加载 Strings.xx.resx
              └→ MainForm/OneNoteService 用 Loc.Get() 获取字符串
          └→ OneNote 加载 AddIn
              └→ GetCustomUI() 读 HKCU → 返回 Ribbon.xx.xml
```

**注册表读取优先级：** HKCU → HKLM → zh-CN 默认值
