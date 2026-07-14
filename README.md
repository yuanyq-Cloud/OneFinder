# OneFinder — OneNote 全文搜索工具

轻量级 OneNote 插件，遍历所有笔记本页面进行全文搜索，不依赖 WSearch 索引，从而防止因内容未索引而造成的搜索遗漏。

A lightweight OneNote add-in that performs full-text search by traversing all pages across all notebooks, without relying on the Windows Search index, to avoid search omissions caused by unindexed content.

## 界面预览

<img src="UI.png" width="600" alt="OneFinder 界面预览">

## 前提条件

**安装.msi（用户）**

- Windows 10/11 x64
- 已安装 Microsoft OneNote 或 Microsoft 365 OneNote 桌面版（OneNote COM 服务器必须存在）
- .NET 8 Desktop Runtime（x64） — 若未安装需从 Microsoft 下载

**开发 / 构建（开发者）**

- Visual Studio 2022+ 或 MSBuild 17+（用于从源码编译和发布）
- .NET SDK 8.x（用于 `dotnet build` / `dotnet publish`）

## 构建

优先使用仓库根目录下的一键脚本 `build.ps1`（会完成 AddIn 的 MSBuild 构建、主程序的 `dotnet publish`，以及使用 WiX 打包 MSI）。


## 使用

1. 工具栏”开始”选项卡中找到OneFinder工具栏，点击”全文搜索”<br>
<img src=”UI-2.png” width=”400” alt=”OneFinder 界面预览”>

2. 在搜索框输入关键词，按 Enter 或点击”搜索”
3. 等待扫描完成（底部状态栏显示当前扫描进度）
4. 双击结果列表中的条目，OneNote 会自动跳转到对应页面

### 更新：**最近修改记录**

首次打开 OneFinder，或搜索框为空时点击搜索，会列出最近修改的页面。预览中“Def.”占位文本会被忽视。

## 注意事项

- 回收站中的页面、受密码保护的页面会被自动跳过
- 同一页最多显示5条匹配结果 [5/5]
- 笔记本越多、页面越多搜索越慢，关键词仅支持完全匹配
- 单个笔记本页面过多时，搜索期间OneNote可能会短暂未响应（由于 OneNote COM API 的架构限制，OneFinder 必须逐页调用 `GetPageContent()` 由 OneNote 主进程同步处理）

## 项目结构

```
<repo-root>/
├── README.md
├── build.ps1
├── nuget.config
├── OneFinder.sln
├── installer/
│   ├── Package.wxs
│   └── OneFinderSetup.wixpdb
├── OneFinder/
│   ├── OneFinder.csproj           # net8.0-windows, x64
│   ├── Program.cs
│   ├── MainForm.cs
│   ├── MainForm.Designer.cs
│   ├── OneNoteService.cs
│   ├── USER_GUIDE.md
│   └── CHANGELOG.md
└── OneFinder.AddIn/
    ├── OneFinder.AddIn.csproj     # .NET Framework 4.8 add-in for OneNote
    ├── AddIn.cs
    ├── Ribbon.xml
    └── bin/                       # build outputs for add-in (net48)
```

## 开发者可调参数

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