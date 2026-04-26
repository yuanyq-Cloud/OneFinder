# OneFinder — OneNote 全文搜索工具

一个轻量 Windows 桌面工具，通过 **OneNote COM API** 直接遍历所有笔记本页面并进行全文搜索，无需 Windows Search 索引。

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

1. 工具栏“开始”选项卡中找到OneFinder工具栏，点击"全文搜索"<br>
<img src="UI-2.png" width="400" alt="OneFinder 界面预览">

2. 在搜索框输入关键词，按 Enter 或点击"搜索"
3. 等待扫描完成（底部状态栏显示当前扫描进度）
4. 双击结果列表中的条目，OneNote 会自动跳转到对应页面

## 注意事项

- 受密码保护的节会被自动跳过
- 回收站中的页面同样被跳过
- 笔记本越多、页面越多，首次扫描越慢；建议关键词尽量精确
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
