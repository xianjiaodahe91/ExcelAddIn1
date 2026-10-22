# Excel 工作表导航

> Excel 侧边导航栏 —— 基于 VSTO + WPF 的加载项，让每个工作表都只差一次点击。

[English](README.md) | **简体中文**

## 简介

工作表一多，标签栏就变成了一场"横向找东西"的运动：标签被压窄、隐藏的表干脆看不到，想切到某张表只能沿着窗口底部来回找。这个加载项在 Excel 窗口左侧停靠一个紧凑面板，把所有工作表列出来，切换、改名、隐藏、删除都不用再去碰底部标签栏。

面板用 WPF 构建，承载在标准的 Office 任务窗格里，因此和 Excel 自带窗格一样停靠在左侧。

## 功能

- **工作表列表** —— 当前工作簿的全部工作表，活动表高亮，隐藏表淡显
- **搜索** —— 输入即筛选
- **切换** —— 点一行就激活对应工作表
- **显隐** —— 每行右侧的眼睛图标可直接隐藏 / 取消隐藏
- **重命名 / 删除 / 新建** —— 面板顶部按钮或行右键菜单
- **标签颜色** —— 读取 Excel 的工作表标签色，显示为每行左侧的彩条
- **设置页** —— 控制是否列出隐藏表，调整面板宽度（220–700 px）
- **记住状态** —— 关闭 Excel 时面板是开着的，下次启动自动打开
- **功能区入口** —— Ribbon 上的「导航」按钮用来开关面板

## 环境要求

| 组件 | 版本 |
|---|---|
| Excel | 2013 及以上（32 位 / 64 位均可） |
| .NET Framework | 4.7.2 及以上 |
| VSTO 运行时 | Microsoft Visual Studio 2010 Tools for Office Runtime |
| 编译环境 | Visual Studio 2022，需勾选「Office/SharePoint 开发」工作负载 |

## 编译

```text
git clone <仓库地址>
```

1. 用 Visual Studio 2022 打开 `ExcelAddIn1.sln`
2. 解决方案配置切到 **Release**
3. 生成 → 重新生成解决方案（`Ctrl+Alt+F7`）
4. 产物在 `ExcelAddIn1\bin\Release`：

```text
ExcelAddIn1.dll
ExcelAddIn1.dll.manifest
ExcelAddIn1.vsto
ExcelAddIn1.pdb
Microsoft.Office.Tools.Common.v4.0.Utilities.dll
```

### 签名

私钥**不在仓库里**，签名配置也不在——项目默认**关闭签名**，所以任何人克隆下来都能直接编译、调试。

要产出可安装的签名包，在 `ExcelAddIn1.csproj` 同级目录放一个本地的 `Signing.props`（已被 git 忽略）：

```xml
<Project>
  <PropertyGroup>
    <SignManifests>true</SignManifests>
    <ManifestKeyFile>$(MSBuildProjectDirectory)\ExcelAddIn1_SigningKey.pfx</ManifestKeyFile>
    <ManifestCertificateThumbprint>你的证书指纹</ManifestCertificateThumbprint>
  </PropertyGroup>
</Project>
```

请使用**你自己的证书**——可以用 `New-SelfSignedCertificate -Type CodeSigningCert` 生成，也可以在 VS 的「项目属性 → 签名」里选择现有 PFX，它会自动填好这两个值。别忘了把配套的 `.cer` 随安装包一起分发。

没有 `Signing.props` 时构建出的是未签名清单：开发够用，但装到不信任该发布者的机器上会被 Office 拒绝。

## 安装到其他电脑

加载项通过 ClickOnce 发布并带签名，目标机需要先信任这个发布者：

```powershell
Import-Certificate -FilePath .\ExcelAddIn1.cer -CertStoreLocation Cert:\CurrentUser\Root
Import-Certificate -FilePath .\ExcelAddIn1.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
```

然后在发布目录里双击 `setup.exe`。不导入证书的话，Excel 会拒绝加载，并提示"用于签署该应用程序的证书不受信任"。

`ExcelAddIn1.cer` 是签名证书的**公钥部分**。它不在仓库里——它跟着“签名的人”走，所以发布时从自己的私钥导出，和安装包一起放到 Release 里即可。

## 配置

用户配置保存在：

```text
%LOCALAPPDATA%\ExcelAddIn1\settings.ini
```

| 键 | 含义 |
|---|---|
| `PaneWidth` | 面板宽度（像素，限制在 220–700） |
| `ShowHiddenSheets` | 是否在列表中显示隐藏的工作表 |
| `LastPaneVisible` | 上次关闭 Excel 时面板是否开着 |

## 项目结构

```text
ExcelAddIn1.sln
ExcelAddIn1/
  ThisAddIn.cs            加载项入口：Excel 事件、面板生命周期、工作表操作
  PaneHost.cs             WinForms 外壳，把 WPF 面板装进任务窗格
  RibbonNavigation.cs     Ribbon 上的开关按钮
  UI/                     WPF 界面层（不引用任何 Excel 类型）
    NavPaneControl.cs     根控件：左侧图标栏 + 右侧页面
    SheetListPage.cs      工作表列表、搜索框、顶部操作按钮
    SettingsPage.cs       设置界面
    SheetInfo.cs          工作表视图模型
    PaneOptions.cs        读写 settings.ini
    UiTheme.cs            配色、字号、图标与控件样式
    UiHelpers.cs
    UiPrompt.cs           重命名 / 确认对话框
  Properties/
```

## 说明与限制

- 任务窗格的**标题栏**（显示"工作表导航"以及折叠、关闭按钮的那一行）由 Office 绘制。VSTO 没有提供修改其字体、字号、样式的接口。
- 面板外壳是纯 WinForms，随 Excel 一起创建，所以能和 Excel 同时出现；WPF 内容稍后才挂上去，这样不会把 WPF 的启动开销加到 Excel 的启动路径上。
- Excel 不允许隐藏当前活动工作表，加载项在需要时会先切换到其它可见工作表再执行隐藏。
