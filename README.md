# ZoteroProLinker
****
###概览：使Word也可以像Latex一样，一键为引文添加跳转至文献处的超链接。增加编辑印象分，降低桌拒率（也许）
### 🌟 核心功能

* **适用于任何引用格式**：完全脱离正则表达式，基于 Zotero 底层 JSON 结构，完美兼容各类复杂排版。
* **多文献独立跳转**：多条引注（如 `(Author A 2024; Author B 2025)`）可实现点击各自的年份跳转至对应位置。
* **学术排版模式**：支持“括号内全链接（JDE风格）”与“年份专属链接（EJ/QJE风格）”两种排版模式。
* **好看**：彻底解决链接点击变紫的问题，提供一键格式清洗。

### 🚀 安装指南

1. 下载本项目中的 `.dotm` 文件。
2. 将文件移动到 Word 的启动文件夹：
* **Windows**: `%appdata%\Microsoft\Word\STARTUP`
* **Mac**: `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word`


3. 重启 Word，你将在顶部菜单栏看到的Link_Engine工具选项卡。

### 🛠 使用方法
* **排版设置**：必须先设置超链接颜色（蓝/黑/红）及排版模式（全部文字/仅年份）。
* **生成超链**：点击按钮，一键为全篇 Zotero 引注，添加跳转至文献处的超链接。
* **卸载链接**：点击删除超链接，彻底移除所有超链，并将格式重置为纯净文本。

### ⚠️ 注意事项
* 该插件基于 Word VBA 开发，需在 Windows 或 macOS 环境下使用。
* 在增删文献时，建议使用 Zotero 插件自带的 `Refresh` 刷新文档，以确保引注数据结构完整。

### 📜 许可协议

本项目采用 [MIT License](https://www.google.com/search?q=LICENSE)，欢迎交流学术科研工具。
