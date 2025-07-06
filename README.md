# Bilibili Data Tool

![Python Version](https://img.shields.io/badge/Python-3.13%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

### ❗ 免责声明
# 本工具仅用于学习和研究目的。请遵守Bilibili的使用条款，不要滥用此工具。开发者对使用此工具造成的任何后果不承担责任。

Bilibili Data Tool 是一个强大的桌面应用程序，用于爬取Bilibili UP主的视频信息并导出为Excel文件。它提供了直观的用户界面，让用户能够轻松获取UP主的视频列表（BVID）并导出详细的视频信息。

## ✨ 功能特点

- **BVID爬取**：输入UP主UID，自动爬取其所有视频的BVID列表
- **视频信息导出**：将BVID列表导出为包含详细信息的Excel文件
- **自定义选项**：可选择包含视频封面、描述和标签
- **用户友好界面**：直观的标签页设计，操作简单
- **日志记录**：详细记录操作过程和状态信息
- **配置保存**：自动保存上次使用的配置项
- **错误处理**：提供详细的错误信息，便于排查问题

## 🚀 安装与使用

### 前提条件

1. 安装 [Python 3.8+](https://www.python.org/downloads/)
2. 安装 [Microsoft Edge](https://www.microsoft.com/edge)
3. 下载匹配的 [Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/) 
4. 将下载的 `msedgedriver.exe` 放置到 `C:\` 根目录

### 🛠 安装步骤

1. 克隆仓库：
   ```bash
   git clone https://github.com/yourusername/bilibili-data-tool.git
   cd bilibili-data-tool
   ```
2. 创建并激活虚拟环境：
   ```bash
   python -m venv venv
   source venv/bin/activate  # 在Windows上使用 `venv\Scripts\activate`
   ```
3. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
4. 运行程序：
   ```bash
   python 爬虫.py
   ```
### 📖 使用说明
# 一. BVID爬取
1. 在"BVID爬取"标签页输入目标UP主的UID
2. （可选）在Cookie区域输入自定义Cookie（用于访问需要登录的内容）
3. 点击"开始爬取"按钮
4. 爬取完成后，点击"保存BVID"将结果保存为文本文件
# 二. 视频信息导出
1. 在"视频信息导出"标签页选择BVID文件（或使用爬取结果）
2. 设置输出文件名（默认为bilibili_videos.xlsx）
3. 选择导出选项（包含封面、描述、标签）
4. 点击"开始导出"按钮
5. 导出完成后，点击"打开文件夹"查看结果

### ⚠ 注意事项
# 一. ​WebDriver要求​：
1. 必须下载匹配Edge浏览器版本的WebDriver
2. 将下载的msedgedriver.exe放置到C:\根目录
3. 确保Edge浏览器已安装并更新到最新版本
# 二. ​Cookie使用​：
1. 对于需要登录才能访问的内容，可以输入自定义Cookie
2. 获取Cookie的方法：登录Bilibili后，在开发者工具中复制Cookie值
# 三. ​爬取限制​：
1. 为避免对Bilibili服务器造成过大压力，请合理使用本工具
2. 爬取过程中有随机延迟，防止请求过快被限制

