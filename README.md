# 📊 数据处理中心

> IT 部门数据自动化处理中心 — Web 版拖拽式数据处理工具箱

![纯前端](https://img.shields.io/badge/纯前端-无需后端-brightgreen)
![SheetJS](https://img.shields.io/badge/SheetJS-0.20.3-blue)
![License](https://img.shields.io/badge/License-MIT-yellow)

## ✨ 功能特性

| 工具 | 功能 | 说明 |
|:---:|------|------|
| 📋 **智能匹配** | 两表关联匹配 | 根据关键列自动匹配并填充数据，支持多条填充规则 |
| 🔄 **格式转换** | 格式互转 | Excel ↔ CSV ↔ JSON 一键互转 |
| 🧹 **数据清洗** | 批量清理 | 去重、去空行、去空格、去换行符等 |
| 📊 **报表合并** | 多表合一 | 多个 Excel 文件纵向追加或 Sheet 合并 |
| 🔍 **数据比对** | 差异发现 | 高亮新增、修改、删除的行 |

## 🚀 快速使用

### 方式一：直接打开
双击 `index.html` 即可在浏览器中使用。

### 方式二：本地服务器
```bash
# 使用 Python 启动本地服务器
python3 -m http.server 8899

# 浏览器访问
open http://localhost:8899
```

## 🔒 数据安全

所有数据处理都在**浏览器端**完成，文件不会上传到任何服务器，数据安全有保障。

## 🎨 设计特色

- 🌓 亮色/暗色模式一键切换
- 📱 响应式布局，适配桌面和移动端
- 🎯 拖拽上传 + 点击上传双模式
- 🔍 可搜索的下拉框，快速定位列名
- ✨ 微动画交互，操作反馈即时可见
- 📈 处理结果统计卡片，清晰直观

## 🛠️ 技术栈

- **HTML5** + **CSS3** + **JavaScript (ES6+)**
- **SheetJS** (xlsx.js) — Excel/CSV 读写
- 零依赖，零构建，零部署成本

## 📁 项目结构

```
数据处理中心/
├── index.html    # 主页面（5个工具视图）
├── style.css     # 设计系统 + 组件样式
├── app.js        # 核心逻辑（5个工具模块）
└── README.md     # 项目说明
```

## 📝 License

MIT
