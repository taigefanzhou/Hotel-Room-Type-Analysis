# 酒店CAD平面图分析器

> 自动识别房型、统计房间数量、分析开门方向、生成Excel报表

## 📥 下载使用

### Windows用户
直接下载 [酒店CAD分析器.exe](https://github.com/你的用户名/hotel-cad-analyzer/releases) ，双击使用，无需安装！

### Mac用户
下载 [酒店CAD分析器-Mac.app](https://github.com/你的用户名/hotel-cad-analyzer/releases) ，双击使用。

---

## 🚀 功能

- 🏨 自动识别房型（大床房/双床房/套房等）
- 📊 统计各房型数量及占比
- 🚪 分析门开启方向
- 🔌 推荐插卡取电位置
- 📑 生成Excel报表（含汇总+明细）

---

## 📋 使用方法

1. **打开软件**，点击确定选择CAD文件
2. **选择文件**（.dxf格式）
3. **等待分析**
4. **查看结果**，Excel报表会自动生成

---

## 📁 CAD文件要求

- **格式**：`.dxf`（如为`.dwg`请先用AutoCAD转换）
- **房型标注**：图中需有"大床房"、"双床房"等文字
- **门图层**：门弧线需在 `FF-门` 图层

---

## 📊 输出文件

分析完成后在同目录生成：
```
分析报告/
├── xxx_房型分析.xlsx    # Excel报表
└── xxx_分析摘要.txt     # 文本报告
```

---

## 🛠 技术栈

- Python 3.11
- ezdxf (CAD解析)
- openpyxl (Excel生成)
- PyInstaller (打包EXE)

---

## 📄 许可证

MIT License
版本: v1.0
