# 酒店CAD分析器

自动分析酒店CAD平面图，识别房型、统计数量、生成Excel报表。

## 功能

- 智能识别房型（大床房/双床房/套房/家庭房等）
- 自动过滤汇总统计信息
- 分析开门方向
- 推荐插卡取电位置
- 生成Excel报表（含楼层分布）

## 使用方法

1. 运行程序
2. 选择CAD文件（.dxf格式）
3. 自动生成Excel报告

## 系统要求

- Windows 10/11
- 无需安装Python

## 开发

```bash
pip install ezdxf openpyxl
python analyzer.py
```
