#!/usr/bin/env python3
"""
酒店CAD分析器 - Windows启动器
独立可执行，无需Python环境
"""
import sys
import os
import subprocess
from pathlib import Path
from tkinter import Tk, messagebox, filedialog

def resource_path(relative_path):
    """获取资源路径（兼容PyInstaller打包后）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

def main():
    # 创建隐藏主窗口
    root = Tk()
    root.withdraw()
    
    # 显示欢迎信息
    messagebox.showinfo(
        "酒店CAD分析器 v1.0",
        "欢迎使用酒店CAD分析器！\n\n"
        "功能：\n"
        "• 自动识别房型（大床/双床/套房）\n"
        "• 统计房间数量\n" 
        "• 分析开门方向\n"
        "• 推荐插卡取电位置\n"
        "• 生成Excel报表\n\n"
        "点击确定选择CAD文件..."
    )
    
    # 选择文件
    file_path = filedialog.askopenfilename(
        title="选择CAD文件",
        filetypes=[
            ("DXF文件", "*.dxf"),
            ("所有文件", "*.*")
        ]
    )
    
    if not file_path:
        return 0
    
    # 检查文件格式
    if not file_path.lower().endswith('.dxf'):
        messagebox.showwarning(
            "格式提示",
            "当前版本只支持 .dxf 格式\n"
            "如果是 .dwg 文件，请先用AutoCAD转换"
        )
    
    # 创建输出目录
    output_dir = Path(file_path).parent / "分析报告"
    
    # 运行分析
    try:
        # 直接导入分析模块
        sys.path.insert(0, resource_path('.'))
        import analyze_cad as analyzer
        
        rooms = analyzer.analyze_cad_file(file_path)
        
        if rooms:
            os.makedirs(output_dir, exist_ok=True)
            base_name = Path(file_path).stem
            excel_path = output_dir / f"{base_name}_房型分析.xlsx"
            txt_path = output_dir / f"{base_name}_分析摘要.txt"
            
            analyzer.generate_excel_report(rooms, str(excel_path))
            analyzer.generate_summary_text(rooms, str(txt_path))
            
            # 统计
            stats = {}
            for room in rooms:
                rt = room['room_type']
                stats[rt] = stats.get(rt, 0) + 1
            
            msg = f"✅ 分析成功！\n\n"
            msg += f"共识别 {len(rooms)} 个房间\n\n"
            msg += "房型统计：\n"
            for rt, count in sorted(stats.items()):
                msg += f"  • {rt}: {count}间\n"
            msg += f"\n报表已保存到:\n{output_dir}"
            
            messagebox.showinfo("分析完成", msg)
            
            # 打开输出目录
            os.startfile(str(output_dir))
        else:
            messagebox.showerror(
                "分析失败", 
                "未能识别到房间信息\n\n"
                "请检查CAD文件中是否有房型标注\n"
                "（如'大床房'、'双床房'等文字）"
            )
    except Exception as e:
        messagebox.showerror(
            "运行错误",
            f"程序出错:\n{str(e)}\n\n"
            f"请确保CAD文件格式正确"
        )
    
    return 0

if __name__ == "__main__":
    sys.exit(main())