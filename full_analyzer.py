#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hotel CAD Analyzer - Full Chinese Version with Door Direction
"""
import sys
import os
import re
import math
from pathlib import Path
from datetime import datetime

def analyze_cad_full(filepath):
    """完整分析，包含门方向"""
    import ezdxf
    
    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()
    
    # 1. 分析房间标注
    rooms = []
    text_entities = []
    for entity in msp:
        if entity.dxftype() in ['TEXT', 'MTEXT']:
            try:
                text = entity.dxf.text if entity.dxftype() == 'TEXT' else entity.text
                if text:
                    pos = entity.dxf.insert if hasattr(entity.dxf, 'insert') else (0, 0, 0)
                    text_entities.append({
                        'text': text.strip(),
                        'pos': (pos[0], pos[1])
                    })
            except:
                pass
    
    # 去重
    seen = set()
    unique = []
    for item in text_entities:
        key = f"{item['text']}_{round(item['pos'][0], -1)}_{round(item['pos'][1], -1)}"
        if key not in seen and item['text']:
            seen.add(key)
            unique.append(item)
    
    # 识别房型
    for item in unique:
        text = item['text']
        room_type = None
        
        if re.search(r'大床', text):
            room_type = '大床房'
        elif re.search(r'双床|标间', text):
            room_type = '双床房'
        elif re.search(r'套房', text):
            room_type = '套房'
        
        if room_type:
            # 提取房间号
            match = re.search(r'\b(\d{3,4})\b', text)
            room_num = match.group(1) if match else None
            
            rooms.append({
                'room_type': room_type,
                'room_number': room_num or '-',
                'raw_text': text,
                'pos': item['pos'],
                'door_direction': '-',
                'card_position': '-'
            })
    
    # 2. 分析门方向
    doors = []
    for entity in msp:
        if entity.dxftype() == 'ARC':
            try:
                if entity.dxf.layer == 'FF-门':
                    center = entity.dxf.center
                    radius = entity.dxf.radius
                    
                    if radius < 15:  # 过滤小弧线
                        continue
                    
                    start = entity.dxf.start_angle
                    end = entity.dxf.end_angle
                    
                    # 计算中点角度
                    mid = (start + end) / 2
                    if end < start:
                        mid = (start + end + 360) / 2
                        if mid > 360:
                            mid -= 360
                    
                    # 判断方向
                    if 45 <= mid < 135:
                        direction = '向上开启'
                        card_pos = '内侧'
                    elif 135 <= mid < 225:
                        direction = '向左开启'
                        card_pos = '左侧'
                    elif 225 <= mid < 315:
                        direction = '向下开启'
                        card_pos = '内侧'
                    else:
                        direction = '向右开启'
                        card_pos = '右侧'
                    
                    doors.append({
                        'center': (center.x, center.y),
                        'direction': direction,
                        'card_pos': card_pos
                    })
            except:
                pass
    
    # 3. 关联门和房间（按距离最近）
    for room in rooms:
        rx, ry = room['pos']
        closest = None
        min_dist = float('inf')
        
        for door in doors:
            dx, dy = door['center']
            dist = math.sqrt((rx-dx)**2 + (ry-dy)**2)
            if dist < min_dist:
                min_dist = dist
                closest = door
        
        if closest and min_dist < 200:
            room['door_direction'] = closest['direction']
            room['card_position'] = closest['card_pos']
    
    return rooms

def generate_excel(rooms, output_path):
    """生成Excel报表"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    
    wb = Workbook()
    
    # 汇总表
    ws1 = wb.active
    ws1.title = '房型汇总'
    
    ws1['A1'] = '酒店房型统计报表'
    ws1['A1'].font = Font(size=14, bold=True)
    ws1.merge_cells('A1:C1')
    
    # 统计
    stats = {}
    for r in rooms:
        t = r['room_type']
        stats[t] = stats.get(t, 0) + 1
    
    headers = ['房型', '数量', '占比']
    for col, h in enumerate(headers, 1):
        cell = ws1.cell(row=3, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    row = 4
    total = len(rooms)
    for rt, count in sorted(stats.items()):
        ws1.cell(row=row, column=1, value=rt)
        ws1.cell(row=row, column=2, value=count)
        ws1.cell(row=row, column=3, value=f'{count/total*100:.1f}%')
        row += 1
    
    ws1.cell(row=row, column=1, value='合计').font = Font(bold=True)
    ws1.cell(row=row, column=2, value=total).font = Font(bold=True)
    
    # 明细表
    ws2 = wb.create_sheet('房间明细')
    headers2 = ['序号', '房间号', '房型', '开门方向', '插卡位置']
    
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    for i, room in enumerate(rooms, 1):
        ws2.cell(row=i+1, column=1, value=i)
        ws2.cell(row=i+1, column=2, value=room['room_number'])
        ws2.cell(row=i+1, column=3, value=room['room_type'])
        ws2.cell(row=i+1, column=4, value=room['door_direction'])
        ws2.cell(row=i+1, column=5, value=room['card_position'])
    
    wb.save(output_path)

def main():
    from tkinter import Tk, messagebox, filedialog
    
    root = Tk()
    root.withdraw()
    
    # 选择文件
    file_path = filedialog.askopenfilename(
        title='选择CAD文件 (.dxf)',
        filetypes=[('DXF文件', '*.dxf'), ('所有文件', '*.*')]
    )
    
    if not file_path:
        return 0
    
    try:
        # 分析
        rooms = analyze_cad_full(file_path)
        
        if not rooms:
            messagebox.showerror('错误', '未识别到房间，请检查CAD文件中是否有房型标注')
            return 1
        
        # 输出到桌面
        desktop = Path.home() / 'Desktop'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = desktop / f'CAD分析报告_{timestamp}'
        output_dir.mkdir(exist_ok=True)
        
        # 生成Excel
        excel_path = output_dir / '房型分析.xlsx'
        generate_excel(rooms, excel_path)
        
        # 生成文本报告
        txt_path = output_dir / '分析摘要.txt'
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write('酒店房型分析报表\n')
            f.write('=' * 50 + '\n\n')
            f.write(f'总房间数: {len(rooms)} 间\n\n')
            
            stats = {}
            for r in rooms:
                t = r['room_type']
                stats[t] = stats.get(t, 0) + 1
            
            f.write('房型统计:\n')
            for rt, count in sorted(stats.items()):
                f.write(f'  {rt}: {count} 间\n')
            
            f.write('\n房间明细:\n')
            f.write(f'{"序号":<6}{"房号":<10}{"房型":<10}{"开门方向":<12}{"插卡位置":<12}\n')
            f.write('-' * 60 + '\n')
            
            for i, room in enumerate(rooms, 1):
                f.write(f'{i:<6}{room["room_number"]:<10}{room["room_type"]:<10}'
                       f'{room["door_direction"]:<12}{room["card_position"]:<12}\n')
        
        # 显示结果
        msg = f'分析完成！\n\n共识别 {len(rooms)} 个房间\n\n'
        msg += '房型统计:\n'
        for rt, count in sorted(stats.items()):
            msg += f'  {rt}: {count} 间\n'
        msg += f'\n报告已保存到桌面:\n{output_dir}'
        
        messagebox.showinfo('成功', msg)
        
        # 打开文件夹
        if sys.platform == 'win32':
            os.startfile(str(output_dir))
        
    except Exception as e:
        messagebox.showerror('错误', f'分析失败:\n{str(e)}')
        return 1
    
    return 0

if __name__ == '__main__':
    sys.exit(main())
