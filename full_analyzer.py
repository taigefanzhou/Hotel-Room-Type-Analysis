#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hotel CAD Analyzer - Suite Recognition
"""
import sys
import os
import re
import math
from pathlib import Path
from datetime import datetime

# 房型特征库
ROOM_FEATURES = {
    "大床房": ["卫生间", "电视", "衣柜", "迷你吧"],
    "双床房": ["卫生间", "电视", "两张床"],
    "套房": ["客厅", "卧室", "卫生间"],
    "行政套房": ["客厅", "卧室", "办公区"],
    "豪华套房": ["客厅", "主卧", "次卧", "餐厅"],
    "总统套房": ["客厅", "主卧", "次卧", "餐厅", "会议室"],
    "家庭房": ["大床", "小床", "儿童区"],
    "无障碍房": ["无障碍", "轮椅", "扶手"]
}

def analyze_cad_full(filepath):
    import ezdxf
    
    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()
    
    # 收集所有标注
    all_texts = []
    for entity in msp:
        if entity.dxftype() in ['TEXT', 'MTEXT']:
            try:
                text = entity.dxf.text if entity.dxftype() == 'TEXT' else entity.text
                if text:
                    pos = entity.dxf.insert if hasattr(entity.dxf, 'insert') else (0, 0, 0)
                    all_texts.append({
                        'text': text.strip(),
                        'pos': (pos[0], pos[1])
                    })
            except:
                pass
    
    # 去重
    seen = set()
    unique = []
    for item in all_texts:
        key = f"{item['text']}_{round(item['pos'][0], -1)}_{round(item['pos'][1], -1)}"
        if key not in seen and item['text']:
            seen.add(key)
            unique.append(item)
    
    # 识别房间
    rooms = []
    for item in unique:
        text = item['text']
        room_type = None
        
        if re.search(r'大床', text):
            room_type = '大床房'
        elif re.search(r'双床|标间', text):
            room_type = '双床房'
        elif re.search(r'套房|suite|行政|豪华', text.lower()):
            room_type = '套房'
        elif re.search(r'总统', text):
            room_type = '总统套房'
        elif re.search(r'家庭|亲子', text):
            room_type = '家庭房'
        
        if room_type:
            rooms.append({
                'room_type': room_type,
                'original_type': room_type,
                'text': text,
                'pos': item['pos'],
                'room_number': '-',
                'door_direction': '-',
                'card_position': '-',
                'suite_score': 0,
                'suite_reason': ''
            })
    
    # 套房智能识别
    for room in rooms:
        rx, ry = room['pos']
        
        has_living = False
        has_bedroom = False
        has_bathroom = False
        label_count = 0
        
        for item in unique:
            ix, iy = item['pos']
            dist = math.sqrt((rx-ix)**2 + (ry-iy)**2)
            
            if dist < 400:
                label_count += 1
                txt = item['text'].lower()
                
                if any(kw in txt for kw in ['客厅', 'living', '会客']):
                    has_living = True
                if any(kw in txt for kw in ['卧室', 'bedroom', '主卧']):
                    has_bedroom = True
                if any(kw in txt for kw in ['卫生间', 'bathroom', '卫浴']):
                    has_bathroom = True
        
        suite_score = 0
        if has_living: suite_score += 3
        if has_bedroom: suite_score += 2
        if has_bathroom: suite_score += 1
        if label_count > 15: suite_score += 1
        
        room['suite_score'] = suite_score
        
        if suite_score >= 5 or (suite_score >= 3 and has_living and has_bedroom):
            room['room_type'] = '套房'
            room['suite_reason'] = '有客厅+卧室分区'
        elif suite_score >= 3 and has_living:
            room['room_type'] = '套房'
            room['suite_reason'] = '有独立客厅'
        else:
            room['suite_reason'] = '标准单间'
    
    # 门方向
    for room in rooms:
        rx, ry = room['pos']
        
        for entity in msp:
            if entity.dxftype() == 'ARC':
                try:
                    if entity.dxf.layer == 'FF-门':
                        center = entity.dxf.center
                        radius = entity.dxf.radius
                        
                        if radius < 15:
                            continue
                        
                        dist = math.sqrt((center.x-rx)**2 + (center.y-ry)**2)
                        if dist < 200:
                            start = entity.dxf.start_angle
                            end = entity.dxf.end_angle
                            mid = (start + end) / 2
                            
                            if end < start:
                                mid = (start + end + 360) / 2
                                if mid > 360: mid -= 360
                            
                            if 45 <= mid < 135:
                                room['door_direction'] = '向上开启'
                                room['card_position'] = '内侧'
                            elif 135 <= mid < 225:
                                room['door_direction'] = '向左开启'
                                room['card_position'] = '左侧'
                            elif 225 <= mid < 315:
                                room['door_direction'] = '向下开启'
                                room['card_position'] = '内侧'
                            else:
                                room['door_direction'] = '向右开启'
                                room['card_position'] = '右侧'
                except:
                    pass
    
    return rooms

def generate_excel(rooms, output_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    
    wb = Workbook()
    
    ws1 = wb.active
    ws1.title = '房型汇总'
    
    ws1['A1'] = '酒店房型统计报表'
    ws1['A1'].font = Font(size=14, bold=True)
    ws1.merge_cells('A1:C1')
    
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
    
    ws2 = wb.create_sheet('房间明细')
    headers2 = ['序号', '房型', '原始标注', '开门方向', '插卡位置', '套房指数', '识别依据']
    
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    for i, room in enumerate(rooms, 1):
        ws2.cell(row=i+1, column=1, value=i)
        ws2.cell(row=i+1, column=2, value=room['room_type'])
        ws2.cell(row=i+1, column=3, value=room['original_type'])
        ws2.cell(row=i+1, column=4, value=room['door_direction'])
        ws2.cell(row=i+1, column=5, value=room['card_position'])
        ws2.cell(row=i+1, column=6, value=room['suite_score'])
        ws2.cell(row=i+1, column=7, value=room['suite_reason'])
    
    wb.save(output_path)

def main():
    from tkinter import Tk, messagebox, filedialog
    
    root = Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title='选择CAD文件 (.dxf)',
        filetypes=[('DXF文件', '*.dxf'), ('所有文件', '*.*')]
    )
    
    if not file_path:
        return 0
    
    try:
        rooms = analyze_cad_full(file_path)
        
        if not rooms:
            messagebox.showerror('错误', '未识别到房间')
            return 1
        
        desktop = Path.home() / 'Desktop'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = desktop / f'CAD分析报告_{timestamp}'
        output_dir.mkdir(exist_ok=True)
        
        excel_path = output_dir / '房型分析.xlsx'
        generate_excel(rooms, excel_path)
        
        stats = {}
        for r in rooms:
            t = r['room_type']
            stats[t] = stats.get(t, 0) + 1
        
        msg = f'分析完成！\\n\\n共识别 {len(rooms)} 个房间\\n\\n'
        for rt, count in sorted(stats.items()):
            msg += f'{rt}: {count} 间\\n'
        msg += f'\\n报告已保存到桌面:\\n{output_dir}'
        
        messagebox.showinfo('成功', msg)
        
        if sys.platform == 'win32':
            os.startfile(str(output_dir))
        
    except Exception as e:
        messagebox.showerror('错误', f'分析失败:\\n{str(e)}')
        return 1
    
    return 0

if __name__ == '__main__':
    sys.exit(main())
