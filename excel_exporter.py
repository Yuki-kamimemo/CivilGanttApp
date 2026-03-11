import json
import re
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ★追加：画面の備考欄（HTML）の<br>や<div>を、Excel用の改行に変換して綺麗にする関数
def strip_html_tags(text):
    if not text:
        return ""
    # <br>や<div>を改行（\n）に変換
    text = re.sub(r'<(br|div|p)[^>]*>', '\n', text, flags=re.IGNORECASE)
    # その他の不要なタグをすべて削除
    text = re.sub(r'<[^>]+>', '', text)
    # 余分な改行を整理
    text = re.sub(r'\n+', '\n', text)
    return text.strip()

def export_data_to_excel(data_str, file_path):
    try:
        export_data = json.loads(data_str)
        state = export_data.get("state", {})
        national_holidays = export_data.get("nationalHolidays", {})
    except Exception as e:
        raise ValueError(f"JSONデータのパースに失敗しました: {e}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active) # デフォルトの空シートを削除
    
    # --- 依存関係（矢印）とクリティカルパスの計算 ---
    periods_dict = {}
    for task in state.get("tasks", []):
        for p in task.get("periods", []):
            if p.get("pid"):
                periods_dict[p["pid"]] = p

    all_edges = []
    adj = {pid: [] for pid in periods_dict}
    preds = {pid: [] for pid in periods_dict}

    # 全ての矢印の繋がりをリストアップ
    for pid, p in periods_dict.items():
        raw_dep = p.get("dep", "")
        raw_dep_str = str(raw_dep) if raw_dep is not None else ""
        deps = [d.strip() for d in raw_dep_str.split(",") if d.strip()]
        for d in deps:
            if d in periods_dict:
                all_edges.append((d, pid))
                adj[d].append(pid)
                preds[pid].append(d)

    critical_edges = set()
    connected_pids = [pid for pid in periods_dict if adj[pid] or preds[pid]]
    
    if connected_pids:
        # 繋がりがあるタスクの中で、一番遅い終了日を探す
        end_dates = [periods_dict[pid].get("end") for pid in connected_pids if periods_dict[pid].get("end")]
        if end_dates:
            max_end = max(end_dates)
            # その日付で終わるタスク（経路のゴール）を特定
            terminals = [pid for pid in connected_pids if periods_dict[pid].get("end") == max_end]
            
            queue = list(terminals)
            visited = set(terminals)
            # ゴールから遡ってクリティカルパスの経路を特定する
            while queue:
                curr = queue.pop(0)
                for p in preds[curr]:
                    critical_edges.add((p, curr))
                    if p not in visited:
                        visited.add(p)
                        queue.append(p)

    daily_tabs = state.get("dailyNoteTabs", [])
    if not daily_tabs:
        daily_tabs = [{"id": "tab_general", "name": "工程表"}]
    daily_data = state.get("dailyNotesData", {})

    for tab_info in daily_tabs:
        tab_id = tab_info.get("id")
        tab_name = tab_info.get("name", "Sheet")
        safe_title = re.sub(r'[\\*?:/\[\]]', '', tab_name)[:31]
        ws = wb.create_sheet(title=safe_title)
        
        # 【A3印刷設定】
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5
        ws.print_options.horizontalCentered = True
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        project_name = state.get("projectName", "")
        company_name = state.get("companyName", "")
        p_start = state.get("projectStart", "")
        p_end = state.get("projectEnd", "")
        
        ws.cell(row=1, column=1, value="工 程 表").font = Font(size=18, bold=True)
        ws.cell(row=2, column=1, value=f"工事名: {project_name}")
        ws.cell(row=2, column=3, value=f"事業者名: {company_name}")
        ws.cell(row=2, column=5, value=f"全体工期: {p_start} ～ {p_end}")
        ws.cell(row=3, column=5, value="■ 太い赤線の矢印はクリティカルパス").font = Font(color="DC3545", bold=True)

        ws.column_dimensions['A'].width = 5   
        ws.column_dimensions['B'].width = 15  
        ws.column_dimensions['C'].width = 15  
        ws.column_dimensions['D'].width = 20  
        
        start_date_str = state.get("displayStart") or p_start
        end_date_str = state.get("displayEnd") or p_end
        if not start_date_str or not end_date_str:
            continue
            
        d_start = datetime.strptime(start_date_str, "%Y-%m-%d")
        d_end = datetime.strptime(end_date_str, "%Y-%m-%d")
        total_days = (d_end - d_start).days + 1
        date_list = [d_start + timedelta(days=i) for i in range(total_days)]
        
        headers = ["No.", "工種", "種別", "細別・規格"]
        for col, text in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=text)
            cell.alignment = align_center
            ws.cell(row=4, column=col).border = thin_border
            ws.cell(row=5, column=col).border = thin_border
            ws.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
            
        cal_start_col = 5
        notes_col = cal_start_col + total_days 
        
        for i in range(total_days):
            ws.column_dimensions[get_column_letter(cal_start_col + i)].width = 3
            
        ws.column_dimensions[get_column_letter(notes_col)].width = 35
        notes_header = ws.cell(row=4, column=notes_col, value="備考欄")
        notes_header.alignment = align_center
        ws.cell(row=4, column=notes_col).border = thin_border
        ws.cell(row=5, column=notes_col).border = thin_border
        ws.merge_cells(start_row=4, start_column=notes_col, end_row=5, end_column=notes_col)
            
        current_month = -1
        month_start_col = cal_start_col
        holiday_fill = PatternFill(start_color='E9ECEF', end_color='E9ECEF', fill_type='solid')
        national_holiday_fill = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
        holidays_settings = state.get("holidays", {})
        custom_holidays = holidays_settings.get("custom", [])

        def is_holiday(dt):
            date_str = dt.strftime("%Y-%m-%d")
            if holidays_settings.get("nationalHolidays") and date_str in national_holidays: return "national"
            if holidays_settings.get("sundays") and dt.weekday() == 6: return "holiday"
            if holidays_settings.get("saturdays") and dt.weekday() == 5: return "holiday"
            if date_str in custom_holidays: return "holiday"
            return None

        holiday_cols = {}
        for i, dt in enumerate(date_list):
            c_col = cal_start_col + i
            cell_d = ws.cell(row=5, column=c_col, value=dt.day)
            cell_d.alignment = align_center
            cell_d.border = thin_border
            
            if dt.month != current_month:
                if current_month != -1:
                    ws.merge_cells(start_row=4, start_column=month_start_col, end_row=4, end_column=c_col-1)
                month_start_col = c_col
                current_month = dt.month
                cell_m = ws.cell(row=4, column=month_start_col, value=f"{dt.year}年{dt.month}月")
                cell_m.alignment = align_center
                
            ws.cell(row=4, column=c_col).border = thin_border
            ht = is_holiday(dt)
            if ht:
                holiday_cols[c_col] = ht
                if ht == "national":
                    cell_d.fill = national_holiday_fill
                    cell_d.font = Font(color="DC3545")
                else:
                    cell_d.fill = holiday_fill
                    cell_d.font = Font(color="6C757D")

        ws.merge_cells(start_row=4, start_column=month_start_col, end_row=4, end_column=cal_start_col + total_days - 1)

        current_row = 6
        bar_coords = {} # 矢印を引くためにバーの座標を記憶
        
        for task in state.get("tasks", []):
            periods = task.get("periods", [])
            max_drow = 0
            for p in periods:
                if p.get("displayRow", 0) > max_drow:
                    max_drow = p.get("displayRow", 0)
            
            row_span = max_drow + 1
            end_row = current_row + row_span - 1
            
            ws.cell(row=current_row, column=1, value=task.get("no", "")).alignment = align_center
            ws.cell(row=current_row, column=2, value=strip_html_tags(task.get("koshu", ""))).alignment = align_center
            ws.cell(row=current_row, column=3, value=strip_html_tags(task.get("shubetsu", ""))).alignment = align_center
            ws.cell(row=current_row, column=4, value=strip_html_tags(task.get("saibetsu", ""))).alignment = align_center
            
            for c in range(1, 5):
                for r in range(current_row, end_row + 1):
                    ws.cell(row=r, column=c).border = thin_border
                if row_span > 1:
                    ws.merge_cells(start_row=current_row, start_column=c, end_row=end_row, end_column=c)

            for r in range(current_row, end_row + 1):
                for i in range(total_days):
                    c_col = cal_start_col + i
                    cell = ws.cell(row=r, column=c_col)
                    cell.border = thin_border
                    ht = holiday_cols.get(c_col)
                    if ht == "national":
                        cell.fill = PatternFill(start_color='FFF1F2', end_color='FFF1F2', fill_type='solid')
                    elif ht == "holiday":
                        cell.fill = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
                ws.cell(row=r, column=notes_col).border = thin_border
            
            for p in periods:
                start_str = p.get("start")
                end_str = p.get("end")
                if not start_str or not end_str:
                    continue
                    
                drow = p.get("displayRow", 0)
                color_code = p.get("color", "#3b82f6").replace("#", "")
                bar_fill = PatternFill(start_color=color_code, end_color=color_code, fill_type='solid')
                
                try:
                    p_s = datetime.strptime(start_str, "%Y-%m-%d")
                    p_e = datetime.strptime(end_str, "%Y-%m-%d")
                except ValueError:
                    continue
                    
                start_col = None
                end_col = None
                for i, dt in enumerate(date_list):
                    if p_s <= dt <= p_e:
                        c_col = cal_start_col + i
                        if start_col is None: start_col = c_col
                        end_col = c_col
                        cell = ws.cell(row=current_row + drow, column=c_col)
                        cell.fill = bar_fill
                        
                if start_col and end_col:
                    bar_coords[p.get("pid")] = {'r': current_row + drow, 'c_start': start_col, 'c_end': end_col}
                        
            current_row += row_span

        # --- 矢印（依存関係）とクリティカルパスの描画 ---
        def set_b(ws_obj, r, c, side, line_color, line_style):
            cell = ws_obj.cell(row=r, column=c)
            b = cell.border
            kw = {'left': b.left, 'right': b.right, 'top': b.top, 'bottom': b.bottom}
            kw[side] = Side(style=line_style, color=line_color)
            cell.border = Border(**kw)

        for pred, curr in all_edges:
            is_critical = (pred, curr) in critical_edges
            line_color = 'DC3545' if is_critical else '000000' # クリティカルパスは赤、通常は黒
            line_style = 'medium' if is_critical else 'thin'
            
            if pred in bar_coords and curr in bar_coords:
                c_pred = bar_coords[pred]
                c_curr = bar_coords[curr]
                r1 = c_pred['r']
                c1 = c_pred['c_end']
                r2 = c_curr['r']
                c2 = c_curr['c_start']
                
                if c2 <= c1: 
                    continue # 逆行はスキップ
                
                # 先端に▶を配置
                if c2 > c1 + 1 or r1 != r2:
                    head_cell = ws.cell(row=r2, column=c2-1)
                    head_cell.value = "▶"
                    head_cell.font = Font(color=line_color, size=8)
                    head_cell.alignment = Alignment(horizontal='right', vertical='center')

                # 罫線による矢印経路の描画
                if r2 > r1: 
                    for r in range(r1, r2):
                        set_b(ws, r, c1, 'right', line_color, line_style)
                    for c in range(c1+1, c2-1):
                        set_b(ws, r2-1, c, 'bottom', line_color, line_style)
                    set_b(ws, r2-1, c1, 'bottom', line_color, line_style) 
                elif r2 < r1: 
                    for r in range(r2, r1+1):
                        set_b(ws, r, c1, 'right', line_color, line_style)
                    for c in range(c1+1, c2-1):
                        set_b(ws, r2-1, c, 'bottom', line_color, line_style)
                    set_b(ws, r2-1, c1, 'bottom', line_color, line_style)
                else: 
                    for c in range(c1+1, c2-1):
                        cell = ws.cell(row=r1, column=c)
                        cell.value = "ー"
                        cell.font = Font(color=line_color, size=8, bold=True)
                        cell.alignment = Alignment(horizontal='center', vertical='center')

        if current_row > 6:
            notes_text = strip_html_tags(state.get("notes", ""))
            notes_cell = ws.cell(row=6, column=notes_col, value=notes_text)
            notes_cell.alignment = Alignment(vertical='top', wrap_text=True)
            ws.merge_cells(start_row=6, start_column=notes_col, end_row=current_row - 1, end_column=notes_col)

        tab_notes = daily_data.get(tab_id, {})
        ws.cell(row=current_row, column=1, value=tab_name).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for c in range(1, 5):
            ws.cell(row=current_row, column=c).border = thin_border
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
        
        max_lines = 1
        for i, dt in enumerate(date_list):
            c_col = cal_start_col + i
            date_str = dt.strftime("%Y-%m-%d")
            month_str = f"{dt.year}-{dt.month:02d}"
            
            note_html = tab_notes.get(date_str, "")
            if not note_html and dt.day == 1:
                note_html = tab_notes.get(month_str, "")
                
            note_text = strip_html_tags(note_html)
            lines = note_text.count('\n') + 1
            if lines > max_lines:
                max_lines = lines
                
            cell = ws.cell(row=current_row, column=c_col, value=note_text)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='top', horizontal='center', wrap_text=True, textRotation=255)
        
        ws.cell(row=current_row, column=notes_col).border = thin_border
        ws.row_dimensions[current_row].height = max_lines * 15 + 40

    try:
        wb.save(file_path)
    except Exception as e:
        raise RuntimeError(f"Excelの保存に失敗しました: {e}")
