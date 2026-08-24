import pandas as pd
import io
import pypdf
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font

def get_active_classes_from_timetable(timetable_file, exam_date, exam_session):
    active_classes = []
    
    if not timetable_file:
        return []
        
    filename = timetable_file.name.lower()
    
    # Standardise user input exam_date
    try:
        exam_date_std = pd.to_datetime(exam_date, dayfirst=True).strftime('%d-%m-%Y')
    except:
        exam_date_std = exam_date
    
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        df = pd.read_excel(timetable_file, header=None)
        header_idx = -1
        for idx, row in df.iterrows():
            row_str = ' '.join([str(x).upper() for x in row.values])
            if 'DATE' in row_str and 'SESSION' in row_str and 'DEPARTMENT' in row_str:
                header_idx = idx
                break
                
        if header_idx != -1:
            df.columns = df.iloc[header_idx]
            df = df[header_idx+1:]
            
            current_date = ''
            current_sess = ''
            
            for idx, row in df.iterrows():
                d = row.get('Date', '')
                if pd.notnull(d) and str(d).strip() != '':
                    try:
                        current_date = pd.to_datetime(d, dayfirst=True).strftime('%d-%m-%Y')
                    except:
                        current_date = str(d).strip()
                    
                s = str(row.get('Session', '')).strip()
                if s != 'nan' and s != '':
                    current_sess = s
                    
                if current_date == exam_date_std and current_sess == exam_session:
                    yr = str(row.get('Year', '')).strip().upper()
                    dpt = str(row.get('Department', '')).strip().upper()
                    y_str = ''
                    if yr == 'I': y_str = '1-YEAR'
                    elif yr == 'II': y_str = '2-YEAR'
                    elif yr == 'III' or yr == 'OG': y_str = '3-YEAR'
                    
                    dept_list = []
                    if 'ALL' in dpt:
                        dept_list = ['CS', 'AI&DS', 'B.COM', 'B A TAMIL', 'CHE', 'BCA', 'BBA', 'MIB']
                    else:
                        for dept_name in ['CS', 'AI&DS', 'B.COM', 'B A TAMIL', 'CHE', 'BCA', 'BBA', 'MIB']:
                            if dept_name.replace('.', '') in dpt.replace('.', ''):
                                dept_list.append(dept_name)
                    
                    for dept_name in dept_list:
                        if y_str:
                            active_classes.append(f'{dept_name} - {y_str}')
                            
        return sorted(list(set(active_classes)))
        
    # Fallback to PDF logic
    reader = pypdf.PdfReader(timetable_file)
    search_str = (exam_date + exam_session).upper().replace(" ", "")
    
    active_classes = set()
    current_degree = ''
    current_sem = -1
    
    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue
        for line in text.split('\n'):
            if line.startswith('DEGREE NAME:'):
                current_degree = line.split('DEGREE NAME:')[1].strip().upper()
            elif 'SEMESTER :' in line:
                sem_str = line.split('SEMESTER')[0].strip()
                if sem_str.isdigit():
                    current_sem = int(sem_str)
            
            # Check if line has the exam
            if search_str in line.replace(' ', '').upper():
                year = ''
                if current_sem in [1, 2]: year = '1-YEAR'
                elif current_sem in [3, 4]: year = '2-YEAR'
                elif current_sem in [5, 6]: year = '3-YEAR'
                
                # Map degree to sheet
                sheets = []
                if 'COMPUTER SCIENCE' in current_degree or 'CS' in current_degree: sheets.append('CS')
                if 'ARTIFICIAL INTELLIGENCE' in current_degree or 'DATA SCIENCE' in current_degree or 'AI&DS' in current_degree: sheets.append('AI&DS')
                if 'COMMERCE' in current_degree or 'B.COM' in current_degree: sheets.append('B.COM')
                if 'TAMIL' in current_degree: sheets.append('B A TAMIL')
                if 'CHEMISTRY' in current_degree or 'CHE' in current_degree: sheets.append('CHE')
                if 'COMPUTER APPLICATIONS' in current_degree or 'BCA' in current_degree: sheets.append('BCA')
                if 'FOUNDATION' in current_degree:
                    sheets = ['CS', 'AI&DS', 'B.COM', 'B A TAMIL', 'CHE', 'BCA', 'BBA', 'MIB']
                
                for s in sheets:
                    if year:
                        active_classes.add(f'{s} - {year}')
                        
    return list(active_classes)

def get_student_classes(student_file):
    xls_students = pd.ExcelFile(student_file)
    classes = set()
    for sheet in xls_students.sheet_names:
        df_sheet = pd.read_excel(xls_students, sheet_name=sheet)
        if len(df_sheet) > 1:
            header_idx = -1
            for i in range(min(5, len(df_sheet))):
                row_vals = [str(x).strip() for x in df_sheet.iloc[i].tolist()]
                if 'REG.NO' in row_vals or 'REG NO' in row_vals or 'REGISTER NUMBER' in row_vals:
                    header_idx = i
                    break
            if header_idx != -1:
                headers = df_sheet.iloc[header_idx].tolist()
                df_data = df_sheet.iloc[header_idx+1:].copy()
                df_data.columns = headers
                
                # Check for year column
                year_col = None
                for c in headers:
                    if isinstance(c, str) and 'YEAR' in c.upper():
                        year_col = c
                        break
                
                if year_col:
                    unique_years = df_data[year_col].dropna().unique()
                    for y in unique_years:
                        classes.add(f"{sheet} - {str(y).strip()}")
                else:
                    classes.add(f"{sheet} - UNKNOWN YEAR")
    return sorted(list(classes))

def process_allocation(hall_file, student_file, timetable_file, exam_date, exam_session, selected_classes, exam_title):
    # 1. Read Hall Details
    if hall_file.name.lower().endswith('.pdf'):
        import pypdf
        reader = pypdf.PdfReader(hall_file)
        halls = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    parts = [p.strip() for p in line.split() if p.strip()]
                    if len(parts) >= 2:
                        hall_no = parts[0]
                        try:
                            total_seat = int(parts[1])
                            rows = 5
                            cols = 6
                            if len(parts) >= 4:
                                try:
                                    rows = int(parts[2])
                                    cols = int(parts[3])
                                except:
                                    pass
                            if 0 < total_seat <= 200:
                                halls.append({'HALL NO': hall_no, 'TOTAL SEAT': total_seat, 'ROWS': rows, 'COLUMNS': cols})
                        except:
                            pass
        if not halls:
            raise ValueError("Could not find valid hall data in the PDF. Ensure lines look like: '101 30 5 6'")
        df_halls = pd.DataFrame(halls)
    else:
        df_halls = pd.read_excel(hall_file)

    # 2. Read Student Details (read all sheets)
    xls_students = pd.ExcelFile(student_file)
    all_students = []
    
    for sheet in xls_students.sheet_names:
        df_sheet = pd.read_excel(xls_students, sheet_name=sheet)
        if len(df_sheet) > 1:
            header_idx = -1
            for i in range(min(5, len(df_sheet))):
                row_vals = [str(x).strip() for x in df_sheet.iloc[i].tolist()]
                if 'REG.NO' in row_vals or 'REG NO' in row_vals or 'REGISTER NUMBER' in row_vals:
                    header_idx = i
                    break
            
            if header_idx != -1:
                headers = df_sheet.iloc[header_idx].tolist()
                df_data = df_sheet.iloc[header_idx+1:].copy()
                df_data.columns = headers
                
                # Find the exact column name that matches REG.NO
                reg_col = [c for c in headers if isinstance(c, str) and 'REG' in c.upper()][0]
                
                # Check for year column
                year_col = None
                for c in headers:
                    if isinstance(c, str) and 'YEAR' in c.upper():
                        year_col = c
                        break
                
                # Clean up empty rows
                df_data = df_data.dropna(subset=[reg_col])
                
                for _, row in df_data.iterrows():
                    student_year = str(row[year_col]).strip() if year_col and not pd.isna(row[year_col]) else "UNKNOWN YEAR"
                    class_key = f"{sheet} - {student_year}"
                    
                    if class_key in selected_classes:
                        all_students.append({
                            'REG_NO': row[reg_col],
                            'NAME': row.iloc[1] if len(row) > 1 else '',
                            'DEPT': class_key
                        })
    
    # 3. Dummy PDF reading for timetable (real extraction logic depends on PDF format)
    # reader = pypdf.PdfReader(timetable_file)
    # text = reader.pages[0].extract_text()
    
    # 4. Allocation Logic
    # Group students by Department
    students_by_dept = {}
    for s in all_students:
        dept = s['DEPT']
        if dept not in students_by_dept:
            students_by_dept[dept] = []
        students_by_dept[dept].append(s)
        
    depts = list(students_by_dept.keys())
    
    # Create the output excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # We will write a single output sheet simulating the printed layout
        out_rows = []
        out_rows.append([exam_title] + [""] * 11)
        
        for _, hall_row in df_halls.iterrows():
            hall_no = hall_row.get('HALL NO', 'UNKNOWN')
            
            # Read ROWS and COLUMNS dynamically, default to 5x6
            grid_cols = int(hall_row.get('COLUMNS', 6))
            grid_rows = int(hall_row.get('ROWS', 5))
            
            # If they provided TOTAL SEAT but no ROWS/COLUMNS, we calculate it:
            if 'TOTAL SEAT' in hall_row and 'ROWS' not in hall_row:
                total_seats = int(hall_row.get('TOTAL SEAT', 30))
                grid_rows = total_seats // grid_cols if total_seats % grid_cols == 0 else (total_seats // grid_cols) + 1
            else:
                total_seats = int(hall_row.get('TOTAL SEAT', grid_rows * grid_cols))
            
            max_cols = grid_cols * 2
            empty_middle = max_cols - 3
            
            # We have visual columns to fill. Each visual column takes `grid_rows` students.
            hall_visual_columns = []
            
            last_dept = None
            hall_assigned = 0
            
            for v_col in range(grid_cols):
                if hall_assigned >= total_seats:
                    break
                    
                col_students = []
                
                # Keep filling this column until it has `grid_rows` students OR hall is full
                while len(col_students) < grid_rows and hall_assigned < total_seats:
                    chosen_dept = None
                    
                    # Try to find a different department with enough students
                    for d in depts:
                        if d != last_dept and len(students_by_dept[d]) > 0:
                            chosen_dept = d
                            break
                    
                    # Fallback: if we couldn't find a different one, just pick any with students
                    if chosen_dept is None:
                        for d in depts:
                            if len(students_by_dept[d]) > 0:
                                chosen_dept = d
                                break
                                
                    if chosen_dept is None:
                        break # No students left at all
                        
                    # Calculate how many more students we need for this column, respecting hall total cap
                    max_allowed_for_hall = total_seats - hall_assigned
                    needed = min(grid_rows - len(col_students), max_allowed_for_hall)
                    take_count = min(needed, len(students_by_dept[chosen_dept]))
                    
                    # Add them to the column
                    col_students.extend(students_by_dept[chosen_dept][:take_count])
                    students_by_dept[chosen_dept] = students_by_dept[chosen_dept][take_count:]
                    hall_assigned += take_count
                    
                    # Update last_dept so we don't pick the same one consecutively (unless fallback)
                    last_dept = chosen_dept
                
                if col_students:
                    hall_visual_columns.append(col_students)
                
            if not hall_visual_columns:
                continue # Skip empty halls
                
            # Format the block for this hall
            out_rows.append([f"Date & Session: {exam_date} & {exam_session}"] + [""] * empty_middle + [f"HALL NO : {hall_no}", ""])

            
            # Calculate actual assigned total and department names
            total_assigned = sum(len(c) for c in hall_visual_columns)
            
            # Calculate actual assigned total and department counts
            total_assigned = sum(len(c) for c in hall_visual_columns)
            
            # --- Extract subcodes from PDF or Excel ---
            dept_subcodes = {}
            if timetable_file:
                try:
                    # Standardise user input exam_date
                    try:
                        exam_date_std = pd.to_datetime(exam_date, dayfirst=True).strftime('%d-%m-%Y')
                    except:
                        exam_date_std = exam_date
                        
                    if hasattr(timetable_file, 'seek'):
                        timetable_file.seek(0)
                        
                    filename = timetable_file.name.lower() if hasattr(timetable_file, 'name') else ""
                    if filename.endswith('.xlsx') or filename.endswith('.xls'):
                        df_tt = pd.read_excel(timetable_file, header=None)
                        header_idx = -1
                        for idx, row_tt in df_tt.iterrows():
                            row_str = ' '.join([str(x).upper() for x in row_tt.values])
                            if 'DATE' in row_str and 'SESSION' in row_str and 'DEPARTMENT' in row_str:
                                header_idx = idx
                                break
                                
                        if header_idx != -1:
                            df_tt.columns = df_tt.iloc[header_idx]
                            df_tt = df_tt[header_idx+1:]
                            
                            c_date = ''
                            c_sess = ''
                            
                            for idx, row_tt in df_tt.iterrows():
                                d = row_tt.get('Date', '')
                                if pd.notnull(d) and str(d).strip() != '':
                                    try:
                                        c_date = pd.to_datetime(d, dayfirst=True).strftime('%d-%m-%Y')
                                    except:
                                        c_date = str(d).strip()
                                    
                                s = str(row_tt.get('Session', '')).strip()
                                if s != 'nan' and s != '': c_sess = s
                                    
                                if c_date == exam_date_std and c_sess == exam_session:
                                    yr = str(row_tt.get('Year', '')).strip().upper()
                                    dpt = str(row_tt.get('Department', '')).strip().upper()
                                    sc = str(row_tt.get('Paper Code', '')).strip()
                                    y_str = ''
                                    if yr == 'I': y_str = '1-YEAR'
                                    elif yr == 'II': y_str = '2-YEAR'
                                    elif yr == 'III' or yr == 'OG': y_str = '3-YEAR'
                                    
                                    dept_list = []
                                    if 'ALL' in dpt:
                                        dept_list = ['CS', 'AI&DS', 'B.COM', 'B A TAMIL', 'CHE', 'BCA', 'BBA', 'MIB']
                                    else:
                                        for dept_name in ['CS', 'AI&DS', 'B.COM', 'B A TAMIL', 'CHE', 'BCA', 'BBA', 'MIB']:
                                            if dept_name.replace('.', '') in dpt.replace('.', ''):
                                                dept_list.append(dept_name)
                                    
                                    for dept_name in dept_list:
                                        if y_str:
                                            key = f'{dept_name} - {y_str}'
                                            if key not in dept_subcodes or 'TA' in sc:
                                                dept_subcodes[key] = sc
                    else:
                        reader = pypdf.PdfReader(timetable_file)
                        search_str = (exam_date + exam_session).upper().replace(' ', '')
                        current_degree = ''
                        current_sem = -1
                        for page in reader.pages:
                            text = page.extract_text()
                            if not text: continue
                            for line in text.split('\n'):
                                if line.startswith('DEGREE NAME:'):
                                    current_degree = line.split('DEGREE NAME:')[1].strip().upper()
                                elif 'SEMESTER :' in line:
                                    sem_str = line.split('SEMESTER')[0].strip()
                                    if sem_str.isdigit():
                                        current_sem = int(sem_str)
                                
                                if search_str in line.replace(' ', '').upper():
                                    sub_code = line.split(' ')[0].strip()
                                    year = ''
                                    if current_sem in [1, 2]: year = '1-YEAR'
                                    elif current_sem in [3, 4]: year = '2-YEAR'
                                    elif current_sem in [5, 6]: year = '3-YEAR'
                                    
                                    sheets = []
                                    if 'COMPUTER SCIENCE' in current_degree or 'CS' in current_degree: sheets.append('CS')
                                    if 'ARTIFICIAL INTELLIGENCE' in current_degree or 'DATA SCIENCE' in current_degree or 'AI&DS' in current_degree: sheets.append('AI&DS')
                                    if 'COMMERCE' in current_degree or 'B.COM' in current_degree: sheets.append('B.COM')
                                    if 'TAMIL' in current_degree: sheets.append('B A TAMIL')
                                    if 'CHEMISTRY' in current_degree or 'CHE' in current_degree: sheets.append('CHE')
                                    if 'COMPUTER APPLICATIONS' in current_degree or 'BCA' in current_degree: sheets.append('BCA')
                                    if 'FOUNDATION' in current_degree:
                                        sheets = ['CS', 'AI&DS', 'B.COM', 'B A TAMIL', 'CHE', 'BCA', 'BBA', 'MIB', 'Sheet3']
                                        
                                    for s_name in sheets:
                                        key = f'{s_name} - {year}'
                                        if key not in dept_subcodes or 'TA' in sub_code: 
                                            dept_subcodes[key] = sub_code
                except:
                    pass
            # ---------------------------------
            
            dept_counts = {}
            for col_students in hall_visual_columns:
                for student in col_students:
                    d = student['DEPT']
                    dept_counts[d] = dept_counts.get(d, 0) + 1
                    
            dept_parts = []
            for d, count in dept_counts.items():
                parts = d.split(" - ")
                dept_name = parts[0]
                roman = ""
                if len(parts) > 1:
                    if parts[1] == "1-YEAR": roman = "I"
                    elif parts[1] == "2-YEAR": roman = "II"
                    elif parts[1] == "3-YEAR": roman = "III"
                
                prefix = f"{roman}-{dept_name}" if roman else dept_name
                subcode = dept_subcodes.get(d, "")
                if subcode:
                    dept_parts.append(f"{prefix} - {count} ({subcode})")
                else:
                    dept_parts.append(f"{prefix} - {count}")
                    
            dept_str = " & ".join(dept_parts)
            
            out_rows.append([f"DEPT/SUB CODE : {dept_str}"] + [""] * empty_middle + [f"TOTAL : {total_assigned}", ""])
            
            header = []
            for _ in range(grid_cols):
                header.extend(["REGISTER NUMBER", "SEAT NO"])
            out_rows.append(header)
            
            # Build the matrix
            seat_matrix = [['' for _ in range(max_cols)] for _ in range(grid_rows)]
            
            # Simple logical column mapping: Fill left-to-right based on actual visual columns available
            # Or you can do a zigzag if you want, but simple 0, 1, 2... mapping is robust for dynamic columns.
            for v_col_idx, col_students in enumerate(hall_visual_columns):
                for row_idx, student in enumerate(col_students):
                    if row_idx < grid_rows:
                        # Simple sequence numbering down the column
                        seat_num = v_col_idx * grid_rows + row_idx + 1
                        
                        seat_matrix[row_idx][v_col_idx*2] = student['REG_NO']
                        seat_matrix[row_idx][v_col_idx*2 + 1] = seat_num
            
            out_rows.extend(seat_matrix)
            out_rows.append([]) # Blank row
            out_rows.append([]) # Blank row

        if not out_rows:
            raise ValueError("No students could be mapped. Please check if the uploaded Student Details file has 'REG.NO' columns.")

        # At this point, find the global max_cols across all halls to ensure safe DataFrame dimensions
        global_max_cols = max([len(r) for r in out_rows] + [0])
        # Pad all rows to global_max_cols
        padded_out_rows = [r + [''] * (global_max_cols - len(r)) for r in out_rows]
        
        out_df = pd.DataFrame(padded_out_rows)
        out_df.to_excel(writer, index=False, header=False, sheet_name='Seating')
        
        # Format the Excel sheet
        ws = writer.sheets['Seating']
        
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        # Set column widths
        for col_idx in range(1, global_max_cols + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            if col_idx % 2 != 0:
                ws.column_dimensions[col_letter].width = 18 # REGISTER NUMBER columns
            else:
                ws.column_dimensions[col_letter].width = 10 # SEAT NO columns
                
        # Apply borders and alignments
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=global_max_cols):
            # Check if the row is entirely empty (separator row)
            is_empty_row = all(c.value is None or str(c.value).strip() == "" for c in row)
            
            if not is_empty_row:
                first_val = str(row[0].value) if row[0].value else ""
                is_header_row = (row[0].row == 1) or ("Date & Session" in first_val) or ("DEPT/SUB CODE" in first_val) or ("REGISTER NUMBER" in first_val)
                
                # Determine how many columns this specific hall uses (look for empty cells at the end)
                # But to keep it simple, we draw borders across the entire populated length of this row.
                row_length = 0
                for idx, c in enumerate(row):
                    if c.value is not None and str(c.value).strip() != "":
                        row_length = idx + 1
                
                # If it's a header row, we might want to span it up to the width it should be.
                # Actually, relying on global_max_cols is fine for borders as long as it looks uniform.
                
                for cell in row:
                    # Draw borders only if it's within the populated area or it's a global header
                    if cell.column <= global_max_cols:
                        cell.border = thin_border
                    
                    val_str = str(cell.value) if cell.value else ""
                    
                    # Align center by default
                    if cell.row == 1 or "HALL NO :" in val_str or "TOTAL :" in val_str:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif "Date & Session:" in val_str or "DEPT/SUB CODE :" in val_str:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    # Make headers bold
                    if is_header_row:
                        if cell.row == 1:
                            cell.font = Font(bold=True, size=14)
                        else:
                            cell.font = Font(bold=True)

        # Merge cells AFTER setting borders
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=global_max_cols):
            first_cell_val = str(row[0].value) if row[0].value else ""
            
            if row[0].row == 1:
                # Top title row
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=global_max_cols)
            elif "Date & Session:" in first_cell_val:
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=global_max_cols-2)
                ws.merge_cells(start_row=row[0].row, start_column=global_max_cols-1, end_row=row[0].row, end_column=global_max_cols)
            elif "DEPT/SUB CODE :" in first_cell_val:
                ws.merge_cells(start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=global_max_cols-2)
                ws.merge_cells(start_row=row[0].row, start_column=global_max_cols-1, end_row=row[0].row, end_column=global_max_cols)
        
    return output.getvalue()
