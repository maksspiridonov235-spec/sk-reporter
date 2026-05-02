import re

with open('agent/inject_agent.py', 'r') as f:
    content = f.read()

old_code = '''            if target_cell and (part1_lines or part2_lines):
                _write_parts_to_cell(target_cell, part1_lines, part2_lines)
                print(f"[INJECT_AGENT] Wrote corrected parts to 'Описание действий' cell")'''

new_code = '''            if target_cell and (part1_lines or part2_lines):
                # Find table and row index
                target_table = None
                target_row_idx = None
                for ti, table in enumerate(doc.tables):
                    for ri, row in enumerate(table.rows):
                        for ci, cell in enumerate(row.cells):
                            if cell == target_cell:
                                target_table = table
                                target_row_idx = ri
                                break
                        if target_row_idx is not None:
                            break
                    if target_row_idx is not None:
                        break
                
                if target_table is not None:
                    # Add two new rows
                    col_count = len(target_table.rows[0].cells) if target_table.rows else 0
                    
                    # Add row for Part 1
                    new_row_1 = target_table.add_row()
                    if part1_lines and col_count > 0:
                        new_row_1.cells[0].text = "\\n".join(part1_lines)
                    
                    # Add row for Part 2  
                    new_row_2 = target_table.add_row()
                    if part2_lines and col_count > 0:
                        new_row_2.cells[0].text = "\\n".join(part2_lines)
                    
                    # Move new rows to correct position (after target_row_idx)
                    tbl = target_table._tbl
                    tr_elements = list(tbl.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'))
                    
                    if len(tr_elements) >= 2 and target_row_idx < len(tr_elements) - 2:
                        # Get the rows we just added (last two)
                        tr_1 = tr_elements[-2]
                        tr_2 = tr_elements[-1]
                        ref_tr = tr_elements[target_row_idx]
                        
                        # Remove from end
                        tbl.remove(tr_2)
                        tbl.remove(tr_1)
                        
                        # Insert after reference row
                        ref_tr.addnext(tr_2)
                        ref_tr.addnext(tr_1)
                    
                    print(f"[INJECT_AGENT] Inserted 2 new rows with Part 1 and Part 2")'''

content = content.replace(old_code, new_code)

with open('agent/inject_agent.py', 'w') as f:
    f.write(content)

print("Done")
