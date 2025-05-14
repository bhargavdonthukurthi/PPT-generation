def slides_to_remove(ppt_file: str, teams: list):
    if set(teams) == {"ADB", "RDM", "CDL", "CDH", "SCUP NA", "Data BAU"}:
        return []

    prs = Presentation(ppt_file)
    remove = []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                if len(table.rows[0].cells) > 2:
                    if table.rows[0].cells[0].text.strip() == "Application" and table.rows[0].cells[1].text.strip() == "Planned user stories" and table.rows[0].cells[2].text.strip() == "Mid Sprint user stories":
                        rows_to_delete_indices = []
                        slide_removed = False
                        for r_idx, row in enumerate(table.rows):
                            app_name = row.cells[0].text.strip()
                            if app_name in ('EMIS UI', 'EMIS Backend'):
                                if "ADB" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name in ('CBIP', 'D&B'):
                                if "ADB" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name == 'RDM':
                                if 'RDM' not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name in ('KPI', 'EDW', 'Trade Credit'):
                                if 'Data BAU' not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name == 'CDL':
                                if 'CDL' not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'CDH':
                                if 'CDH' not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'SCUP NA':
                                if 'SCUP NA' not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break

                        # Delete rows using the working method
                        for row_index in sorted(rows_to_delete_indices, reverse=True):
                            if len(table.rows) > row_index:  # Ensure index is valid
                                table._tbl.remove(table._tbl.tr_lst[row_index])

                        if slide_removed:
                            break

                elif len(table.rows[0].cells) == 2:
                    if table.rows[0].cells[0].text.strip() == "Application" and table.rows[0].cells[1].text.strip().lower() == "planned user stories":
                        rows_to_delete_indices = []
                        slide_removed = False
                        for r_idx, row in enumerate(table.rows):
                            app_name = row.cells[0].text.strip()
                            if app_name in ('EMIS UI', 'EMIS Backend', 'CBIP', 'D&B'):
                                if "ADB" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name == 'RDM':
                                if "RDM" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name in ('KPI', 'EDW', 'Trade Credit'):
                                if "Data BAU" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name == 'CDL':
                                if 'CDL' not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'CDH':
                                if 'CDH' not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'SCUP NA':
                                if 'SCUP NA' not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break

                        # Delete rows using the working method
                        for row_index in sorted(rows_to_delete_indices, reverse=True):
                            if len(table.rows) > row_index:  # Ensure index is valid
                                table._tbl.remove(table._tbl.tr_lst[row_index])

                        if slide_removed:
                            break
                    elif table.rows[0].cells[0].text.strip() == "Application" and table.rows[0].cells[1].text.strip().lower() == "implemented user stories":
                        rows_to_delete_indices = []
                        slide_removed = False
                        for r_idx, row in enumerate(table.rows):
                            app_name = row.cells[0].text.strip()
                            if app_name in ('EMIS UI', 'EMIS Backend', 'CBIP', 'D&B'):
                                if "ADB" not in teams and "RDM" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                                elif "ADB" not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'RDM':
                                if "ADB" not in teams and "RDM" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                                elif "RDM" not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name in ('KPI', 'EDW', 'Trade Credit'):
                                if "Data BAU" not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break
                            elif app_name == 'CDL':
                                if 'CDL' not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'CDH':
                                if 'CDH' not in teams:
                                    rows_to_delete_indices.append(r_idx)
                            elif app_name == 'SCUP NA':
                                if 'SCUP NA' not in teams:
                                    remove.append(i)
                                    slide_removed = True
                                    break

                        # Delete rows using the working method
                        for row_index in sorted(rows_to_delete_indices, reverse=True):
                            if len(table.rows) > row_index:  # Ensure index is valid
                                table._tbl.remove(table._tbl.tr_lst[row_index])

                        if slide_removed:
                            break
    print(remove)
    return sorted(list(set(remove)))
