import os

import utilities.support_functions as sup_functions
import json
from datetime import datetime


def get_data_from_haf_xlsx(data):
    print("========> get_data_from_haf_xlsx")
    """open excel app"""
    excel_wrapper = sup_functions.ExcelWrapper()
    with excel_wrapper as excel:
        if excel.Workbooks.Count == 0:
            quit_excel_after_process = True
        else:
            quit_excel_after_process = False

        """loop each ha file and fetch required information and tag nc slide type"""
        for ha_file in data['ha_files']:
            wb = sup_functions.open_workbook(excel, ha_file['ha_file'])
            data['part_num'] = sup_functions.get_part_number_by_jap_sting(wb)
            nc_sheet = sup_functions.get_worksheet(wb, "NonConformance")

            if nc_sheet:
                ha_file['eval_nc_remarks'] = ''
                ha_file['nc_data'] = get_nc_data(nc_sheet)
                ha_file['nc_images'] = get_nc_image(nc_sheet)
                con_sheet = sup_functions.get_worksheet(wb, "Conclusion")

                inspect_id = ha_file['nc_data']['t1_insp_id']
                inspect_sheet = sup_functions.get_worksheet_by_partial_text(wb, inspect_id + '_NC')
                dwb_sheet = sup_functions.get_worksheet(wb, 'DWG')
                mcb_sheet = sup_functions.get_worksheet(wb, 'OMax_MCB')
                fbo_sheet = sup_functions.get_worksheet(wb, 'LCF_FBO')
                if con_sheet:
                    ha_file['eval_data'], ha_file['eval_nc_data'], ha_file['eval_nc_remarks'] = get_eval_data(con_sheet)
                if inspect_sheet or dwb_sheet:
                    ha_file['inspect_id_data'], ha_file['eval_remarks'] = get_inspect_sheet_data(inspect_sheet, dwb_sheet)
                    ha_file['template_type'] = '3'
                elif mcb_sheet or fbo_sheet:
                    ha_file['mcb_data'], ha_file['fbo_data'] = get_mcb_fbo_sheet_data(mcb_sheet, fbo_sheet)
                    ha_file['template_type'] = '2'
                else:
                    ha_file['template_type'] = '1'

                if ha_file['eval_nc_remarks'] == '':
                    ha_file['eval_nc_remarks'] = get_nc_eval_remarks(nc_sheet)
                    
            sup_functions.close_workbook(wb)
            
        """cleanup"""
        if quit_excel_after_process:
            sup_functions.quit_excel(excel)

    return data


def get_mcb_fbo_sheet_data(mcb_sheet, fbo_sheet):
    """get_mcb_fbo_sheet_data"""
    fbo_data = {}
    mcb_data = {}

    if mcb_sheet:
        green_cell = rgb_to_excel_color(146, 208, 80)
        nc_head_row = 0
        nc_start_row = 0
        nc_end_row = 0
        nc_eval_remarks = ''
        for row in range(2, mcb_sheet.UsedRange.Rows.Count + 1):
            cell_text = mcb_sheet.Cells(row, 1).Text
            desc_text = mcb_sheet.Cells(row, 2).Text
            cell_color = mcb_sheet.Cells(row, 2).Interior.Color
            if nc_head_row == 0 and 'NC' in cell_text:
                nc_head_row = row
            elif nc_head_row > 0 and nc_start_row == 0 and 'Item' in cell_text:
                nc_start_row = row + 1
            elif nc_start_row > 0 and nc_end_row == 0 and '1.5 RSS of Tolerances'.lower() in desc_text.lower():
                nc_end_row = row
            elif nc_end_row > 0 and cell_color == green_cell and len(desc_text) > 0:
                if nc_eval_remarks == '':
                    nc_eval_remarks = desc_text
                else:
                    nc_eval_remarks = '\n' + desc_text

        mcb_data['nc_eval_remarks'] = nc_eval_remarks
        mcb_data['image1'] = ''
        if nc_head_row > 0:
            sup_functions.clear_clip_board()
            copy_range = mcb_sheet.Range(mcb_sheet.Cells(nc_start_row, 1), mcb_sheet.Cells(nc_end_row, 'N'))
            time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
            img_name = f'mcb_nc_table_{time_stamp}.png'
            is_copy_success, file_path = sup_functions.save_range_as_image(copy_range, img_name)
            mcb_data['image1'] = file_path if is_copy_success else ''

    if fbo_sheet:
        fbo_data['image1'] = ''
        fbo_data['image2'] = ''
        fbo_data['image3'] = ''
        for shape in fbo_sheet.Shapes:
            if shape.Name == 'fbo_1':
                sup_functions.clear_clip_board()
                time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
                img_name = f'fbo_image1_{time_stamp}.png'
                is_copy_success, file_path = sup_functions.save_shape_as_image(shape, img_name)
                fbo_data['image1'] = file_path if is_copy_success else ''
            elif shape.Name == 'fbo_2':
                sup_functions.clear_clip_board()
                time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
                img_name = f'fbo_image2_{time_stamp}.png'
                is_copy_success, file_path = sup_functions.save_shape_as_image(shape, img_name)
                fbo_data['image2'] = file_path if is_copy_success else ''
            elif shape.Name == 'fbo_3':
                sup_functions.clear_clip_board()
                time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
                img_name = f'fbo_image3_{time_stamp}.png'
                is_copy_success, file_path = sup_functions.save_shape_as_image(shape, img_name)
                fbo_data['image3'] = file_path if is_copy_success else ''

        fbo_data['eval_1_text1'] = fbo_sheet.Cells(2, 2).Text
        fbo_data['eval_1_text2'] = fbo_sheet.Cells(5, 2).Text
        eval_1_text3 = ''
        for row in range(24, 33):
            row_text = ''
            for col in range(2,5):
                cell_text = fbo_sheet.Cells(row, col).Text
                if len(cell_text) > 0:
                    row_text = cell_text if row_text == '' else row_text + ' ' + cell_text
            if row_text != '':
                eval_1_text3 = row_text if eval_1_text3 == '' else eval_1_text3 + '\n' + row_text
        fbo_data['eval_1_text3'] = eval_1_text3

        fbo_data['eval_2_text1'] = fbo_sheet.Cells(2, 2).Text
        fbo_data['eval_2_text2'] = fbo_sheet.Cells(5, 2).Text
        eval_2_text3 = ''
        for row in range(24, 33):
            row_text = ''
            for col in range(15,18):
                cell_text = fbo_sheet.Cells(row, col).Text
                if len(cell_text) > 0:
                    row_text = cell_text if row_text == '' else row_text + ' ' + cell_text
            if row_text != '':
                eval_2_text3 = row_text if eval_2_text3 == '' else eval_2_text3 + '\n' + row_text
        fbo_data['eval_2_text3'] = eval_2_text3

        fbo_data['eval_2_remarks'] = fbo_sheet.Cells(37, 2).Text

    return mcb_data, fbo_data

def rgb_to_excel_color(red, green, blue):
    """rgb_to_excel_color"""
    return red + (green * 256) + (blue * 256 * 256)

def get_inspect_sheet_data(nc_sheet, dwb_sheet):
    """get_inspect_sheet_data"""
    data = {}
    text_list = ['image1', 'image2', 'image3', 'image4']
    eval_remarks = ''
    if nc_sheet:
        sup_functions.clear_clip_board()
        copy_range = nc_sheet.Range('B2:C16')
        time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
        img_name = f'inspect_table_{text_list[0]}_{time_stamp}.png'
        is_copy_success, file_path = sup_functions.save_range_as_image(copy_range, img_name)
        data[text_list[0]] = file_path if is_copy_success else ''

        data[text_list[1]] = ''
        data[text_list[2]] = ''
        for shape in nc_sheet.Shapes:
            if shape.Name == 'nc_1':
                sup_functions.clear_clip_board()
                time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
                img_name = f'inspect_table_{text_list[1]}_{time_stamp}.png'
                is_copy_success, file_path = sup_functions.save_shape_as_image(shape, img_name)
                data[text_list[1]] = file_path if is_copy_success else ''
            elif shape.Name == 'nc_2':
                sup_functions.clear_clip_board()
                time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
                img_name = f'inspect_table_{text_list[2]}_{time_stamp}.png'
                is_copy_success, file_path = sup_functions.save_shape_as_image(shape, img_name)
                data[text_list[2]] = file_path if is_copy_success else ''
        eval_remarks = nc_sheet.Range('B34').Text

    if dwb_sheet:
        sup_functions.clear_clip_board()
        copy_range = dwb_sheet.UsedRange
        time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
        img_name = f'inspect_table_{text_list[3]}_{time_stamp}.png'
        is_copy_success, file_path = sup_functions.save_range_as_image(copy_range, img_name)
        data[text_list[3]] = file_path if is_copy_success else ''

    return data, eval_remarks


def get_nc_data(sheet):
    """get_nc_data"""
    column_mapping = {'B': 't1_ln', 'C': 't1_zone', 'D': 't1_dwg', 'E': 't1_desc', 'F': 't1_insp_id', 'G': 't1_nc',
                      'I': 't1_ihi_no', 'J': 't1_wc_nc', 'K': 't1_exceed', 'L': 't1_meas', 'M': 't1_eval',
                      'N': 't1_remarks'}
    result_dict = {}
    for col_letter, key in column_mapping.items():
        if col_letter == 'G':
            value = f"{sheet.Range('G6').Text} {sheet.Range('H6').Text} " \
                    f"({sheet.Range('G7').Text} {sheet.Range('H7').Text})"
        else:
            value = sheet.Range(col_letter + '6').Text
        result_dict[key] = value

    return result_dict

def get_nc_eval_remarks(nc_sheet):
    """get_nc_eval_remarks"""
    eval_nc_remarks = ''
    for row in range(9, nc_sheet.UsedRange.Rows.Count + 1):
        for col in range(1, nc_sheet.UsedRange.Columns.Count + 1):
            cell_value = str(nc_sheet.Cells(row, col).Value).lower()
            if 'this nonconformance' in cell_value:
                eval_nc_remarks = cell_value
                return eval_nc_remarks
    return eval_nc_remarks

def get_nc_image(sheet):
    """get_nc_image"""
    images_dict = {}
    for index, shape in enumerate(sheet.Shapes):
        shape_name = shape.Name
        time_stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
        img_name = 'nc_image_{}_{}.png'.format(shape_name, time_stamp)
        img_saved, file_path = sup_functions.save_shape_as_image(shape, img_name)
        if img_saved:
            images_dict[str(index+1)] = file_path
    return images_dict

def get_eval_data(sheet):
    """get_eval_data"""
    eval_header_row = 0
    eval_data_row = 0
    eval_nc_header_row = 0
    eval_nc_data_row = 0
    nc_remarks_data_row = 0
    eval_result_dict = {}
    eval_nc_result_dict = {}
    eval_nc_remarks = ''
    for row in range(2, sheet.UsedRange.Rows.Count + 1):
        cell_text = sheet.Cells(row, 2).Text
        remarks_text = sheet.Cells(row, 4).Text
        if eval_header_row == 0 and 'Evaluation  Method and Criteria' in cell_text:
            eval_header_row = row
        elif eval_header_row >0 and eval_data_row == 0 and 'Line no' in cell_text:
            eval_data_row = row + 1

        elif eval_nc_header_row == 0 and 'Nonconformance evaluation:' in cell_text:
            eval_nc_header_row = row
        elif eval_nc_header_row > 0 and eval_nc_data_row == 0 and 'Line no' in cell_text:
            eval_nc_data_row = row +1
        elif eval_nc_data_row > 0 and 'nonconformance' in remarks_text.lower():
            nc_remarks_data_row = row

    if eval_data_row > 0:
        cell = sheet.Cells(eval_data_row, 2)
        if cell.MergeCells:
            merged_range = cell.MergeArea
            eval_min_row = merged_range.Cells(1, 1).Row
            eval_max_row = merged_range.Cells(merged_range.Rows.Count, 1).Row
        else:
            eval_min_row = eval_data_row
            eval_max_row = eval_data_row

        column_mapping = {'B': 't2_ln', 'C': 't2_dwg', 'D': 't2_eval', 'E': 't2_criteria'}

        for col_letter, key in column_mapping.items():
            cell_text = ''
            for row in range(eval_min_row, eval_max_row+1):
                value = sheet.Range(col_letter + str(row)).Text
                if value != '':
                    cell_text = value if cell_text == '' else cell_text + '\n' + value

            eval_result_dict[key] = cell_text

    if eval_nc_data_row > 0:
        cell = sheet.Cells(eval_nc_data_row, 2)
        if cell.MergeCells:
            merged_range = cell.MergeArea
            eval_nc_min_row = merged_range.Cells(1, 1).Row
            eval_nc_max_row = merged_range.Cells(merged_range.Rows.Count, 1).Row
        else:
            eval_nc_min_row = eval_nc_data_row
            eval_nc_max_row = eval_nc_data_row

        column_mapping = {'B': 't3_ln', 'C': 't3_dwg', 'D': 't3_eval', 'E': 't3_dispo'}

        for col_letter, key in column_mapping.items():
            cell_text = ''
            for row in range(eval_nc_min_row, eval_nc_max_row + 1):
                value = sheet.Range(col_letter + str(row)).Text
                if value != '':
                    cell_text = value if cell_text == '' else cell_text + '\n' + value
            eval_nc_result_dict[key] = cell_text
    if nc_remarks_data_row > 0:
        eval_nc_remarks = sheet.Range('D' + str(nc_remarks_data_row)).Text
    return eval_result_dict, eval_nc_result_dict, eval_nc_remarks


def main():
    input_path = r"C:\Users\HP\PycharmProjects\GE_Passport\input"
    ppt_template = os.path.join(input_path,"template.pptx")
    file_path1 = os.path.join(input_path,"HA16K230237_1.xls")
    file_path2 = os.path.join(input_path,"HA16K230237_2.xlsx")
    file_path3 = os.path.join(input_path,"HA16K230237_3.xls")
    file_path4 = os.path.join(input_path,"HA16K230237_4.xlsx")
    output_file = os.path.join(input_path,"output.pptx")
    serial_num = 'SERIAL_NUMBER'
    cr_num = 'CR_NUMBER'
    data = {
        'ppt_template': ppt_template,
        'output_file': output_file,
        'ser_num': serial_num,
        'cr_num': cr_num,
        'ha_files': [{'ha_file': file_path1, 'sec_num': '01'},
                     {'ha_file': file_path2, 'sec_num': '02'},
                     {'ha_file': file_path3, 'sec_num': '03'},
                     {'ha_file': file_path4, 'sec_num': '04'}
                     ]
    }

    result = get_data_from_haf_xlsx(data)
    output_file = os.path.join(input_path, "ref_data.json")
    with open(output_file, 'w') as f:
        json.dump(result, f, indent=4)


if __name__ == '__main__':

    main()

