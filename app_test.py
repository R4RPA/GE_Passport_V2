import os
from utilities.get_data_from_haf_xlsx import get_data_from_haf_xlsx
from utilities.create_output_pptx import create_ouput_pptx


def main():
    """Initiate data params to process"""
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

    """get data from ha excel files"""
    result = get_data_from_haf_xlsx(data)

    """create ppt report"""
    create_ouput_pptx(result)


if __name__ == '__main__':

    main()