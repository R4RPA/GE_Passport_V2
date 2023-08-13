#pyinstaller --onefile --windowed --paths Lib\site-packages -i "icon.ico" app.py

import os
import sys
import json

from datetime import datetime
from PyQt5 import QtWidgets

from utilities.get_data_from_haf_xlsx import get_data_from_haf_xlsx
from utilities.create_output_pptx import create_ouput_pptx
from utilities.app_gui import Ui_MainWindow
import utilities.support_functions as sup_functions


class UiWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(UiWindow, self).__init__()
        """Initiate GUI Window"""
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        """On Click and On Change Actions"""
        self.ui.browse_ha_files_folder.clicked.connect(self.browse_ha_files_folder)
        self.ui.browse_ppt_template_file.clicked.connect(self.browse_ppt_template_file)
        self.ui.browse_output_ppt_folder.clicked.connect(self.browse_output_ppt_folder)
        
        self.ui.reset_selections.clicked.connect(self.reset_selections)
        self.ui.generate_reports.clicked.connect(self.generate_reports)
        self.ui.exit_app.clicked.connect(self.close)

    def browse_file(self):
        """Browse for File"""
        return QtWidgets.QFileDialog.getOpenFileName(self, 'Select File', '', 'All Files (*);;Text Files (*.txt)')[0]
    
    def browse_folder(self):
        self.ui.statusbar.showMessage('')
        """Browse for Folder """
        return QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Folder', '')

    def browse_ha_files_folder(self):
        """Browse for folder to save INPUT files folder"""
        folder = self.browse_folder()
        self.ui.ha_files_folder.setText(folder)
        self.ui.ha_files_folder.setToolTip(folder)

    def browse_ppt_template_file(self):
        """Browse for folder to save output file"""
        file_path = self.browse_file()
        if file_path.endswith('.pptx'):
            self.ui.ppt_template_file.setText(file_path)
            self.ui.ppt_template_file.setToolTip(file_path)
        else:
            self.ui.ppt_template_file.setText("path....")
            self.ui.ppt_template_file.setToolTip("path....")
            self.ui.statusbar.showMessage('Please select PPTX Template File')
            print('   Please select PPTX Template File')
            
    def browse_output_ppt_folder(self):
        """Browse for folder to save output file folder"""
        folder = self.browse_folder()
        self.ui.output_ppt_folder.setText(folder)
        self.ui.output_ppt_folder.setToolTip(folder)

    def reset_selections(self):
        """Reset all  form fields to default status"""
        self.ui.ha_files_folder.setText("path....")
        self.ui.ppt_template_file.setText("path....")
        self.ui.output_ppt_folder.setText("path....")
        
        self.ui.ha_files_folder.setToolTip("path....")
        self.ui.ppt_template_file.setToolTip("path....")
        self.ui.output_ppt_folder.setToolTip("path....")
        
        self.ui.cr_number.setText("")
        self.ui.sr_number.setText("")
        self.ui.statusbar.showMessage('')
    
    def generate_reports(self):
        print('=== GENERATE PASSPORT REPORT - START')
        inputs = [self.ui.ha_files_folder.text(), 
                   self.ui.ppt_template_file.text(),
                   self.ui.output_ppt_folder.text(),
                   self.ui.cr_number.text(),
                   self.ui.sr_number.text()]
        
        for i, inputval in enumerate(inputs):
            inputs[i] = '' if inputval == "path...." else inputval
        
        
        if '' in inputs:
            self.ui.statusbar.showMessage('Missing Input params: fill the form and try again')
            print('   Missing Input params: fill the form and try again')
        else:
            self.ui.statusbar.showMessage('Validating input files')
            print('   Validating input files')
            msg, input_dict = self.validate_input_folders(inputs)
            if len(msg) == 0:
                self.ui.statusbar.showMessage('Starting Analysis....')
                print('   Starting NC Analysis with given inputs....')
                print(json.dumps(input_dict, indent=2))
                            
                """get data from ha excel files"""
                result = get_data_from_haf_xlsx(input_dict)
            
                """create ppt report"""
                create_ouput_pptx(result)
                
                self.ui.statusbar.showMessage('DONE!!   - Check Output Folder for Reports')
                print('=== GENERATE REPORT - END')
            else:
                self.ui.statusbar.showMessage(msg)
    
    
    
    def validate_input_folders(self, inputs):
        [ha_folder, template_file, output_folder, cr_num, sr_num] = inputs
        ha_files = sup_functions.check_for_files(ha_folder, "*.xls", "*.xlsx")
        print('ha_files', ha_files)
        msg = ''
        if ha_files:
            ha_file_list = []
            for index, file_path in enumerate(ha_files):
                ha_file_list.append({'ha_file': file_path, 'sec_num': str(index+1).zfill(2)})
            
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
            output_file = os.path.join(output_folder, f'output_{timestamp}.pptx')
            input_dict = {
                'ppt_template': template_file,
                'output_file': output_file,
                'ser_num': sr_num,
                'cr_num': cr_num,
                'ha_files': ha_file_list
            }
        else:
            msg = 'No input files found in given folders...'
            print('   No input files found in given folders...')
            input_dict = {}
        return msg, input_dict
    
def create_app():
    """Initiate PyQT Application"""
    app = QtWidgets.QApplication(sys.argv)
    win = UiWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    create_app()
