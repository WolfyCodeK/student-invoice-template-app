from datetime import datetime
import os
import shutil
import sys
from invoice_app import InvoiceApp
from ctypes import windll

if __name__ == '__main__':
    windll.shcore.SetProcessDpiAwareness(1)

    # Delete previous version if app has just been updated
    if len(sys.argv) > 2:
        version = str(sys.argv[2])
        
        with open('error_log.txt', 'a') as f:
            f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - Version found -> v{version}\n')
            f.close()
        
        if (version != "") and (version is not None):
            # Get the absolute path of this file
            current_file_path = os.path.abspath(__file__)

            parent_path =  os.path.dirname(os.path.dirname(current_file_path))
            old_app_path = os.path.join(parent_path, f'StudentInvoice-{version}')
            
            try:
                with open('error_log.txt', 'a') as f:
                    f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - Deleting directory -> {old_app_path}\n')
                    f.close()
                
                with open('error_log.txt', 'a') as f:
                    f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - Path: {old_app_path} -> Permissions -> {os.stat(old_app_path)}\n')
                    f.close()
                
                shutil.rmtree(old_app_path)
                
            except OSError as e:
                with open('error_log.txt', 'a') as f:
                    f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [ERROR] - OSERROR DELETEING PREVIOUS VERSION!!!\n')
                    f.close()
    else:
        with open('error_log.txt', 'a') as f:
            f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - App manually launched by user.\n')
            f.close()

    app = InvoiceApp()  
    app.run()