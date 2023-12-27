from datetime import datetime
import os
import shutil
import sys
from invoice_app import InvoiceApp
from ctypes import windll

if __name__ == '__main__':
    windll.shcore.SetProcessDpiAwareness(1)

    

    # Delete previous version if app has just been updated
    if len(sys.argv) > 1:
        version = str(sys.argv[1])
        
        with open('error_log.txt', 'a') as f:
            f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - Version found -> v{version}\n')
            f.close()
        
        if version is not None:
            # Get the absolute path of this file
            current_file_path = os.path.abspath(__file__)

            parent_path = os.path.dirname(current_file_path)
            
            try:
                with open('error_log.txt', 'a') as f:
                    f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - Deleting version -> v{version} inside directory -> {parent_path}\n')
                    f.close()
                # shutil.rmtree(directory_to_delete)
            except OSError as e:
                with open('error_log.txt', 'a') as f:
                    f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [ERROR] - FAILED TO DELETE PREVIOUS VERSION!!!\n')
                    f.close()
    else:
        with open('error_log.txt', 'a') as f:
            f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [INFO] - No args found\n')
            f.close()

    app = InvoiceApp()  
    app.run()