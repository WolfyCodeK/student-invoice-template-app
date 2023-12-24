from invoice_app import InvoiceApp
from ctypes import windll

if __name__ == '__main__':
    windll.shcore.SetProcessDpiAwareness(1)
    
    app = InvoiceApp()  
    app.run()