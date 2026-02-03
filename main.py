import tkinter as tk
from gui.login_window import LoginWindow
from gui.main_app import LetterGeneratorApp
from db.database import create_database,create_default_admin

def main():
    create_database()
    create_default_admin()
    
    root = tk.Tk()
    app = LoginWindow(root, lambda officer, new_root: LetterGeneratorApp(new_root, officer))
    root.mainloop()

if __name__ == "__main__":
    main()