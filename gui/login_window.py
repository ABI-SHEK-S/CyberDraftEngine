from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import bcrypt
from db.database import connect_db
from gui.admin_panel import AdminPanel
import sqlite3
import os
import sys


# Writable database path (in user's Documents folder)
WRITABLE_DB_DIR = os.path.join(os.path.expanduser("~"), "Documents", "LetterGeneratorDB")
os.makedirs(WRITABLE_DB_DIR, exist_ok=True)
WRITABLE_DB_PATH = os.path.join(WRITABLE_DB_DIR, "letter_requests.db")

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class LoginWindow:
    def __init__(self, root, app_callback):
        self.root = root
        self.app_callback = app_callback
        self.root.title("Cyber Crime Wing - Login")
        self.root.geometry("400x600")
        self.root.minsize(400, 550)
        self.root.configure(bg="#f0f4f8")
        self.root.resizable(True, True)
        self.after_id = None

        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

        container = tk.Frame(self.root, bg="white", bd=5, relief="flat")
        container.place(relx=0.5, rely=0.5, anchor="center", width=350)
        container.configure(highlightbackground="#cccccc", highlightthickness=1)

        # Logo canvas
        self.logo_canvas = tk.Canvas(container, bg="white", width=300, height=100, highlightthickness=0)
        self.logo_canvas.pack(pady=10)

        try:
            # Instead of: self.logo = tk.PhotoImage(file="assets/police.png")
            self.logo = tk.PhotoImage(file=resource_path("assets/police.png"))

            self.logo = self.logo.subsample(self.logo.width() // 100, self.logo.height() // 100)
            self.logo_canvas.create_image(150, 50, image=self.logo)
        except tk.TclError as e:
            tk.Label(self.logo_canvas, text="Logo Not Found", font=("Arial", 12), bg="white", fg="gray").pack()
            print(f"Logo load error: {str(e)}")

        header_frame = tk.Frame(container, bg="#003087")
        header_frame.pack(fill="x")
        tk.Label(
            header_frame,
            text="Cyber Crime Wing",
            font=("Arial", 18, "bold"),
            fg="white",
            bg="#003087",
            pady=10,
            wraplength=300,
            justify="center"
        ).pack()

        main_frame = tk.Frame(container, bg="white")
        main_frame.pack(expand=True, fill="both", padx=20, pady=10)

        tk.Label(
            main_frame,
            text="Username",
            font=("Arial", 12, "bold"),
            bg="white",
            fg="#333333"
        ).pack(pady=(10, 5))
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(
            main_frame,
            textvariable=self.username_var,
            font=("Arial", 12),
            width=25
        )
        self.username_entry.pack(pady=5)
        self.username_entry.focus()

        tk.Label(
            main_frame,
            text="Password",
            font=("Arial", 12, "bold"),
            bg="white",
            fg="#333333"
        ).pack(pady=(10, 5))
        self.password_entry = ttk.Entry(
            main_frame,
            font=("Arial", 12),
            width=25,
            show="*"
        )
        self.password_entry.pack(pady=5)
        self.password_entry.bind("<Return>", lambda event: self.login())

        self.show_password = tk.BooleanVar()
        ttk.Checkbutton(
            main_frame,
            text="Show Password",
            variable=self.show_password,
            command=self.toggle_password,
            style="TCheckbutton"
        ).pack(pady=(5, 10))

        self.login_button = ttk.Button(
            main_frame,
            text="Login",
            command=self.login,
            style="TButton"
        )
        self.login_button.pack(pady=20)

        forgot_btn = tk.Button(
            main_frame,
            text="Reset Password",
            command=self.open_reset_password_window,
            fg="blue",
            bg="white",
            relief="flat",
            cursor="hand2"
        )
        forgot_btn.pack(pady=(5, 0))

        self.error_label = tk.Label(
            main_frame,
            text="",
            font=("Arial", 10),
            fg="red",
            bg="white",
            wraplength=300
        )
        self.error_label.pack(pady=5)

        style = ttk.Style()
        style.configure("TButton", font=("Arial", 14, "bold"), padding=12)
        style.map("TButton",
                  background=[('active', '#0052cc'), ('!active', '#003087')],
                  foreground=[('active', 'white'), ('!active', 'red')],
                  relief=[('active', 'raised'), ('!active', 'flat')])

        self.officer = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        if self.after_id is not None:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        self.root.destroy()

    def toggle_password(self):
        self.password_entry.config(show="" if self.show_password.get() else "*")

    def open_reset_password_window(self):
        """Simple password reset window - just username and new password"""
        reset_popup = tk.Toplevel(self.root)
        reset_popup.title("Reset Password")
        reset_popup.geometry("400x300")
        reset_popup.configure(bg="white")
        reset_popup.resizable(False, False)
        reset_popup.grab_set()
        reset_popup.transient(self.root)

        # Center the popup
        reset_popup.update_idletasks()
        width = reset_popup.winfo_width()
        height = reset_popup.winfo_height()
        x = (reset_popup.winfo_screenwidth() // 2) - (width // 2)
        y = (reset_popup.winfo_screenheight() // 2) - (height // 2)
        reset_popup.geometry(f'{width}x{height}+{x}+{y}')

        tk.Label(reset_popup, text="Reset Password", font=("Arial", 16, "bold"), bg="white").pack(pady=20)

        # Username field
        tk.Label(reset_popup, text="Enter Username:", font=("Arial", 12), bg="white").pack(pady=(10, 5))
        username_entry = tk.Entry(reset_popup, font=("Arial", 12), width=25)
        username_entry.pack(pady=5)
        username_entry.focus()

        # New password field
        tk.Label(reset_popup, text="Enter New Password:", font=("Arial", 12), bg="white").pack(pady=(15, 5))
        new_password_entry = tk.Entry(reset_popup, show="*", font=("Arial", 12), width=25)
        new_password_entry.pack(pady=5)

        # Confirm password field
        tk.Label(reset_popup, text="Confirm New Password:", font=("Arial", 12), bg="white").pack(pady=(15, 5))
        confirm_password_entry = tk.Entry(reset_popup, show="*", font=("Arial", 12), width=25)
        confirm_password_entry.pack(pady=5)

        def reset_password():
            username = username_entry.get().strip()
            new_password = new_password_entry.get().strip()
            confirm_password = confirm_password_entry.get().strip()

            # Validation
            if not username:
                messagebox.showerror("Error", "Username is required.", parent=reset_popup)
                return
            
            if not new_password:
                messagebox.showerror("Error", "New password is required.", parent=reset_popup)
                return
            
            if len(new_password) < 6:
                messagebox.showerror("Error", "Password must be at least 6 characters long.", parent=reset_popup)
                return
            
            if new_password != confirm_password:
                messagebox.showerror("Error", "Passwords do not match.", parent=reset_popup)
                return

            # Database operations
            conn = connect_db()
            if not conn:
                messagebox.showerror("Error", "Database connection failed.", parent=reset_popup)
                return

            try:
                cursor = conn.cursor()
                
                # Check if username exists
                cursor.execute("SELECT Id, OfficerName FROM Officers WHERE Username = ?", (username,))
                officer = cursor.fetchone()
                
                if not officer:
                    messagebox.showerror("Error", "Username does not exist.", parent=reset_popup)
                    return

                # Hash the new password
                hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
                
                # Update the password
                cursor.execute("UPDATE Officers SET Password = ? WHERE Username = ?", (hashed_password, username))
                conn.commit()

                messagebox.showinfo(
                    "Success", 
                    f"Password reset successfully for {officer[1]}!\nYou can now login with your new password.", 
                    parent=reset_popup
                )
                reset_popup.destroy()

            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"Failed to reset password: {str(e)}", parent=reset_popup)
            except Exception as e:
                messagebox.showerror("Error", f"Unexpected error: {str(e)}", parent=reset_popup)
            finally:
                conn.close()

        # Reset button
        reset_btn = tk.Button(
            reset_popup,
            text="Reset Password",
            command=reset_password,
            bg="#003087",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=5
        )
        reset_btn.pack(pady=20)

        # Cancel button
        cancel_btn = tk.Button(
            reset_popup,
            text="Cancel",
            command=reset_popup.destroy,
            bg="#666666",
            fg="white",
            font=("Arial", 12),
            padx=20,
            pady=5
        )
        cancel_btn.pack(pady=5)

        # Bind Enter key to reset password
        reset_popup.bind('<Return>', lambda event: reset_password())

    def login(self):
        print("Login button clicked!")
        if self.after_id is not None:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        
        username = self.username_var.get().strip()
        password = self.password_entry.get().strip()
        
        if not username:
            self.error_label.config(text="Please enter a username.")
            return
        if not password:
            self.error_label.config(text="Please enter your password.")
            return

        conn = connect_db()
        if not conn:
            self.error_label.config(text="Failed to connect to the database.")
            return
        
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM Officers WHERE Username = ?", (username,))
            self.officer = cursor.fetchone()
            
            if self.officer:
                stored_hash = self.officer[2]  # Password is at index 2
                if bcrypt.checkpw(password.encode('utf-8'), stored_hash):
                    self.root.destroy()
                    new_root = tk.Tk()
                    username = self.officer[1]  # Username is at index 1
                    if username.lower() == "admin":
                        AdminPanel(new_root)
                    else:
                        self.app_callback(self.officer, new_root)
                else:
                    self.error_label.config(text="Invalid username or password.")
            else:
                self.error_label.config(text="Invalid username or password.")
                
        except sqlite3.Error as e:
            self.error_label.config(text=f"Database error: {str(e)}")
        except Exception as e:
            self.error_label.config(text=f"Login error: {str(e)}")
        finally:
            if conn:
                conn.close()
