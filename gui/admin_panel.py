import pathlib
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import bcrypt
import pandas as pd
from db.database import connect_db
import logging

# Set up logging
log_dir = pathlib.Path.home() / 'Documents' / 'LetterGeneratorLogs'
log_dir.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    filename=log_dir / 'admin_panel.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class AdminPanel:
    def __init__(self, root):
        self.root = root
        self.root.title("Admin Panel - Officer Management")
        self.root.geometry("900x600")
        self.root.configure(bg="#f4f6f9")
        self.root.resizable(True, True)

        # Title
        title = tk.Label(root, text="⚙️ Admin Dashboard", font=("Segoe UI", 16, "bold"), bg="#f4f6f9", fg="#333")
        title.pack(pady=10)

        # Search and Filter Frame
        search_frame = tk.Frame(root, bg="#f4f6f9")
        search_frame.pack(pady=5, fill="x", padx=20)
        tk.Label(search_frame, text="Search:", font=("Segoe UI", 10, "bold"), bg="#f4f6f9", fg="#333").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, font=("Segoe UI", 10))
        self.search_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self.filter_officers)
        tk.Label(search_frame, text="Filter by:", font=("Segoe UI", 10, "bold"), bg="#f4f6f9", fg="#333").pack(side="left", padx=5)
        self.filter_var = tk.StringVar(value="Username")
        filter_combo = ttk.Combobox(search_frame, textvariable=self.filter_var, values=["Username", "Designation"], state="readonly", width=12)
        filter_combo.pack(side="left", padx=5)
        filter_combo.bind("<<ComboboxSelected>>", self.filter_officers)

        # Officer Treeview
        self.tree = ttk.Treeview(root, columns=("Id", "Username", "Name", "Designation", "Phone", "Email"), show="headings", height=12)
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")
        self.tree.pack(padx=20, pady=10, fill="both", expand=True)

        # Button Frame
        button_frame = tk.Frame(root, bg="#f4f6f9")
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="➕ Add Officer", command=self.add_officer_dialog).pack(side="left", padx=10)
        ttk.Button(button_frame, text="✏️ Edit Officer", command=self.edit_officer_dialog).pack(side="left", padx=10)
        ttk.Button(button_frame, text="🗑️ Delete Selected", command=self.delete_selected).pack(side="left", padx=10)
        ttk.Button(button_frame, text="🔄 Refresh", command=self.load_officers).pack(side="left", padx=10)
        ttk.Button(button_frame, text="📄 Export to CSV", command=self.export_to_csv).pack(side="left", padx=10)
        ttk.Button(button_frame, text="🔑 Change Admin Password", command=self.change_admin_password).pack(side="left", padx=10)
        ttk.Button(button_frame, text="🚪 Logout", command=self.logout).pack(side="left", padx=10)

        # Load officers initially
        self.load_officers()

        # Apply consistent styling
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8)
        style.configure("TCombobox", font=("Segoe UI", 10), padding=5)
        style.configure("TEntry", font=("Segoe UI", 10), padding=5)
        logging.debug("AdminPanel initialized")

    def load_officers(self):
        """Load all officers into the Treeview."""
        for row in self.tree.get_children():
            self.tree.delete(row)
        try:
            conn = connect_db()
            if not conn:
                messagebox.showerror("Error", "Failed to connect to database")
                logging.error("Failed to connect to database")
                return
            cursor = conn.cursor()
            cursor.execute("SELECT Id, Username, OfficerName, Designation, Phone, Email FROM Officers")
            for row in cursor.fetchall():
                self.tree.insert("", "end", values=row)
            logging.debug(f"Loaded {len(self.tree.get_children())} officers")
        except sqlite3.OperationalError as e:
            messagebox.showerror("Error", f"Database operation failed: {e}")
            logging.error(f"Database operation failed: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load officers: {e}")
            logging.error(f"Failed to load officers: {e}")
        finally:
            if conn:
                conn.close()

    def filter_officers(self, event=None):
        """Filter officers based on search term and column."""
        search_term = self.search_var.get().lower()
        filter_by = self.filter_var.get()
        column_index = {"Username": 1, "Designation": 3}[filter_by]
        for row in self.tree.get_children():
            self.tree.delete(row)
        try:
            conn = connect_db()
            if not conn:
                messagebox.showerror("Error", "Failed to connect to database")
                logging.error("Failed to connect to database")
                return
            cursor = conn.cursor()
            query = f"SELECT Id, Username, OfficerName, Designation, Phone, Email FROM Officers WHERE LOWER({filter_by}) LIKE ?"
            cursor.execute(query, (f"%{search_term}%",))
            for row in cursor.fetchall():
                self.tree.insert("", "end", values=row)
            logging.debug(f"Filtered officers by {filter_by}: {search_term}")
        except sqlite3.OperationalError as e:
            messagebox.showerror("Error", f"Database operation failed: {e}")
            logging.error(f"Database operation failed: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to filter officers: {e}")
            logging.error(f"Failed to filter officers: {e}")
        finally:
            if conn:
                conn.close()

    def add_officer_dialog(self):
        """Open dialog to add a new officer."""
        win = tk.Toplevel(self.root)
        win.title("Add Officer")
        win.geometry("400x500")
        win.configure(bg="#f4f6f9")
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)

        # Create a canvas with scrollbar
        canvas = tk.Canvas(win, bg="#f4f6f9")
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f4f6f9")
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        fields = ["Username", "Password", "OfficerName", "Designation", "Phone", "Email"]
        entries = {}

        for field in fields:
            tk.Label(scrollable_frame, text=field, bg="#f4f6f9", anchor="w", font=("Segoe UI", 10, "bold")).pack(fill="x", padx=10, pady=2)
            entry = ttk.Entry(scrollable_frame, show="*" if field == "Password" else "", font=("Segoe UI", 10))
            entry.pack(fill="x", padx=10, pady=2)
            entries[field] = entry

        def save():
            values = {f: e.get().strip() for f, e in entries.items()}
            if not all(values.values()):
                messagebox.showwarning("Validation", "All fields are required.")
                logging.warning("Add officer failed: All fields are required")
                return
            hashed_pw = bcrypt.hashpw(values["Password"].encode("utf-8"), bcrypt.gensalt())
            try:
                conn = connect_db()
                if not conn:
                    messagebox.showerror("Error", "Failed to connect to database")
                    logging.error("Failed to connect to database")
                    return
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO Officers (Username, Password, OfficerName, Designation, Phone, Email) VALUES (?, ?, ?, ?, ?, ?)",
                    (values["Username"], hashed_pw, values["OfficerName"], values["Designation"], values["Phone"], values["Email"])
                )
                conn.commit()
                messagebox.showinfo("Success", "Officer added successfully.")
                logging.debug(f"Added officer: {values['Username']}")
                win.destroy()
                self.load_officers()
            except sqlite3.IntegrityError as e:
                messagebox.showerror("Error", f"Username '{values['Username']}' already exists. Please choose a different username.")
                logging.error(f"Add officer failed: Username '{values['Username']}' already exists - {e}")
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e):
                    messagebox.showerror("Error", "Database is locked. Please try again later.")
                    logging.error("Database is locked during add officer")
                else:
                    messagebox.showerror("Error", f"Database operation failed: {e}")
                    logging.error(f"Database operation failed: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add officer: {e}")
                logging.error(f"Failed to add officer: {e}")
            finally:
                if conn:
                    conn.close()

        ttk.Button(scrollable_frame, text="Save Officer", command=save).pack(pady=10)


    def logout(self):
        if messagebox.askokcancel("Logout", "Are you sure you want to log out?"):
            self.root.destroy()
            from main import main
            main()

    def edit_officer_dialog(self):
        """Open dialog to edit a selected officer."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Select an officer to edit.")
            logging.warning("Edit officer failed: No officer selected")
            return
        officer_id = self.tree.item(selected[0])['values'][0]
        try:
            conn = connect_db()
            if not conn:
                messagebox.showerror("Error", "Failed to connect to database")
                logging.error("Failed to connect to database")
                return
            cursor = conn.cursor()
            cursor.execute("SELECT Username, OfficerName, Designation, Phone, Email FROM Officers WHERE Id = ?", (officer_id,))
            officer_data = cursor.fetchone()
            if not officer_data:
                messagebox.showerror("Error", "Officer not found.")
                logging.error(f"Officer not found for ID: {officer_id}")
                return
        except sqlite3.OperationalError as e:
            messagebox.showerror("Error", f"Database operation failed: {e}")
            logging.error(f"Database operation failed: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load officer data: {e}")
            logging.error(f"Failed to load officer data: {e}")
        finally:
            if conn:
                conn.close()

        win = tk.Toplevel(self.root)
        win.title("Edit Officer")
        win.geometry("400x500")
        win.configure(bg="#f4f6f9")
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)

        # Create a canvas with scrollbar
        canvas = tk.Canvas(win, bg="#f4f6f9")
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f4f6f9")
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        fields = ["Username", "OfficerName", "Designation", "Phone", "Email"]
        entries = {}
        for i, field in enumerate(fields):
            tk.Label(scrollable_frame, text=field, bg="#f4f6f9", anchor="w", font=("Segoe UI", 10, "bold")).pack(fill="x", padx=10, pady=2)
            entry = ttk.Entry(scrollable_frame, font=("Segoe UI", 10))
            entry.insert(0, officer_data[i])
            entry.pack(fill="x", padx=10, pady=2)
            entries[field] = entry

        tk.Label(scrollable_frame, text="New Password (optional)", bg="#f4f6f9", anchor="w", font=("Segoe UI", 10, "bold")).pack(fill="x", padx=10, pady=2)
        password_entry = ttk.Entry(scrollable_frame, show="*", font=("Segoe UI", 10))
        password_entry.pack(fill="x", padx=10, pady=2)

        def save():
            values = {f: e.get().strip() for f, e in entries.items()}
            if not all(values.values()):
                messagebox.showwarning("Validation", "All fields are required.")
                logging.warning("Edit officer failed: All fields are required")
                return
            try:
                conn = connect_db()
                if not conn:
                    messagebox.showerror("Error", "Failed to connect to database")
                    logging.error("Failed to connect to database")
                    return
                cursor = conn.cursor()
                if password_entry.get().strip():
                    hashed_pw = bcrypt.hashpw(password_entry.get().encode("utf-8"), bcrypt.gensalt())
                    cursor.execute(
                        "UPDATE Officers SET Username = ?, Password = ?, OfficerName = ?, Designation = ?, Phone = ?, Email = ? WHERE Id = ?",
                        (values["Username"], hashed_pw, values["OfficerName"], values["Designation"], values["Phone"], values["Email"], officer_id)
                    )
                else:
                    cursor.execute(
                        "UPDATE Officers SET Username = ?, OfficerName = ?, Designation = ?, Phone = ?, Email = ? WHERE Id = ?",
                        (values["Username"], values["OfficerName"], values["Designation"], values["Phone"], values["Email"], officer_id)
                    )
                conn.commit()
                messagebox.showinfo("Success", "Officer updated successfully.")
                logging.debug(f"Updated officer ID: {officer_id}")
                win.destroy()
                self.load_officers()
            except sqlite3.IntegrityError as e:
                messagebox.showerror("Error", f"Username '{values['Username']}' already exists. Please choose a different username.")
                logging.error(f"Edit officer failed: Username '{values['Username']}' already exists - {e}")
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e):
                    messagebox.showerror("Error", "Database is locked. Please try again later.")
                    logging.error("Database is locked during edit officer")
                else:
                    messagebox.showerror("Error", f"Database operation failed: {e}")
                    logging.error(f"Database operation failed: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update officer: {e}")
                logging.error(f"Failed to update officer: {e}")
            finally:
                if conn:
                    conn.close()

        ttk.Button(scrollable_frame, text="Save Changes", command=save).pack(pady=10)

    def delete_selected(self):
        """Delete the selected officer."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Select an officer to delete.")
            logging.warning("Delete officer failed: No officer selected")
            return
        officer_id = self.tree.item(selected[0])['values'][0]
        username = self.tree.item(selected[0])['values'][1]
        if username == 'admin':
            messagebox.showerror("Error", "The admin account cannot be deleted.")
            logging.error("Attempted to delete admin account")
            return
        confirm = messagebox.askyesno("Confirm", "Are you sure you want to delete the selected officer?")
        if not confirm:
            return
        try:
            conn = connect_db()
            if not conn:
                messagebox.showerror("Error", "Failed to connect to database")
                logging.error("Failed to connect to database")
                return
            cursor = conn.cursor()
            cursor.execute("DELETE FROM Officers WHERE Id = ?", (officer_id,))
            conn.commit()
            self.tree.delete(selected[0])
            messagebox.showinfo("Success", "Officer deleted successfully.")
            logging.debug(f"Deleted officer ID: {officer_id}")
        except sqlite3.OperationalError as e:
            if "database is locked" in str(e):
                messagebox.showerror("Error", "Database is locked. Please try again later.")
                logging.error("Database is locked during delete officer")
            else:
                messagebox.showerror("Error", f"Database operation failed: {e}")
                logging.error(f"Database operation failed: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete officer: {e}")
            logging.error(f"Failed to delete officer: {e}")
        finally:
            if conn:
                conn.close()

    def export_to_csv(self):
        """Export officer list to CSV."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            conn = connect_db()
            if not conn:
                messagebox.showerror("Error", "Failed to connect to database")
                logging.error("Failed to connect to database")
                return
            cursor = conn.cursor()
            cursor.execute("SELECT Id, Username, OfficerName, Designation, Phone, Email FROM Officers")
            data = cursor.fetchall()
            df = pd.DataFrame(data, columns=["Id", "Username", "OfficerName", "Designation", "Phone", "Email"])
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Officer list exported to {file_path}")
            logging.debug(f"Exported officer list to {file_path}")
        except sqlite3.OperationalError as e:
            messagebox.showerror("Error", f"Database operation failed: {e}")
            logging.error(f"Database operation failed: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export to CSV: {e}")
            logging.error(f"Failed to export to CSV: {e}")
        finally:
            if conn:
                conn.close()

    def change_admin_password(self):
        """Open dialog to change admin password."""
        win = tk.Toplevel(self.root)
        win.title("Change Admin Password")
        win.geometry("400x250")
        win.configure(bg="#f4f6f9")
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)

        tk.Label(win, text="New Password:", bg="#f4f6f9", font=("Segoe UI", 10, "bold")).pack(pady=5, padx=10)
        new_password = ttk.Entry(win, show="*", font=("Segoe UI", 10))
        new_password.pack(pady=5, padx=10)
        tk.Label(win, text="Confirm Password:", bg="#f4f6f9", font=("Segoe UI", 10, "bold")).pack(pady=5, padx=10)
        confirm_password = ttk.Entry(win, show="*", font=("Segoe UI", 10))
        confirm_password.pack(pady=5, padx=10)

        def save():
            if new_password.get() != confirm_password.get():
                messagebox.showerror("Error", "Passwords do not match.")
                logging.error("Change admin password failed: Passwords do not match")
                return
            if not new_password.get():
                messagebox.showerror("Error", "Password cannot be empty.")
                logging.error("Change admin password failed: Password is empty")
                return
            try:
                hashed_pw = bcrypt.hashpw(new_password.get().encode("utf-8"), bcrypt.gensalt())
                conn = connect_db()
                if not conn:
                    messagebox.showerror("Error", "Failed to connect to database")
                    logging.error("Failed to connect to database")
                    return
                cursor = conn.cursor()
                cursor.execute("UPDATE Officers SET Password = ? WHERE Username = 'admin'", (hashed_pw,))
                conn.commit()
                messagebox.showinfo("Success", "Admin password updated successfully.")
                logging.debug("Admin password updated successfully")
                win.destroy()
            except sqlite3.OperationalError as e:
                if "database is locked" in str(e):
                    messagebox.showerror("Error", "Database is locked. Please try again later.")
                    logging.error("Database is locked during change admin password")
                else:
                    messagebox.showerror("Error", f"Database operation failed: {e}")
                    logging.error(f"Database operation failed: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update password: {e}")
                logging.error(f"Failed to update password: {e}")
            finally:
                if conn:
                    conn.close()

        ttk.Button(win, text="Save Password", command=save).pack(pady=10)