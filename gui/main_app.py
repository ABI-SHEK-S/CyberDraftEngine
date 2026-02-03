import tkinter as tk
from tkinter import ttk, messagebox, Text
import sqlite3
import re
from tkinter import filedialog
from db.database import connect_db
from .bank_letters import BankLetters
from .inter_letters import InterLetters
from .tsp_letters import TSPLetters
import os
from pathlib import Path
import json
import logging
import sys

CONFIG_FILE = 'config.json'  # Config file to store user settings persistently


    
class LetterGeneratorApp:
    class ToolTip:
        def __init__(self, widget, text, app):
            self.widget = widget
            self.text = text
            self.app = app
            self.tip_window = None
            self.widget.bind("<Enter>", self.show_tip)
            self.widget.bind("<Leave>", self.hide_tip)

        def show_tip(self, event=None):
            if self.tip_window or not self.text:
                return
            x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
            self.tip_window = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                tw, text=self.text, justify=tk.LEFT,
                background="#ffffe0", foreground=self.app.text_color,
                relief=tk.RAISED, borderwidth=2, font=("Segoe UI", 8, "normal"),
                padx=5, pady=5
            )
            label.pack(ipadx=1)

        def hide_tip(self, event=None):
            tw = self.tip_window
            self.tip_window = None
            if tw:
                tw.destroy()

    def __init__(self, root, officer):
        self.root = root
        self.root.title("Cyber Crime Wing - Letter Generator")
        self.root.geometry("1000x800")
        self.root.minsize(800, 600)
        self.root.resizable(True, True)

        # Initialize logging
        log_dir = os.path.join(Path.home(), 'Documents', 'LetterGeneratorLogs')
        os.makedirs(log_dir, exist_ok=True)
        logging.basicConfig(
            filename=os.path.join(log_dir, 'app.log'),
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        logging.debug("LetterGeneratorApp initialized")

        # Initialize default theme colors
        self.default_colors = {
            'bg_color': '#f0f4f8', 'header_bg': '#003087', 'accent_color': '#007bff',
            'text_color': '#333333', 'error_color': '#dc3545', 'success_color': '#28a745',
            'warning_color': '#ffc107'
        }
        self.bg_color = self.default_colors['bg_color']
        self.header_bg = self.default_colors['header_bg']
        self.accent_color = self.default_colors['accent_color']
        self.text_color = self.default_colors['text_color']
        self.error_color = self.default_colors['error_color']
        self.success_color = self.default_colors['success_color']
        self.warning_color = self.default_colors['warning_color']

        # Set default font
        default_font = ("Segoe UI", 10)
        self.root.option_add("*Font", default_font)

        # Initialize case details
        self.crime_number = None
        self.ncrp_id = None
        self.last_case_details = {'CrimeNumber': None, 'NCRP_ID': None}
        self.date_from = None
        self.date_to = None
        self.profile_window = None

        # Initialize template directory
        self.config_dir = os.path.join(Path.home(), 'Documents', 'LetterGenerator')
        self.config_file = os.path.join(self.config_dir, 'config.json')
        self.template_dir = None
        

        self.officer = {
            'Id': officer[0], 'Username': officer[1], 'OfficerName': officer[3],
            'Designation': officer[4], 'Phone': officer[5], 'Email': officer[6]
        }

        self.config = self.load_config()
        self.template_dir = self.config.get('template_dir', "")

        if not self.template_dir:
            print("Template directory not set yet.")
        else:
            print(f"Loaded template directory: {self.template_dir}")

        # Center the window
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

        # Inside your main app class initialization (__init__ or setup_ui method)

        

        # Create header frame
        header_frame = tk.Frame(self.root, bg=self.header_bg)
        header_frame.pack(side=tk.TOP, fill=tk.X)

        # Title Label centered at the top
        title_frame = tk.Frame(header_frame, bg=self.header_bg)
        title_frame.pack(side=tk.TOP, fill=tk.X)
        title_label = tk.Label(
            title_frame,
            text="Cyber Crime Wing - Letter Generator",
            font=("Segoe UI", 16, "bold"),
            bg=self.header_bg,
            fg="white"
        )
        title_label.pack(pady=10)  # pack center by default

        # Right side frame for buttons below title
        right_header = tk.Frame(header_frame, bg=self.header_bg)
        right_header.pack(side=tk.TOP, anchor="e", pady=0, padx=10)


        # Buttons container frame using grid to arrange buttons in 2x2 grid
        buttons_frame = tk.Frame(right_header, bg=self.header_bg)
        buttons_frame.pack(anchor="ne")

        # Define a fixed button width (number of characters)
        button_width = 16

        # Row 1
        self.template_dir_button = ttk.Button(
            buttons_frame, text="Template Directory",
            command=self.change_template_dir,
            style="TButton", width=button_width
        )
        self.template_dir_button.grid(row=0, column=0, padx=5, pady=5)
        self.ToolTip(self.template_dir_button, "Change directory for templates", self)

        self.profile_btn = ttk.Button(
            buttons_frame, text="Profile",
            command=self.toggle_profile,
            style="TButton", width=button_width
        )
        self.profile_btn.grid(row=0, column=1, padx=5, pady=5)
        self.ToolTip(self.profile_btn, "Edit your profile details", self)

        # Row 2
        self.help_button = ttk.Button(
            buttons_frame, text="Help",
            command=self.show_help,
            style="TButton", width=button_width
        )
        self.help_button.grid(row=1, column=0, padx=5, pady=5)
        self.ToolTip(self.help_button, "Instructions and help", self)

        self.logout_button = ttk.Button(
            buttons_frame, text="Logout",
            command=self.logout,
            style="TButton", width=button_width
        )
        self.logout_button.grid(row=1, column=1, padx=5, pady=5)
        self.ToolTip(self.logout_button, "Logout", self)

        # Optionally make the columns expand equally for better alignment
        buttons_frame.grid_columnconfigure(0, weight=1)
        buttons_frame.grid_columnconfigure(1, weight=1)



        # Keyboard shortcuts
        self.root.bind('<Control-c>', lambda e: self.set_case_details())
        self.root.bind('<Control-t>', lambda e: self.next_tab())
        self.root.bind('<Control-T>', lambda e: self.prev_tab())

        # Main Frame with Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.bank_frame = tk.Frame(self.notebook, bg=self.bg_color)
        self.inter_frame = tk.Frame(self.notebook, bg=self.bg_color)
        self.tsp_frame = tk.Frame(self.notebook, bg=self.bg_color)

        self.notebook.add(self.bank_frame, text="Bank Letters")
        self.notebook.add(self.inter_frame, text="Intermediaries")
        self.notebook.add(self.tsp_frame, text="TSP Letters")

        # Initialize tab modules
        
        self.inter_letters = InterLetters(self.inter_frame, self)
        self.tsp_letters = TSPLetters(self.tsp_frame, self)

        # Welcome Message and Case Details in Bank tab
        welcome_frame = tk.Frame(self.bank_frame, bg=self.bg_color)
        welcome_frame.pack(pady=20, fill="x")

        tk.Label(
            welcome_frame, text=f"Welcome, {self.officer['OfficerName']} ({self.officer['Designation']})",
            font=("Segoe UI", 14, "bold"), bg=self.bg_color, fg=self.header_bg
        ).pack(pady=10, anchor="center")

        self.case_details_button = ttk.Button(
            welcome_frame, text="Enter Case Details (Ctrl+C)", command=self.set_case_details, style="TButton"
        )
        self.case_details_button.pack(pady=5)
        self.ToolTip(self.case_details_button, "Enter crime number and NCRP ID for the case (Ctrl+C)", self)

        self.case_details_label = tk.Label(
            welcome_frame, text="Case Details: Not Set", font=("Segoe UI", 10, "italic"),
            bg=self.bg_color, fg=self.text_color
        )
        self.case_details_label.pack(pady=5)

        self.bank_letters = BankLetters(self.bank_frame, self)

        # Apply default theme
        self.set_theme()
        self.update_button_states()

        # Footer
        footer_frame = tk.Frame(self.root, bg=self.header_bg)
        footer_frame.pack(side=tk.BOTTOM, fill="x")
        tk.Label(
            footer_frame, text="Letter Generator v1.0 | © 2025 Cyber Crime Wing",
            font=("Segoe UI", 8), fg="white", bg=self.header_bg
        ).pack(pady=5)

    def load_template_dir(self):
        """Load template directory from config file or set to None if invalid."""
        os.makedirs(self.config_dir, exist_ok=True)
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
                template_dir = config.get('template_dir')
                # Verify the template directory and subdirectories exist
                subdirs = ['banks', 'inter', 'tsp']
                if (template_dir and os.path.isdir(template_dir) and
                    all(os.path.isdir(os.path.join(template_dir, subdir)) for subdir in subdirs) and
                    os.path.exists(os.path.join(template_dir, 'banks', 'bank.docx')) and
                    any(f.endswith('.docx') for f in os.listdir(os.path.join(template_dir, 'inter'))) and
                    any(f.endswith('.docx') for f in os.listdir(os.path.join(template_dir, 'tsp')))):
                    self.template_dir = template_dir
                    logging.debug(f"Loaded template directory: {template_dir}")
                    return
                else:
                    logging.debug(f"Invalid template directory: {template_dir}")
        except (FileNotFoundError, json.JSONDecodeError, KeyError) as e:
            logging.debug(f"Config file issue: {str(e)}")
        self.template_dir = None

    def prompt_template_dir(self):
        """Prompt user to select the root template directory."""
        messagebox.showinfo(
            "Select Template Directory",
            "Please select the root template directory containing 'banks', 'inter', and 'tsp' subdirectories with .docx templates."
        )
        template_dir = filedialog.askdirectory(
            title="Select Template Directory",
            initialdir=os.path.join(Path.home(), 'Documents')
        )
        if not template_dir:
            logging.warning("User cancelled template directory selection")
            return None
        # Verify the template directory contains required subdirectories and files
        subdirs = ['banks', 'inter', 'tsp']
        if not all(os.path.isdir(os.path.join(template_dir, subdir)) for subdir in subdirs):
            logging.error(f"Selected directory {template_dir} lacks required subdirectories")
            messagebox.showerror(
                "Error",
                "Selected directory must contain 'banks', 'inter', and 'tsp' subdirectories."
            )
            return None
        if not os.path.exists(os.path.join(template_dir, 'banks', 'bank.docx')):
            logging.error(f"Selected directory {template_dir}/banks lacks bank.docx")
            messagebox.showerror(
                "Error",
                "The 'banks' subdirectory must contain 'bank.docx'."
            )
            return None
        if not any(f.endswith('.docx') for f in os.listdir(os.path.join(template_dir, 'inter'))):
            logging.error(f"Selected directory {template_dir}/inter lacks .docx templates")
            messagebox.showerror(
                "Error",
                "The 'inter' subdirectory must contain at least one .docx template."
            )
            return None
        if not any(f.endswith('.docx') for f in os.listdir(os.path.join(template_dir, 'tsp'))):
            logging.error(f"Selected directory {template_dir}/tsp lacks .docx templates")
            messagebox.showerror(
                "Error",
                "The 'tsp' subdirectory must contain at least one .docx template."
            )
            return None
        logging.debug(f"User selected template directory: {template_dir}")
        return template_dir

    def save_template_dir(self, template_dir):
        """Save template directory to config file."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump({'template_dir': template_dir}, f)
            logging.debug(f"Saved template directory to config: {template_dir}")
        except Exception as e:
            logging.error(f"Failed to save template directory to config: {str(e)}")
            messagebox.showerror("Error", f"Failed to save template directory: {str(e)}")

    def change_template_dir(self):
        current_config = self.load_config()
        current_dir = current_config.get('template_dir', '')

        popup = tk.Toplevel(self.root)
        popup.title("Select Template Directory")
        popup.geometry("600x250")
        popup.resizable(False, False)
        popup.grab_set()
        popup.transient(self.root)

        tk.Label(popup, text="Current Template Directory:", font=("Arial", 12, "bold")).pack(pady=(15, 5))
        
        dir_var = tk.StringVar(value=current_dir)
        dir_label = tk.Label(popup, textvariable=dir_var, font=("Arial", 10), wraplength=480, justify="left")
        dir_label.pack(pady=5, padx=10)

        def select_new_dir():
            new_dir = filedialog.askdirectory(title="Select Template Directory")
            if new_dir:
                dir_var.set(new_dir)

        select_button = ttk.Button(popup, text=" select/change Directory", command=select_new_dir)
        select_button.pack(pady=10)

        def save_and_close():
            selected_dir = dir_var.get()
            if not selected_dir or not os.path.isdir(selected_dir):
                messagebox.showerror("Error", "Please select a valid directory.", parent=popup)
                return
            
            # Save to config
            config_data = self.load_config()
            config_data['template_dir'] = selected_dir
            self.save_config(config_data)
            
            # Update in-memory and sync with bank_letters
            self.template_dir = selected_dir
            self.bank_letters.bank_template_dir = os.path.join(selected_dir, 'banks')  # Assumes 'banks' subfolder
            
            messagebox.showinfo("Saved", f"Template directory saved:\n{selected_dir}", parent=popup)
            popup.destroy()

        save_button = ttk.Button(popup, text="Save", command=save_and_close)
        save_button.pack(pady=10)


    def set_theme(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 11, "bold"), padding=10, background=self.accent_color, foreground="white")
        style.map("TButton", background=[("active", "#0056b3")])
        style.configure("TCombobox", font=("Segoe UI", 10), padding=5, fieldbackground=self.bg_color, foreground=self.text_color)
        style.configure("TEntry", font=("Segoe UI", 10), padding=5, fieldbackground=self.bg_color, foreground=self.text_color)
        style.configure("TLabelFrame.Label", font=("Segoe UI", 13, "bold"), foreground=self.header_bg)

        self.root.configure(bg=self.bg_color)
        for frame in [self.bank_frame, self.inter_frame, self.tsp_frame]:
            frame.configure(bg=self.bg_color)
        self.bank_letters.bank_status_label.configure(bg="white", fg=self.success_color)
        self.inter_letters.inter_status_label.configure(bg="white", fg=self.success_color)
        self.tsp_letters.tsp_status_label.configure(bg="white", fg=self.success_color)
        self.case_details_label.configure(bg=self.bg_color, fg=self.text_color)
        self.bank_letters.excel_label.configure(bg="white", fg=self.text_color)
        for frame in [self.bank_frame, self.inter_frame, self.tsp_frame]:
            for widget in frame.winfo_children():
                if isinstance(widget, tk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Label):
                            child.configure(bg="white", fg=self.text_color)
                        elif isinstance(child, tk.Radiobutton):
                            child.configure(bg="white", fg=self.text_color)

    def next_tab(self):
        current = self.notebook.index(self.notebook.select())
        next_tab = (current + 1) % self.notebook.index("end")
        self.notebook.select(next_tab)

    def prev_tab(self):
        current = self.notebook.index(self.notebook.select())
        prev_tab = (current - 1) % self.notebook.index("end")
        self.notebook.select(prev_tab)

    def show_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Help - Letter Generator")
        help_window.geometry("600x400")
        help_window.resizable(False, False)
        help_window.transient(self.root)
        help_window.grab_set()
        help_window.configure(bg=self.bg_color)

        text = Text(help_window, font=("Segoe UI", 10), bg=self.bg_color, fg=self.text_color, wrap="word")
        text.pack(pady=10, padx=10, fill="both", expand=True)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=text.yview)
        scrollbar.pack(side="right", fill="y")
        text.configure(yscrollcommand=scrollbar.set)

        help_content = """
        # Letter Generator Help

        ## Overview
        This application generates letters for the Cyber Crime Wing, including Bank Letters, Intermediary Letters, and TSP Letters.

        ## Case Details
        - **Crime Number**: Enter in format XX/YYYY (e.g., 21/2025).
        - **NCRP ID**: Must be a 14-digit number.
        - Use the "Load Recent Cases" dropdown to select previously entered cases.
        - Shortcut: Ctrl+C to open case details.
        - Case details must be entered before generating any letters.

        ## Template Directory
        - Use the "Template Directory" button in the header to set the root template directory containing 'banks', 'inter', and 'tsp' subdirectories with .docx templates.
        - You will be prompted to select a template directory if none is set or if the current directory is invalid when generating a letter.

        ## Bank Letters
        - Select an Excel file with a 'Layer1' sheet containing columns: account_no, ifsc_code, transaction_amount, date_from, date_to, transaction_id_/_utr_number2, bank/fis.
        - Click "Generate Letters" (Ctrl+G) to create letters in the 'generated_letters/bank' folder.
        - View errors in the error log window if issues occur.

        ## Intermediary Letters
        - Select a platform (e.g., WhatsApp, Google) and enter URLs or IDs.
        - For Google, choose Gmail ID or GAID and enter the corresponding ID.
        - Generate letters in the 'generated_letters/inter' folder.

        ## TSP Letters
        - Select a TSP (e.g., Airtel) and optionally enter Phone or IMEI numbers.
        - Choose request type (CAF, CDR, Both) and enter dates in YYYY-MM-DD format for CAF or Both.
        - If tkcalendar is installed, use the calendar widget for date selection.
        - Generate letters in the 'generated_letters/tsp' folder.

        ## Shortcuts
        - Ctrl+C: Open case details.
        - Ctrl+G: Generate bank letters (when enabled).
        - Ctrl+T: Switch to next tab.
        - Ctrl+Shift+T: Switch to previous tab.

        ## Tips
        - Ensure the template directory contains 'banks', 'inter', and 'tsp' subdirectories with required .docx files.
        - Check the error log for detailed error messages.
        - Install tkcalendar (pip install tkcalendar) for a date picker in TSP and Intermediary Letters.
        """
        text.insert(tk.END, help_content)
        text.config(state="disabled")

    def update_button_states(self):
        has_case = self.crime_number and self.ncrp_id
        self.bank_letters.generate_button.config(state="normal" if has_case and getattr(self.bank_letters, 'selected_file', None) else "disabled")
        self.inter_letters.inter_generate_button.config(state="normal" if has_case else "disabled")
        self.tsp_letters.tsp_generate_button.config(state="normal" if has_case else "disabled")

    def prompt_case_details(self, anchor_widget):
        widget_x = anchor_widget.winfo_rootx()
        widget_y = anchor_widget.winfo_rooty()
        offset_x = widget_x + anchor_widget.winfo_width() + 10
        offset_y = widget_y

        dialog = tk.Toplevel(self.root)
        dialog.title("Enter Case Details")
        dialog.geometry(f"300x350+{offset_x}+{offset_y}")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.bg_color)

        recent_cases = []
        conn = connect_db()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT CrimeNumber, NCRP_ID FROM Cases ORDER BY id DESC LIMIT 5")
                recent_cases = [f"{row[0]} | {row[1]}" for row in cursor.fetchall()]
            finally:
                conn.close()

        tk.Label(dialog, text="Recent Cases:", font=("Segoe UI", 10, "bold"), bg=self.bg_color, fg=self.text_color).pack(pady=5)
        recent_combo = ttk.Combobox(dialog, values=recent_cases, state="readonly", style="TCombobox")
        recent_combo.pack(pady=5)
        self.ToolTip(recent_combo, "Select a recent case to autofill", self)

        tk.Label(dialog, text="Crime Number:", font=("Segoe UI", 10, "bold"), bg=self.bg_color, fg=self.text_color).pack(pady=5)
        self.crime_no_entry = ttk.Entry(dialog, font=("Segoe UI", 10), style="TEntry")
        self.crime_no_entry.pack(pady=5)
        crime_no_feedback = tk.Label(dialog, text="", font=("Segoe UI", 9), bg=self.bg_color, fg=self.error_color)
        crime_no_feedback.pack()

        tk.Label(dialog, text="NCRP ID:", font=("Segoe UI", 10, "bold"), bg=self.bg_color, fg=self.text_color).pack(pady=5)
        self.ncrp_id_entry = ttk.Entry(dialog, font=("Segoe UI", 10), style="TEntry")
        self.ncrp_id_entry.pack(pady=5)
        ncrp_id_feedback = tk.Label(dialog, text="", font=("Segoe UI", 9), bg=self.bg_color, fg=self.error_color)
        ncrp_id_feedback.pack()

        def validate_crime_no(event=None):
            value = self.crime_no_entry.get().strip()
            if not value:
                crime_no_feedback.config(text="Required field", fg=self.error_color)
            else:
                crime_no_feedback.config(text="Valid", fg=self.success_color)
            # Removed: Format check and specific error message (allows any non-empty value)
            # Original elif and else blocks commented out for reference
            # elif re.match(r'^\d{2}/\d{4}$', value):
            #     crime_no_feedback.config(text="Valid format", fg=self.success_color)
            # else:
            #     crime_no_feedback.config(text="Use XX/YYYY (e.g., 21/2025)", fg=self.error_color)


        def validate_ncrp_id(event=None):
            value = self.ncrp_id_entry.get().strip()
            if not value:
                ncrp_id_feedback.config(text="Required field", fg=self.error_color)
            elif len(value) == 14 and value.isdigit():
                ncrp_id_feedback.config(text="Valid format", fg=self.success_color)
            else:
                ncrp_id_feedback.config(text="Must be 14 digits", fg=self.error_color)

        self.crime_no_entry.bind("<KeyRelease>", validate_crime_no)
        self.ncrp_id_entry.bind("<KeyRelease>", validate_ncrp_id)

        def load_recent_case(event=None):
            selection = recent_combo.get()
            if selection:
                crime_no, ncrp_id = selection.split(" | ")
                self.crime_no_entry.delete(0, tk.END)
                self.crime_no_entry.insert(0, crime_no)
                self.ncrp_id_entry.delete(0, tk.END)
                self.ncrp_id_entry.insert(0, ncrp_id)
                validate_crime_no()
                validate_ncrp_id()

        recent_combo.bind("<<ComboboxSelected>>", load_recent_case)

        case_details = {'CrimeNumber': None, 'NCRP_ID': None}

        def submit():
            crime_no = self.crime_no_entry.get().strip()
            ncrp_id = self.ncrp_id_entry.get().strip()
            if not crime_no or not ncrp_id:
                messagebox.showerror("Error", "Crime Number and NCRP ID are required.")
                return

            # Removed: Format check for NCRP ID (allows any non-empty value)
            # Original: if len(ncrp_id) != 14 or not ncrp_id.isdigit():
            #     messagebox.showerror("Error", "NCRP ID must be a 14-digit number.")
            #     return

            if len(ncrp_id) != 14 or not ncrp_id.isdigit():
                messagebox.showerror("Error", "NCRP ID must be a 14-digit number.")
                return
            try:
                conn = connect_db()
                if not conn:
                    messagebox.showerror("Error", "Database connection failed.")
                    return
                cursor = conn.cursor()
                cursor.execute("INSERT INTO Cases (CrimeNumber, NCRP_ID) VALUES (?, ?)", (crime_no, ncrp_id))
                conn.commit()
                case_details['CrimeNumber'] = crime_no
                case_details['NCRP_ID'] = ncrp_id
                self.last_case_details['CrimeNumber'] = crime_no
                self.last_case_details['NCRP_ID'] = ncrp_id
                messagebox.showinfo("Success", "Case details saved successfully.")
                dialog.destroy()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"An error occurred: {e}")
            finally:
                if conn:
                    conn.close()

        ttk.Button(dialog, text="Submit", command=submit, style="TButton").pack(pady=10)
        dialog.wait_window()
        return case_details

    def set_case_details(self):
        case_details = self.prompt_case_details(self.case_details_button)
        if case_details and case_details.get('CrimeNumber') and case_details.get('NCRP_ID'):
            self.crime_number = case_details['CrimeNumber']
            self.ncrp_id = case_details['NCRP_ID']
            self.case_details_label.config(
                text=f"Case Details: Crime No: {self.crime_number}, NCRP ID: {self.ncrp_id}"
            )
            self.update_button_states()
        else:
            self.crime_number = None
            self.ncrp_id = None
            self.case_details_label.config(text="Case Details: Not Set")
            self.update_button_states()

    def toggle_profile(self):
        if self.profile_window and self.profile_window.winfo_exists():
            self.profile_window.destroy()

        self.profile_window = tk.Toplevel(self.root)
        self.profile_window.title("Edit Profile")
        self.profile_window.geometry("400x400")
        self.profile_window.resizable(False, False)
        self.profile_window.transient(self.root)
        self.profile_window.grab_set()
        self.profile_window.configure(bg=self.bg_color)

        self.root.update_idletasks()

        help_x = self.help_button.winfo_rootx()
        help_y = self.help_button.winfo_rooty()
        help_height = self.help_button.winfo_height()

        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()

        popup_width = 400
        popup_height = 400
        right_margin = 40
        x = root_x + root_width - popup_width - right_margin
        offset_y = help_y + help_height + 10
        self.profile_window.geometry(f"{popup_width}x{popup_height}+{x}+{offset_y}")

        conn = connect_db()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT Username, OfficerName, Designation, Phone, Email FROM Officers WHERE Id = ?",
                    (self.officer['Id'],)
                )
                officer_data = cursor.fetchone()
                if officer_data:
                    self.officer.update({
                        'Username': officer_data[0], 'OfficerName': officer_data[1],
                        'Designation': officer_data[2], 'Phone': officer_data[3],
                        'Email': officer_data[4]
                    })
            finally:
                conn.close()

        ttk.Label(self.profile_window, text="Edit Profile", font=("Segoe UI", 16, "bold")).pack(pady=15)
        form_frame = tk.Frame(self.profile_window, bg=self.bg_color)
        form_frame.pack(fill="both", expand=True, padx=20)

        fields = ['Username', 'OfficerName', 'Designation', 'Phone', 'Email']
        self.entries = {}
        for field in fields:
            field_label = field.replace('OfficerName', 'Name').capitalize()
            row = tk.Frame(form_frame, bg=self.bg_color)
            row.pack(fill="x", pady=8)
            tk.Label(
                row, text=f"{field_label}:", font=("Segoe UI", 10, "bold"), width=12, anchor="w",
                bg=self.bg_color, fg=self.text_color
            ).pack(side="left")
            entry = ttk.Entry(row, font=("Segoe UI", 10), style="TEntry")
            entry.insert(0, self.officer.get(field, ''))
            entry.pack(side="right", fill="x", expand=True)
            self.entries[field] = entry

        save_button = ttk.Button(
            self.profile_window, text="Save Changes", command=self.save_profile, style="TButton"
        )
        save_button.pack(pady=20)
        self.ToolTip(save_button, "Save the updated profile details", self)


    def logout(self):
        if messagebox.askokcancel("Logout", "Are you sure you want to log out?"):
            self.root.destroy()
            from main import main
            main()



    def save_profile(self):
        updated_officer = {
            'Id': self.officer['Id'], 'Username': self.entries['Username'].get(),
            'OfficerName': self.entries['OfficerName'].get(), 'Designation': self.entries['Designation'].get(),
            'Phone': self.entries['Phone'].get(), 'Email': self.entries['Email'].get()
        }
        conn = connect_db()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE Officers SET Username = ?, OfficerName = ?, Designation = ?, Phone = ?, Email = ? WHERE Id = ?",
                    (updated_officer['Username'], updated_officer['OfficerName'], updated_officer['Designation'],
                     updated_officer['Phone'], updated_officer['Email'], updated_officer['Id'])
                )
                conn.commit()
                self.officer.update(updated_officer)
                messagebox.showinfo("Success", "Profile updated successfully!")
                self.profile_window.destroy()
                self.root.title(f"Cyber Crime Wing - Letter Generator ({self.officer['OfficerName']})")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update profile: {str(e)}")
            finally:
                conn.close()

    def fetch_officer_details(self):
        conn = connect_db()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT Username, OfficerName, Designation, Phone, Email FROM Officers WHERE Id = ?",
                    (self.officer['Id'],)
                )
                officer_data = cursor.fetchone()
                if officer_data:
                    self.officer.update({
                        'Username': officer_data[0], 'OfficerName': officer_data[1],
                        'Designation': officer_data[2], 'Phone': officer_data[3],
                        'Email': officer_data[4]
                    })
            except Exception as e:
                logging.error(f"Error fetching officer details: {str(e)}")
            finally:
                conn.close()

    def save_config(self, config_data):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config_data, f)
        except Exception as e:
            print(f"Error saving config: {e}")

    def load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return {}
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
            return {}


    def show_error_log(self, errors):
        error_window = tk.Toplevel(self.root)
        error_window.title("Error Log")
        error_window.geometry("600x400")
        error_window.resizable(False, False)
        error_window.transient(self.root)
        error_window.grab_set()
        error_window.configure(bg=self.bg_color)

        text = Text(error_window, font=("Segoe UI", 10), bg=self.bg_color, fg=self.text_color, wrap="word")
        text.pack(pady=10, padx=10, fill="both", expand=True)
        scrollbar = ttk.Scrollbar(error_window, orient="vertical", command=text.yview)
        scrollbar.pack(side="right", fill="y")
        text.configure(yscrollcommand=scrollbar.set)

        for error in errors:
            text.insert(tk.END, f"{error}\n")
        text.config(state="disabled")

        def copy_to_clipboard():
            self.root.clipboard_clear()
            self.root.clipboard_append(text.get("1.0", tk.END))
            messagebox.showinfo("Success", "Error log copied to clipboard")

        ttk.Button(error_window, text="Copy to Clipboard", command=copy_to_clipboard, style="TButton").pack(pady=10)

    def logout(self):
        self.root.destroy()
        from main import main
        main()