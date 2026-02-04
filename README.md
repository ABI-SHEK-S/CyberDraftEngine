# 🚔 CyberDraft Engine – Automated Letter Generator for Cyber Crime Wing

## 📌 Overview
CyberDraft Engine is a desktop-based automation tool developed to assist law enforcement officers in generating formal investigation letters quickly and accurately.  

The application reduces manual paperwork, minimizes human error, and significantly speeds up communication with Telecom Service Providers (TSPs), banks, and digital intermediaries.

Designed with usability and efficiency in mind, this system streamlines critical workflows involved in cybercrime investigations.

---

## 🎯 Objective
To build a secure, user-friendly GUI application that automates official letter generation for cybercrime investigations, improving operational efficiency for officers.

---

## 🚀 Key Features

### 🔐 Secure Authentication
- Role-based login system (Admin & Officer)
- Password hashing using **bcrypt**
- Officer profile management

### 📄 Automated Letter Generation
Supports multiple investigation request types:

✅ Telecom Service Providers (Airtel, Jio, Vodafone, BSNL)
- CAF Details  
- Call Detail Records (CDR)  
- IMEI Tracking  
- Aadhar-linked numbers  
- PoS requests  

✅ Intermediary Platforms
- WhatsApp  
- Google  
- Facebook  
- Instagram  
- Twitter  

✅ Banking Requests
- Bulk processing via Excel  
- Transaction and account detail extraction  
- Batch letter generation with progress tracking  

---

## 🏗️ System Architecture

**Frontend:**  
- Python Tkinter (GUI)

**Backend:**  
- SQLite3 (Local Database)

**Libraries Used:**
- `python-docx` → Word document generation  
- `pandas` → Excel data processing  
- `bcrypt` → Password security  
- `tkcalendar` → Date selection (optional)

---

## 🧠 How It Works
1. Officer logs into the system.
2. Creates or selects a case.
3. Chooses the request type (TSP / Bank / Intermediary).
4. Inputs required details.
5. The system dynamically fills pre-built templates.
6. A formatted `.docx` letter is generated instantly.

---

## 📂 Project Structure

letter_generator/
│
├── main.py # Application entry point
├── main_app.py # Core GUI logic
├── login_window.py # Authentication interface
├── admin_panel.py # Officer & template management
├── database.py # Database creation & operations
│
├── tsp_letters.py # Telecom request module
├── inter_letters.py # Intermediary request module
├── bank_letters.py # Banking request module
│
├── templates/ # Letter templates
├── generated_letters/ # Output folder
├── db/ # SQLite database
└── utils.py # Helper functions


---

## 💡 Why This Project Matters
Manual letter drafting is time-consuming during investigations.

CyberDraft Engine:

✅ Reduces paperwork  
✅ Saves officer time  
✅ Ensures formatting consistency  
✅ Minimizes errors  
✅ Enables faster case processing  

---

## ⚙️ Installation

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/your-username/CyberDraftEngine.git
cd CyberDraftEngine
2️⃣ Install Dependencies
pip install python-docx pandas bcrypt tkcalendar
3️⃣ Run the Application
python main.py
🔐 Security Considerations
Password hashing implemented

Structured database design

Role-based system

Input validation included

(Future enhancements will include stronger sanitization and access control.)

🚧 Current Limitations
Limited edge-case testing

Styling can be further improved

Multi-date support not yet implemented

Advanced error handling in progress

🔮 Future Enhancements
Cloud database integration

Advanced role-based access

Improved UI styling

Export to PDF

Automated audit logs

Multi-date investigation support

🏆 What This Project Demonstrates
This project highlights:

✔ Real-world problem solving
✔ Secure authentication
✔ Database design
✔ GUI development
✔ Document automation
✔ Modular architecture

👨‍💻 Author
Dhanush V N

If you found this project interesting, feel free to connect or provide feedback!
