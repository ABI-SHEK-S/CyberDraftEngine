import sqlite3
import os
import sys
import shutil
from pathlib import Path
import logging
import bcrypt

# Set up logging
log_dir = os.path.join(Path.home(), 'Documents', 'LetterGeneratorLogs')
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    filename=os.path.join(log_dir, 'database.log'),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def connect_db():
    """Connect to the SQLite database with proper error handling."""
    try:
        # Define writable database path
        user_db_dir = os.path.join(Path.home(), 'Documents', 'LetterGeneratorData')
        user_db_path = os.path.join(user_db_dir, 'letter_requests.db')
        
        if getattr(sys, 'frozen', False):
            # Running as executable
            base_path = sys._MEIPASS
            bundled_db_path = os.path.join(base_path, 'db', 'letter_requests.db')
            # Copy bundled database to user directory if it doesn't exist
            if not os.path.exists(user_db_path):
                os.makedirs(user_db_dir, exist_ok=True)
                if os.path.exists(bundled_db_path):
                    shutil.copyfile(bundled_db_path, user_db_path)
                    logging.debug(f"Copied database from {bundled_db_path} to {user_db_path}")
                else:
                    logging.error(f"Bundled database not found: {bundled_db_path}")
                    return None
        else:
            # Running in development
            user_db_path = os.path.join(Path(__file__).parent.parent, 'db', 'letter_requests.db')
            logging.debug(f"Using development database path: {user_db_path}")
        
        # Ensure the database directory exists
        os.makedirs(os.path.dirname(user_db_path), exist_ok=True)
        # Use timeout to prevent locking issues
        conn = sqlite3.connect(user_db_path, timeout=10)
        conn.execute("PRAGMA foreign_keys = ON")
        # Verify Officers table exists and has data
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Officers'")
        if not cursor.fetchone():
            logging.warning("Officers table not found, creating it")
            cursor.execute("""
                CREATE TABLE Officers (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Username TEXT UNIQUE NOT NULL,
                    Password TEXT NOT NULL,
                    OfficerName TEXT NOT NULL,
                    Designation TEXT NOT NULL,
                    Phone TEXT NOT NULL,
                    Email TEXT NOT NULL
                )
            """)
            conn.commit()
            # Create default admin user
            hashed_pw = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt())
            cursor.execute(
                "INSERT OR IGNORE INTO Officers (Username, Password, OfficerName, Designation, Phone, Email) VALUES (?, ?, ?, ?, ?, ?)",
                ('admin', hashed_pw, 'Administrator', 'Admin', '0000000000', 'admin@example.com')
            )
            conn.commit()
            logging.debug("Created default admin user")
        cursor.execute("SELECT COUNT(*) FROM Officers")
        count = cursor.fetchone()[0]
        logging.debug(f"Officers table contains {count} records")
        return conn
    except sqlite3.Error as e:
        logging.error(f"Database connection error: {e}")
        return None
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return None

def create_database():
    """Create necessary database tables."""
    conn = None
    try:
        conn = connect_db()
        if not conn:
            logging.error("Failed to connect to database for table creation")
            return
        
        cursor = conn.cursor()
        
        # Create Cases table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Cases (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                CrimeNumber TEXT NOT NULL,
                NCRP_ID TEXT NOT NULL,
                CreatedAt TEXT DEFAULT (datetime('now'))
            )
        """)
        
        # Create OTPs table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS OTPs (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Username TEXT NOT NULL,
                OTP TEXT NOT NULL,
                CreatedAt TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (Username) REFERENCES Officers(Username)
            )
        """)
        
        conn.commit()
        logging.debug("Database tables created successfully")
        
    except sqlite3.Error as e:
        logging.error(f"Error creating database: {e}")
    finally:
        if conn:
            conn.close()

def save_case(case, officer_id, letter_type):
    """Save case data to database."""
    conn = None
    try:
        conn = connect_db()
        if not conn:
            logging.error("Database connection failed")
            return "Database connection failed"
        
        cursor = conn.cursor()
        
        # Check if case already exists
        cursor.execute("SELECT Id FROM Cases WHERE CrimeNumber = ? AND NCRP_ID = ?", 
                      (case.get('CrimeNumber', ''), case.get('NCRP_ID', '')))
        case_row = cursor.fetchone()
        
        if case_row:
            case_id = case_row[0]
        else:
            cursor.execute("INSERT INTO Cases (CrimeNumber, NCRP_ID) VALUES (?, ?)",
                          (case.get('CrimeNumber', ''), case.get('NCRP_ID', '')))
            case_id = cursor.lastrowid
        
        conn.commit()
        logging.debug(f"Saved case ID {case_id} for CrimeNumber {case['CrimeNumber']}")
        return None
        
    except sqlite3.Error as e:
        logging.error(f"Database error: {str(e)}")
        return f"Database error: {str(e)}"
    except Exception as e:
        logging.error(f"Unexpected error: {str(e)}")
        return f"Unexpected error: {str(e)}"
    finally:
        if conn:
            conn.close()

def create_default_admin():
    """Create default admin user."""
    # Moved to connect_db to ensure creation when Officers table is created
    pass

def validate_case(case):
    """Validate case data before saving."""
    errors = []
    
    # Basic validation
    if not case.get('CrimeNumber', '').strip():
        errors.append("Crime Number is required")
    if not case.get('NCRP_ID', '').strip():
        errors.append("NCRP ID is required")  # Keeps it required (non-empty) but allows any value
    
    # If there's a format check like this elsewhere in the function, comment it out:
    # if not re.match(r'^\d+/\d+$', case.get('NCRP_ID', '')):
    #     errors.append("NCRP ID must be in format like '21/5698'")
    
    # (Rest of the function unchanged, if any)
    return errors  # Assuming this is how errors are returned

# Initialize database when module is imported
if __name__ == "__main__":
    create_database()