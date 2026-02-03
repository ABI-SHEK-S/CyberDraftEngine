import bcrypt
from db.database import connect_db

conn = connect_db()
cursor = conn.cursor()
hashed_pw = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt())
cursor.execute('''
    INSERT OR IGNORE INTO Officers (Username, Password, OfficerName, Designation, Phone, Email, Address)
    VALUES (?, ?, ?, ?, ?, ?, ?)
''', ('admin', hashed_pw, 'Administrator', 'Admin', '0000000000', 'admin@example.com', 'Head Office'))
conn.commit()
conn.close()
