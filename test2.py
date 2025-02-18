import sqlite3

def check_users():
    connection = sqlite3.connect("participants.db")
    cursor = connection.cursor()

    cursor.execute("SELECT username, password FROM users")
    users = cursor.fetchall()
    connection.close()

    if users:
        for user in users:
            print(f"Username: {user[0]}, Hashed Password: {user[1]}")
    else:
        print("⚠️ No users found in database!")

check_users()
