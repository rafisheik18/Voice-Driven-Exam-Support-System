import openpyxl

USERS_FILE = "users.xlsx"

def load_users():
    users = {}
    wb = openpyxl.load_workbook(USERS_FILE)
    sheet = wb.active
    for row in sheet.iter_rows(min_row=2, values_only=True):  # skip header
        username, password, role = row
        if username and password:
            users[username] = {"password": str(password), "role": role}
    return users

# preload users
users = load_users()
