"""
=====================================================
  PYTHON ATM MANAGEMENT SYSTEM (Excel Integrated)
=====================================================
Features:
  1. Create Account
  2. Deposit Money
  3. Withdraw Money
  4. Show Balance
  5. Exit

All account data is stored in an Excel file (atm_database.xlsx)
using the openpyxl library. Every transaction updates the Excel
sheet immediately, so the sheet always reflects the latest data.
"""

import os
import random
from openpyxl import Workbook, load_workbook

DB_FILE = "atm_database.xlsx"
SHEET_NAME = "Accounts"
HEADERS = ["Account Number", "Name", "Age", "Address", "Phone Number", "PIN", "Balance"]


# ---------------------------------------------------------
# 1. DATABASE (EXCEL) HELPER FUNCTIONS
# ---------------------------------------------------------

def init_database():
    """Create the Excel file with headers if it doesn't already exist."""
    if not os.path.exists(DB_FILE):
        wb = Workbook()
        sheet = wb.active
        sheet.title = SHEET_NAME
        sheet.append(HEADERS)
        wb.save(DB_FILE)


def load_sheet():
    """Load workbook and return (workbook, sheet)."""
    wb = load_workbook(DB_FILE)
    sheet = wb[SHEET_NAME]
    return wb, sheet


def generate_account_number(sheet):
    """Generate a unique 6-digit account number."""
    existing = {row[0].value for row in sheet.iter_rows(min_row=2) if row[0].value}
    while True:
        acc_no = random.randint(100000, 999999)
        if acc_no not in existing:
            return acc_no


def find_account_row(sheet, acc_no):
    """Return the row number (Excel row index) of an account, or None."""
    for row in sheet.iter_rows(min_row=2):
        if row[0].value == acc_no:
            return row[0].row
    return None


# ---------------------------------------------------------
# 2. CORE ATM FEATURES
# ---------------------------------------------------------

def create_account():
    print("\n----- CREATE NEW ACCOUNT -----")
    name = input("Enter your full name: ").strip()
    age = input("Enter your age: ").strip()
    address = input("Enter your address: ").strip()
    phone = input("Enter your phone number: ").strip()
    pin = input("Set a 4-digit PIN: ").strip()

    while not pin.isdigit() or len(pin) != 4:
        pin = input("Invalid PIN. Enter a 4-digit numeric PIN: ").strip()

    try:
        initial_deposit = float(input("Enter initial deposit amount (₹): ").strip())
        if initial_deposit < 0:
            raise ValueError
    except ValueError:
        print("Invalid amount. Setting initial balance to 0.")
        initial_deposit = 0.0

    wb, sheet = load_sheet()
    acc_no = generate_account_number(sheet)
    sheet.append([acc_no, name, age, address, phone, pin, initial_deposit])
    wb.save(DB_FILE)

    print("\n✅ Account created successfully!")
    print(f"   Your Account Number is: {acc_no}")
    print("   (Please note this down — you'll need it for all transactions)\n")


def authenticate():
    """Ask for account number + PIN, return (sheet, wb, row_number) if valid."""
    try:
        acc_no = int(input("Enter your Account Number: ").strip())
    except ValueError:
        print("❌ Invalid account number format.")
        return None, None, None

    wb, sheet = load_sheet()
    row_num = find_account_row(sheet, acc_no)

    if row_num is None:
        print("❌ Account not found.")
        return None, None, None

    pin = input("Enter your PIN: ").strip()
    stored_pin = str(sheet.cell(row=row_num, column=6).value)

    if pin != stored_pin:
        print("❌ Incorrect PIN.")
        return None, None, None

    return wb, sheet, row_num


def deposit_money():
    print("\n----- DEPOSIT MONEY -----")
    wb, sheet, row_num = authenticate()
    if row_num is None:
        return

    try:
        amount = float(input("Enter amount to deposit (₹): ").strip())
        if amount <= 0:
            raise ValueError
    except ValueError:
        print("❌ Invalid amount.")
        return

    balance_cell = sheet.cell(row=row_num, column=7)
    balance_cell.value = balance_cell.value + amount
    wb.save(DB_FILE)

    print(f"✅ ₹{amount:.2f} deposited successfully!")
    print(f"   New Balance: ₹{balance_cell.value:.2f}\n")


def withdraw_money():
    print("\n----- WITHDRAW MONEY -----")
    wb, sheet, row_num = authenticate()
    if row_num is None:
        return

    try:
        amount = float(input("Enter amount to withdraw (₹): ").strip())
        if amount <= 0:
            raise ValueError
    except ValueError:
        print("❌ Invalid amount.")
        return

    balance_cell = sheet.cell(row=row_num, column=7)

    if amount > balance_cell.value:
        print("❌ Insufficient balance!")
        print(f"   Current Balance: ₹{balance_cell.value:.2f}\n")
        return

    balance_cell.value = balance_cell.value - amount
    wb.save(DB_FILE)

    print(f"✅ ₹{amount:.2f} withdrawn successfully!")
    print(f"   New Balance: ₹{balance_cell.value:.2f}\n")


def show_balance():
    print("\n----- CHECK BALANCE -----")
    wb, sheet, row_num = authenticate()
    if row_num is None:
        return

    name = sheet.cell(row=row_num, column=2).value
    balance = sheet.cell(row=row_num, column=7).value
    print(f"\n👤 Account Holder: {name}")
    print(f"💰 Current Balance: ₹{balance:.2f}\n")


# ---------------------------------------------------------
# 3. MAIN MENU
# ---------------------------------------------------------

def main_menu():
    init_database()

    while True:
        print("=" * 40)
        print("        WELCOME TO PYTHON ATM")
        print("=" * 40)
        print("1. Create Account")
        print("2. Deposit Money")
        print("3. Withdraw Money")
        print("4. Show Balance")
        print("5. Exit")
        print("=" * 40)

        choice = input("Enter your choice (1-5): ").strip()

        if choice == "1":
            create_account()
        elif choice == "2":
            deposit_money()
        elif choice == "3":
            withdraw_money()
        elif choice == "4":
            show_balance()
        elif choice == "5":
            print("\nThank you for using Python ATM. Goodbye! 👋")
            break
        else:
            print("❌ Invalid choice. Please enter a number between 1 and 5.\n")


if __name__ == "__main__":
    main_menu()
    
