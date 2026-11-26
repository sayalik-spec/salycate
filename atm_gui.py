import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from openpyxl import Workbook, load_workbook
from fpdf import FPDF
from datetime import datetime
import os
import subprocess

# Dummy accounts
accounts = {
    "123456": {"pin": "1111", "balance": 50000},
    "987654": {"pin": "2222", "balance": 17000},
    "244210": {"pin": "2814", "balance": 97500},
    "244225": {"pin": "2508", "balance": 5000},
}

class ATM_GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ATM System")
        self.root.geometry("700x550")
        self.root.resizable(False, False)

        # Load background
        bg_image = Image.open("background.jpg").resize((700, 550))
        self.bg_photo = ImageTk.PhotoImage(bg_image)

        self.bg_label = tk.Label(self.root, image=self.bg_photo)
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # Load logo
        self.logo_photo = None
        if os.path.exists("logo.png"):
            logo_image = Image.open("logo.png").resize((60, 60))
            self.logo_photo = ImageTk.PhotoImage(logo_image)

        self.login_screen()

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        self.bg_label = tk.Label(self.root, image=self.bg_photo)
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    def login_screen(self):
        self.clear_frame()

        if self.logo_photo:
            tk.Label(self.root, image=self.logo_photo, bg="#f0f0f0").place(relx=0.5, rely=0.15, anchor="center")

        frame = tk.Frame(self.root, bg="#f0f0f0", bd=2, relief="ridge")
        frame.place(relx=0.5, rely=0.5, anchor="center", width=350, height=300)

        tk.Label(frame, text="ATM Login", font=("Arial", 20, "bold"), bg="#dbe9f4").pack(pady=5)

        tk.Label(frame, text="Account Number:", font=("Arial", 12), bg="#f0f0f0").pack(pady=5)
        self.acc_entry = tk.Entry(frame, font=("Arial", 12))
        self.acc_entry.pack(pady=5)

        tk.Label(frame, text="PIN:", font=("Arial", 12), bg="#f0f0f0").pack(pady=5)
        self.pin_entry = tk.Entry(frame, show="*", font=("Arial", 12))
        self.pin_entry.pack(pady=5)

        tk.Button(frame, text="Login", font=("Arial", 14, "bold"), bg="green", fg="white", command=self.authenticate).pack(pady=15)

    def authenticate(self):
        acc = self.acc_entry.get()
        pin = self.pin_entry.get()

        if acc in accounts and accounts[acc]["pin"] == pin:
            self.current_account = acc
            self.dashboard_screen()
        else:
            messagebox.showerror("Login Failed", "Invalid account number or PIN.")

    def dashboard_screen(self):
        self.clear_frame()

        frame = tk.Frame(self.root, bg="#ffffff", bd=2, relief="ridge")
        frame.place(relx=0.5, rely=0.5, anchor="center", width=400, height=470)

        tk.Label(frame, text="ATM Dashboard", font=("Arial", 20, "bold"), bg="#ffffff").pack(pady=10)

        balance = accounts[self.current_account]['balance']
        tk.Label(frame, text=f"Balance: ₹ {balance:.2f}", font=("Arial", 16, "bold"), fg="green", bg="#ffffff").pack(pady=5)

        tk.Button(frame, text="Deposit", font=("Arial", 14), width=25, bg="blue", fg="white", command=self.deposit_screen).pack(pady=5)
        tk.Button(frame, text="Withdraw", font=("Arial", 14), width=25, bg="red", fg="white", command=self.withdraw_screen).pack(pady=5)
        tk.Button(frame, text="Transfer", font=("Arial", 14), width=25, bg="orange", fg="white", command=self.transfer_screen).pack(pady=5)
        tk.Button(frame, text="View Transactions", font=("Arial", 14), width=25, bg="#555", fg="white", command=self.open_excel).pack(pady=5)
        tk.Button(frame, text="Export to PDF", font=("Arial", 14), width=25, bg="#007bff", fg="white", command=self.export_to_pdf).pack(pady=5)
        tk.Button(frame, text="Logout", font=("Arial", 14), width=25, bg="gray", fg="white", command=self.login_screen).pack(pady=10)

    def deposit_screen(self):
        self.transaction_screen("Deposit", "Enter deposit amount:", self.deposit)

    def withdraw_screen(self):
        self.transaction_screen("Withdraw", "Enter withdrawal amount:", self.withdraw)

    def transfer_screen(self):
        self.transaction_screen("Transfer", "Enter amount and recipient account:", self.transfer, is_transfer=True)

    def transaction_screen(self, title, label, action, is_transfer=False):
        self.clear_frame()

        frame = tk.Frame(self.root, bg="#ffffff", bd=2, relief="ridge")
        frame.place(relx=0.5, rely=0.5, anchor="center", width=400, height=300)

        tk.Label(frame, text=title, font=("Arial", 18, "bold"), bg="#ffffff").pack(pady=10)
        tk.Label(frame, text=label, font=("Arial", 12), bg="#ffffff").pack(pady=5)

        self.amount_entry = tk.Entry(frame, font=("Arial", 12))
        self.amount_entry.pack(pady=5)

        if is_transfer:
            tk.Label(frame, text="Recipient Account Number:", font=("Arial", 12), bg="#ffffff").pack(pady=5)
            self.recipient_entry = tk.Entry(frame, font=("Arial", 12))
            self.recipient_entry.pack(pady=5)

        tk.Button(frame, text="Confirm", font=("Arial", 12), bg="green", fg="white", command=action).pack(pady=10)
        tk.Button(frame, text="Back", font=("Arial", 12), bg="gray", fg="white", command=self.dashboard_screen).pack()

    def deposit(self):
        amount = self.get_amount()
        if amount:
            accounts[self.current_account]["balance"] += amount
            self.log_transaction(self.current_account, "Deposit", amount)
            messagebox.showinfo("Success", f"Deposited ₹.{amount}")
            self.dashboard_screen()

    def withdraw(self):
        amount = self.get_amount()
        if amount and amount <= accounts[self.current_account]["balance"]:
            accounts[self.current_account]["balance"] -= amount
            self.log_transaction(self.current_account, "Withdraw", amount)
            messagebox.showinfo("Success", f"Withdrew ₹.{amount}")
            self.dashboard_screen()
        else:
            messagebox.showerror("Error", "Insufficient funds or invalid amount.")

    def transfer(self):
        amount = self.get_amount()
        recipient = self.recipient_entry.get()

        if recipient == self.current_account:
            messagebox.showerror("Error", "Cannot transfer to the same account.")
            return

        if recipient in accounts and amount and amount <= accounts[self.current_account]["balance"]:
            accounts[self.current_account]["balance"] -= amount
            accounts[recipient]["balance"] += amount
            self.log_transaction(self.current_account, "Transfer", amount, recipient)
            self.log_transaction(recipient, "Received", amount, self.current_account)
            messagebox.showinfo("Success", f"Transferred ₹.{amount} to {recipient}")
            self.dashboard_screen()
        else:
            messagebox.showerror("Error", "Invalid recipient or insufficient funds.")

    def get_amount(self):
        try:
            amount = float(self.amount_entry.get())
            if amount > 0:
                return amount
            else:
                messagebox.showerror("Error", "Amount must be greater than zero.")
                return None
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number.")
            return None

    def log_transaction(self, account, type_, amount, recipient=None):
        file_name = "transactions.xlsx"

        if not os.path.exists(file_name):
            wb = Workbook()
            ws = wb.active
            ws.append(["Account", "Type", "Amount", "Recipient", "Balance", "Timestamp"])
            wb.save(file_name)

        wb = load_workbook(file_name)
        ws = wb.active

        balance = accounts[account]['balance']
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append([account, type_, amount, recipient or "-", balance, timestamp])
        wb.save(file_name)

    def open_excel(self):
        file_name = "transactions.xlsx"
        if os.path.exists(file_name):
            subprocess.Popen(["start", file_name], shell=True)
        else:
            messagebox.showinfo("Info", "No transactions found yet.")

    def export_to_pdf(self):
        file_name = "transactions.xlsx"
        if not os.path.exists(file_name):
            messagebox.showinfo("Info", "No transactions to export.")
            return

        wb = load_workbook(file_name)
        ws = wb.active

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, f"Transaction History for {self.current_account}", ln=True, align="C")
        pdf.ln(10)

        headers = ["Account", "Type", "Amount", "Recipient", "Balance", "Timestamp"]
        for header in headers:
            pdf.cell(32, 10, header, 1)
        pdf.ln()

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == self.current_account:
                for item in row:
                    pdf.cell(32, 10, str(item), 1)
                pdf.ln()

        pdf_file = f"transactions_{self.current_account}.pdf"
        pdf.output(pdf_file)
        subprocess.Popen(["start", pdf_file], shell=True)

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = ATM_GUI(root)
    root.mainloop()
