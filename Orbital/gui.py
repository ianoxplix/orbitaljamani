# gui.py

import tkinter as tk
from tkinter import ttk, messagebox
from models import Customer, QuoteMeta, QuoteItem, Quote
from excel_writer import save_quote_to_excel

class QuoteGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Aluminium Quotation System")

        # --- Customer Info ---
        tk.Label(root, text="Customer Name").grid(row=0, column=0)
        self.name_var = tk.StringVar()
        tk.Entry(root, textvariable=self.name_var).grid(row=0, column=1)

        tk.Label(root, text="Address").grid(row=1, column=0)
        self.address_var = tk.StringVar()
        tk.Entry(root, textvariable=self.address_var).grid(row=1, column=1)

        tk.Label(root, text="Town").grid(row=2, column=0)
        self.town_var = tk.StringVar()
        tk.Entry(root, textvariable=self.town_var).grid(row=2, column=1)

        tk.Label(root, text="Contact Person").grid(row=3, column=0)
        self.contact_var = tk.StringVar()
        tk.Entry(root, textvariable=self.contact_var).grid(row=3, column=1)

        tk.Label(root, text="Phone").grid(row=4, column=0)
        self.phone_var = tk.StringVar()
        tk.Entry(root, textvariable=self.phone_var).grid(row=4, column=1)

        tk.Label(root, text="Email").grid(row=5, column=0)
        self.email_var = tk.StringVar()
        tk.Entry(root, textvariable=self.email_var).grid(row=5, column=1)

        # --- Quote Info ---
        tk.Label(root, text="Quotation No").grid(row=0, column=2)
        self.quotation_no_var = tk.StringVar()
        tk.Entry(root, textvariable=self.quotation_no_var).grid(row=0, column=3)

        tk.Label(root, text="Date").grid(row=1, column=2)
        self.date_var = tk.StringVar()
        tk.Entry(root, textvariable=self.date_var).grid(row=1, column=3)

        tk.Label(root, text="Order No").grid(row=2, column=2)
        self.order_no_var = tk.StringVar()
        tk.Entry(root, textvariable=self.order_no_var).grid(row=2, column=3)

        tk.Label(root, text="Service").grid(row=3, column=2)  # ✅ Service Field
        self.service_var = tk.StringVar()
        tk.Entry(root, textvariable=self.service_var).grid(row=3, column=3)

        # --- Items Section ---
        self.items = []

        self.tree = ttk.Treeview(root, columns=("Qty", "Description", "Units", "Unit Price", "Total"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.grid(row=6, column=0, columnspan=4)

        # --- Add Item Inputs ---
        self.qty_var = tk.StringVar()
        self.desc_var = tk.StringVar()
        self.units_var = tk.StringVar()
        self.unit_price_var = tk.StringVar()

        tk.Entry(root, textvariable=self.qty_var, width=5).grid(row=7, column=0)
        tk.Entry(root, textvariable=self.desc_var, width=20).grid(row=7, column=1)
        tk.Entry(root, textvariable=self.units_var, width=10).grid(row=7, column=2)
        tk.Entry(root, textvariable=self.unit_price_var, width=10).grid(row=7, column=3)

        tk.Button(root, text="Add Item", command=self.add_item).grid(row=7, column=4)

        # --- Labour & Total ---
        tk.Label(root, text="Labour Description").grid(row=8, column=0)
        self.labour_desc_var = tk.StringVar()
        tk.Entry(root, textvariable=self.labour_desc_var).grid(row=8, column=1)

        tk.Label(root, text="Labour Cost").grid(row=8, column=2)
        self.labour_cost_var = tk.StringVar()
        tk.Entry(root, textvariable=self.labour_cost_var).grid(row=8, column=3)

        tk.Label(root, text="Grand Total").grid(row=9, column=2)
        self.total_var = tk.StringVar()
        tk.Entry(root, textvariable=self.total_var).grid(row=9, column=3)

        # --- Save Button ---
        tk.Button(root, text="Save Quote", command=self.save_quote).grid(row=10, column=3)

    def add_item(self):
        try:
            qty = float(self.qty_var.get())
            unit_price = float(self.unit_price_var.get())
            total = qty * unit_price

            item = QuoteItem(
                qty=qty,
                description=self.desc_var.get(),
                units=self.units_var.get(),
                unit_price=unit_price,
                total=total
            )

            self.items.append(item)
            self.tree.insert("", "end", values=(qty, item.description, item.units, unit_price, total))

            self.qty_var.set("")
            self.desc_var.set("")
            self.units_var.set("")
            self.unit_price_var.set("")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter numeric values for Quantity and Unit Price.")

    def save_quote(self):
        try:
            customer = Customer(
                name=self.name_var.get(),
                address=self.address_var.get(),
                town=self.town_var.get(),
                contact=self.contact_var.get(),
                phone=self.phone_var.get(),
                email=self.email_var.get()
            )

            quote_data = QuoteMeta(
                quotation_no=self.quotation_no_var.get(),
                date=self.date_var.get(),
                order_no=self.order_no_var.get(),
                service=self.service_var.get()  # ✅ Service
            )

            labour_item = QuoteItem(
                qty=1,
                description=self.labour_desc_var.get(),
                units="Lumpsum",
                unit_price=float(self.labour_cost_var.get() or 0),
                total=float(self.labour_cost_var.get() or 0)
            )

            grand_total = float(self.total_var.get() or 0)

            full_quote = Quote(
                customer=customer,
                quote=quote_data,
                items=self.items,
                labour=labour_item,
                grand_total=grand_total
            )

            save_quote_to_excel(full_quote)
            messagebox.showinfo("Success", "Quote saved to Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save quote:\n{e}")


