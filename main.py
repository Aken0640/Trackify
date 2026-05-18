import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
import csv
import os
from tkcalendar import DateEntry
from datetime import datetime


class IntroPage:

    def __init__(self, root):

        self.root = root
        self.root.title("Trackify")
        self.root.geometry("600x650")
        self.root.configure(bg="black")

        self.frame = tk.Frame(root,bg="black")
        self.frame.pack(expand=True)

        self.title = tk.Label(
            self.frame,
            text="Trackify",
            font=("Helvetica",72,"bold"),
            fg="#00FFFF",
            bg="black"
        )
        self.title.pack(pady=(120,10))

        self.subtitle = tk.Label(
            self.frame,
            text="The Complete Expense Tracker",
            font=("Helvetica",18),
            fg="white",
            bg="black"
        )
        self.subtitle.pack(pady=(0,40))

        self.button = tk.Button(
            self.frame,
            text="Continue",
            font=("Helvetica",14,"bold"),
            bg="red",
            fg="white",
            command=self.start_app
        )
        self.button.pack()

        self.glow_state = 0
        self.glow_job = None
        self.animate_glow()


    def animate_glow(self):

        colors = ["#00FFFF","#66FFFF","#FFFFFF","#66FFFF"]

        if not self.title.winfo_exists():
            return

        self.title.config(fg=colors[self.glow_state])

        self.glow_state = (self.glow_state + 1) % len(colors)

        self.glow_job = self.root.after(400,self.animate_glow)


    def start_app(self):

        if self.glow_job:
            self.root.after_cancel(self.glow_job)

        self.frame.destroy()
        ExpenseTracker(self.root)



class ExpenseTracker:

    def __init__(self, root):

        self.root = root
        self.root.title("Trackify")
        self.root.geometry("600x650")
        self.root.configure(bg="black")

        self.balance = 0
        self.total_income = 0
        self.total_expense = 0
        self.expenses = {}

        self.income_categories = ["Allowance","Salary","Petty Cash","Bonus","Other"]
        self.expense_categories = ["Food","Transport","Shopping","Bills","Entertainment","Other"]

        self.file_name = "transactions.csv"

        if not os.path.exists(self.file_name):
            with open(self.file_name,"w",newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Date","Type","Category","Amount","Notes"])

        self.create_gui()
        self.load_transactions()


    def create_gui(self):

        title = tk.Label(self.root,
                         text="Trackify",
                         font=("Arial",20,"bold"),
                         fg="white",
                         bg="black")
        title.pack(pady=10)

        self.balance_label = tk.Label(self.root,text="Balance: ₹0",
                                      font=("Arial",14),
                                      fg="white",
                                      bg="black")
        self.balance_label.pack()

        self.income_label = tk.Label(self.root,text="Total Income: ₹0",
                                     font=("Arial",12),
                                     fg="cyan",
                                     bg="black")
        self.income_label.pack()

        self.expense_label = tk.Label(self.root,text="Total Expense: ₹0",
                                      font=("Arial",12),
                                      fg="red",
                                      bg="black")
        self.expense_label.pack()


        table_frame = tk.Frame(self.root,bg="black")
        table_frame.pack(pady=20)

        style = ttk.Style()
        style.theme_use("default")

        style.configure("Treeview",
                        background="#303030",
                        fieldbackground="#303030",
                        foreground="white")

        style.configure("Treeview.Heading",
                        background="black",
                        foreground="white")


        self.table = ttk.Treeview(table_frame,
                                  columns=("Date","Type","Category","Amount","Notes"),
                                  show="headings",
                                  height=15)

        self.table.heading("Date",text="Date")
        self.table.heading("Type",text="Type")
        self.table.heading("Category",text="Category")
        self.table.heading("Amount",text="Amount")
        self.table.heading("Notes",text="Notes")

        self.table.column("Date",width=90)
        self.table.column("Type",width=90)
        self.table.column("Category",width=160)
        self.table.column("Amount",width=90)
        self.table.column("Notes",width=150)

        self.table.tag_configure("expense",foreground="red")
        self.table.tag_configure("income",foreground="cyan")

        self.table.pack()


        chart_btn = tk.Button(self.root,
                              text="Show Expense Chart",
                              command=self.show_chart,
                              bg="black",
                              fg="white")

        chart_btn.pack(pady=10)


        self.add_btn = tk.Button(self.root,
                                 text="+",
                                 font=("Arial",20,"bold"),
                                 bg="red",
                                 fg="white",
                                 width=3,
                                 command=self.open_add_window)

        self.add_btn.place(relx=0.9,rely=0.9,anchor="center")


    def open_add_window(self):

        self.popup = tk.Toplevel(self.root)
        self.popup.title("Add Transaction")
        self.popup.geometry("300x350")
        self.popup.configure(bg="black")

        tk.Label(self.popup,text="Date",bg="black",fg="white").pack(pady=5)

        self.date_entry = DateEntry(
            self.popup,
            width=18,
            background="black",
            foreground="white",
            borderwidth=2,
            date_pattern="dd-mm-yyyy"
        )

        self.date_entry.pack()
        self.date_entry.set_date(datetime.now())


        tk.Label(self.popup,text="Amount",bg="black",fg="white").pack(pady=5)
        self.amount_entry = tk.Entry(self.popup)
        self.amount_entry.pack()


        tk.Label(self.popup,text="Type",bg="black",fg="white").pack(pady=5)

        self.type_box = ttk.Combobox(self.popup,
                                     values=["Income","Expense"],
                                     state="readonly")
        self.type_box.pack()
        self.type_box.current(1)
        self.type_box.bind("<<ComboboxSelected>>",self.update_categories)


        tk.Label(self.popup,text="Category",bg="black",fg="white").pack(pady=5)

        self.category_box = ttk.Combobox(self.popup,
                                         values=self.expense_categories,
                                         state="readonly")
        self.category_box.pack()
        self.category_box.current(0)

        tk.Label(self.popup,text="Notes",bg="black",fg="white").pack(pady=5)
        self.notes_entry = tk.Entry(self.popup)
        self.notes_entry.pack()


        add_btn = tk.Button(self.popup,
                            text="Add",
                            command=self.add_transaction,
                            bg="red",
                            fg="white")

        add_btn.pack(pady=15)


    def update_categories(self,event):

        if self.type_box.get() == "Income":
            self.category_box["values"] = self.income_categories
        else:
            self.category_box["values"] = self.expense_categories

        self.category_box.current(0)


    def add_transaction(self):

        date = self.date_entry.get()
        notes = self.notes_entry.get()

        try:
            amount = float(self.amount_entry.get())
        except:
            messagebox.showerror("Error","Enter valid amount")
            return

        t_type = self.type_box.get()
        category = self.category_box.get()

        if t_type == "Income":

            self.balance += amount
            self.total_income += amount
            tag = "income"

        else:

            self.balance -= amount
            self.total_expense += amount
            self.expenses[category] = self.expenses.get(category,0) + amount
            tag = "expense"


        self.balance_label.config(text="Balance: ₹"+str(self.balance))
        self.income_label.config(text="Total Income: ₹"+str(self.total_income))
        self.expense_label.config(text="Total Expense: ₹"+str(self.total_expense))


        self.table.insert("",tk.END,values=(date,t_type,category,amount,notes),tags=(tag,))


        with open(self.file_name,"a",newline="") as f:
            writer = csv.writer(f)
            writer.writerow([date,t_type,category,amount,notes])

        self.popup.destroy()


    def load_transactions(self):

        with open(self.file_name,"r") as f:

            reader = csv.DictReader(f)

            for row in reader:

                date = row["Date"]
                t_type = row["Type"]
                category = row["Category"]
                amount = float(row["Amount"])
                notes = row.get("Notes","")

                if t_type == "Income":
                    self.balance += amount
                    self.total_income += amount
                    tag="income"
                else:
                    self.balance -= amount
                    self.total_expense += amount
                    self.expenses[category] = self.expenses.get(category,0) + amount
                    tag="expense"

                self.table.insert("",tk.END,values=(date,t_type,category,amount,notes),tags=(tag,))

        self.balance_label.config(text="Balance: ₹"+str(self.balance))
        self.income_label.config(text="Total Income: ₹"+str(self.total_income))
        self.expense_label.config(text="Total Expense: ₹"+str(self.total_expense))


    def show_chart(self):

        if len(self.expenses) == 0:
            return

        labels = list(self.expenses.keys())
        values = list(self.expenses.values())

        explode = [0]*len(values)

        fig, ax = plt.subplots()

        fig.patch.set_facecolor("black")
        ax.set_facecolor("black")

        colors = ["#ff4d4d","#ff944d","#4da6ff","#cc66ff","#00cccc","#ff66b2"]

        def draw_chart():
            ax.clear()
            ax.set_facecolor("black")

            wedges, texts, autotexts = ax.pie(
                values,
                labels=labels,
                autopct="%1.1f%%",
                explode=explode,
                startangle=90,
                colors=colors,
                labeldistance=1.15,
                wedgeprops={'edgecolor':'black'}
            )

            ax.set_title("Expense Distribution",color="white")

            for t in texts:
                t.set_color("white")
                t.set_fontsize(10)
                t.set_weight("bold")

            for t in autotexts:
                t.set_color("white")

            fig.canvas.draw_idle()

            return wedges

        wedges = draw_chart()


        def on_hover(event):

            found=False

            for i,wedge in enumerate(wedges):

                if wedge.contains_point([event.x,event.y]):

                    explode[:] = [0]*len(values)
                    explode[i] = 0.18
                    draw_chart()

                    found=True
                    break

            if not found and max(explode)!=0:

                explode[:] = [0]*len(values)
                draw_chart()

        fig.canvas.mpl_connect("motion_notify_event",on_hover)

        plt.show()


root = tk.Tk()

IntroPage(root)

root.mainloop()