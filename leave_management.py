import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
import os

EXCEL_FILE = 'leave_management.xlsx'


def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            summary_cols = ['ID', 'Employee', 'CL_Total', 'CL_Rem', 'SL_Total', 'SL_Rem', 'PL_Total', 'PL_Rem']
            pd.DataFrame(columns=summary_cols).to_excel(writer, sheet_name='Summary', index=False)
            log_cols = ['ID', 'Employee', 'Date', 'Month', 'Leave_Type', 'Reason']
            pd.DataFrame(columns=log_cols).to_excel(writer, sheet_name='Logs', index=False)


initialize_excel()


class LeaveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Leave Management System")
        self.root.geometry("1100x700")

        # --- SECTION 1: REGISTRATION ---
        frame_reg = tk.LabelFrame(root, text="Register / Update Quota", padx=10, pady=10)
        frame_reg.pack(fill="x", padx=10, pady=5)

        labels = ["ID:", "Name:", "CL Quota:", "SL Quota:", "PL Quota:"]
        self.reg_entries = {}
        for i, text in enumerate(labels):
            tk.Label(frame_reg, text=text).grid(row=0, column=i * 2)
            ent = tk.Entry(frame_reg, width=10)
            ent.grid(row=0, column=i * 2 + 1, padx=5)
            self.reg_entries[text] = ent

        tk.Button(frame_reg, text="Save Employee", command=self.add_or_update, bg="#2196F3", fg="white").grid(row=0,
                                                                                                              column=10,
                                                                                                              padx=10)

        # --- SECTION 2: APPLY LEAVE ---
        frame_apply = tk.LabelFrame(root, text="Apply for Leave", padx=10, pady=10)
        frame_apply.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_apply, text="Employee:").grid(row=0, column=0)
        self.emp_combo = ttk.Combobox(frame_apply, postcommand=self.refresh_list, width=30)
        self.emp_combo.grid(row=0, column=1, padx=5)
        # Bind the selection event to update history
        self.emp_combo.bind("<<ComboboxSelected>>", self.on_employee_select)

        tk.Label(frame_apply, text="Type:").grid(row=0, column=2)
        self.type_combo = ttk.Combobox(frame_apply, values=["Casual Leave", "Sick Leave", "Privilege Leave"], width=15)
        self.type_combo.grid(row=0, column=3, padx=5)

        tk.Label(frame_apply, text="Date:").grid(row=1, column=0, pady=10)
        self.cal = DateEntry(frame_apply, width=12)
        self.cal.grid(row=1, column=1, sticky="w", padx=5)

        tk.Label(frame_apply, text="Reason:").grid(row=1, column=2)
        self.reason_ent = tk.Entry(frame_apply, width=40)
        self.reason_ent.grid(row=1, column=3, columnspan=3, padx=5)

        tk.Button(frame_apply, text="Submit Application", command=self.apply_leave, bg="#4CAF50", fg="white").grid(
            row=1, column=6, padx=10)

        # --- SECTION 3: VIEW LOGS ---
        self.label_tree = tk.Label(root, text="Employee Leave History", font=("Arial", 10, "bold"))
        self.label_tree.pack(pady=5)
        self.tree = ttk.Treeview(root, columns=("ID", "Name", "Date", "Type", "Reason"), show='headings')
        for col in ("ID", "Name", "Date", "Type", "Reason"):
            self.tree.heading(col, text=col)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def refresh_list(self):
        df = pd.read_excel(EXCEL_FILE, sheet_name='Summary', dtype={'ID': str})
        self.emp_combo['values'] = [f"{row['ID']} | {row['Employee']}" for _, row in df.iterrows()]

    def on_employee_select(self, event):
        selected_val = self.emp_combo.get()
        if selected_val:
            eid = selected_val.split(" | ")[0]
            self.update_treeview(emp_id=eid)

    def add_or_update(self):
        data = {k: v.get().strip() for k, v in self.reg_entries.items()}
        if not all(data.values()):
            messagebox.showerror("Error", "All fields required")
            return

        df = pd.read_excel(EXCEL_FILE, sheet_name='Summary', dtype={'ID': str})
        eid = data["ID:"]

        new_row_data = {
            'ID': eid, 'Employee': data["Name:"],
            'CL_Total': int(data["CL Quota:"]), 'CL_Rem': int(data["CL Quota:"]),
            'SL_Total': int(data["SL Quota:"]), 'SL_Rem': int(data["SL Quota:"]),
            'PL_Total': int(data["PL Quota:"]), 'PL_Rem': int(data["PL Quota:"])
        }

        if eid in df['ID'].values:
            df.update(pd.DataFrame([new_row_data], index=df[df['ID'] == eid].index))
        else:
            df = pd.concat([df, pd.DataFrame([new_row_data])], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Summary', index=False)
        messagebox.showinfo("Success", "Employee Record Saved")

    def apply_leave(self):
        emp_val = self.emp_combo.get()
        if not (emp_val and self.type_combo.get()):
            messagebox.showerror("Error", "Selection incomplete")
            return

        eid = emp_val.split(" | ")[0]
        l_type = self.type_combo.get()
        reason = self.reason_ent.get()
        date_sel = self.cal.get_date()

        mapping = {"Casual Leave": "CL", "Sick Leave": "SL", "Privilege Leave": "PL"}
        prefix = mapping[l_type]

        summary_df = pd.read_excel(EXCEL_FILE, sheet_name='Summary', dtype={'ID': str})
        idx = summary_df.index[summary_df['ID'] == eid].tolist()[0]

        if summary_df.at[idx, f'{prefix}_Rem'] >= 1:
            summary_df.at[idx, f'{prefix}_Rem'] -= 1
            log_df = pd.read_excel(EXCEL_FILE, sheet_name='Logs', dtype={'ID': str})
            new_log = {
                'ID': eid, 'Employee': summary_df.at[idx, 'Employee'],
                'Date': date_sel.strftime('%Y-%m-%d'), 'Month': date_sel.strftime('%b'),
                'Leave_Type': l_type, 'Reason': reason
            }
            log_df = pd.concat([log_df, pd.DataFrame([new_log])], ignore_index=True)

            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                log_df.to_excel(writer, sheet_name='Logs', index=False)

            messagebox.showinfo("Approved", "Leave Logged Successfully")
            self.update_treeview(emp_id=eid)  # Refresh for the specific employee
        else:
            messagebox.showwarning("Denied", f"No {l_type} balance left!")

    def update_treeview(self, emp_id=None):
        for i in self.tree.get_children(): self.tree.delete(i)
        log_df = pd.read_excel(EXCEL_FILE, sheet_name='Logs', dtype={'ID': str})

        # Filter logic
        if emp_id:
            log_df = log_df[log_df['ID'] == emp_id]

        for _, row in log_df.tail(10).iterrows():
            self.tree.insert("", 0, values=(row['ID'], row['Employee'], row['Date'], row['Leave_Type'], row['Reason']))


if __name__ == "__main__":
    root = tk.Tk()
    app = LeaveApp(root)
    root.mainloop()