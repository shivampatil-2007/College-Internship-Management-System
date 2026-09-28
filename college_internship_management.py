import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
import os


FILE_NAME = "internship_management.xlsx"


# -------------------- EXCEL SETUP --------------------

def create_excel_file():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()

        ws = wb.active
        ws.title = "Students"
        ws.append(["Name", "Roll No", "Branch", "Year", "CGPA", "Skills"])

        ws = wb.create_sheet("Companies")
        ws.append(["Company Name", "Location", "Role", "Stipend"])

        ws = wb.create_sheet("Internships")
        ws.append(["Company Name", "Role", "Duration", "Stipend"])

        ws = wb.create_sheet("Applications")
        ws.append(["Student Name", "Roll No", "Company", "Role", "Status"])

        wb.save(FILE_NAME)


create_excel_file()


# -------------------- MAIN WINDOW --------------------

root = tk.Tk()
root.title("College Internship Management System")
root.geometry("1100x650")
root.resizable(False, False)


# -------------------- FUNCTIONS --------------------

def add_student():
    name = student_name.get()
    roll = student_roll.get()
    branch = student_branch.get()
    year = student_year.get()
    cgpa = student_cgpa.get()
    skills = student_skills.get()

    if name == "" or roll == "" or branch == "":
        messagebox.showwarning("Warning", "Please enter required details.")
        return

    wb = load_workbook(FILE_NAME)
    ws = wb["Students"]

    ws.append([name, roll, branch, year, cgpa, skills])

    wb.save(FILE_NAME)

    messagebox.showinfo("Success", "Student added successfully!")

    clear_student_fields()
    show_students()


def clear_student_fields():
    student_name.delete(0, tk.END)
    student_roll.delete(0, tk.END)
    student_branch.delete(0, tk.END)
    student_year.delete(0, tk.END)
    student_cgpa.delete(0, tk.END)
    student_skills.delete(0, tk.END)


def show_students():
    for item in student_table.get_children():
        student_table.delete(item)

    wb = load_workbook(FILE_NAME)
    ws = wb["Students"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        student_table.insert("", tk.END, values=row)


def add_company():
    company = company_name.get()
    location = company_location.get()
    role = company_role.get()
    stipend = company_stipend.get()

    if company == "" or role == "":
        messagebox.showwarning("Warning", "Please enter company and role.")
        return

    wb = load_workbook(FILE_NAME)
    ws = wb["Companies"]

    ws.append([company, location, role, stipend])

    wb.save(FILE_NAME)

    messagebox.showinfo("Success", "Company added successfully!")

    clear_company_fields()
    show_companies()


def clear_company_fields():
    company_name.delete(0, tk.END)
    company_location.delete(0, tk.END)
    company_role.delete(0, tk.END)
    company_stipend.delete(0, tk.END)


def show_companies():
    for item in company_table.get_children():
        company_table.delete(item)

    wb = load_workbook(FILE_NAME)
    ws = wb["Companies"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        company_table.insert("", tk.END, values=row)


def add_internship():
    company = internship_company.get()
    role = internship_role.get()
    duration = internship_duration.get()
    stipend = internship_stipend.get()

    if company == "" or role == "":
        messagebox.showwarning("Warning", "Please enter company and role.")
        return

    wb = load_workbook(FILE_NAME)
    ws = wb["Internships"]

    ws.append([company, role, duration, stipend])

    wb.save(FILE_NAME)

    messagebox.showinfo("Success", "Internship added successfully!")

    clear_internship_fields()
    show_internships()


def clear_internship_fields():
    internship_company.delete(0, tk.END)
    internship_role.delete(0, tk.END)
    internship_duration.delete(0, tk.END)
    internship_stipend.delete(0, tk.END)


def show_internships():
    for item in internship_table.get_children():
        internship_table.delete(item)

    wb = load_workbook(FILE_NAME)
    ws = wb["Internships"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        internship_table.insert("", tk.END, values=row)


def add_application():
    name = application_student.get()
    roll = application_roll.get()
    company = application_company.get()
    role = application_role.get()
    status = application_status.get()

    if name == "" or roll == "" or company == "":
        messagebox.showwarning("Warning", "Please enter required details.")
        return

    wb = load_workbook(FILE_NAME)
    ws = wb["Applications"]

    ws.append([name, roll, company, role, status])

    wb.save(FILE_NAME)

    messagebox.showinfo("Success", "Application added successfully!")

    clear_application_fields()
    show_applications()


def clear_application_fields():
    application_student.delete(0, tk.END)
    application_roll.delete(0, tk.END)
    application_company.delete(0, tk.END)
    application_role.delete(0, tk.END)
    application_status.set("Pending")


def show_applications():
    for item in application_table.get_children():
        application_table.delete(item)

    wb = load_workbook(FILE_NAME)
    ws = wb["Applications"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        application_table.insert("", tk.END, values=row)


def update_dashboard():
    wb = load_workbook(FILE_NAME)

    students = wb["Students"].max_row - 1
    companies = wb["Companies"].max_row - 1
    internships = wb["Internships"].max_row - 1
    applications = wb["Applications"].max_row - 1

    student_count.config(text=str(max(students, 0)))
    company_count.config(text=str(max(companies, 0)))
    internship_count.config(text=str(max(internships, 0)))
    application_count.config(text=str(max(applications, 0)))


# -------------------- TITLE --------------------

title = tk.Label(
    root,
    text="COLLEGE INTERNSHIP MANAGEMENT SYSTEM",
    font=("Arial", 22, "bold")
)
title.pack(pady=15)


# -------------------- NOTEBOOK --------------------

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=15, pady=10)


# ==========================================================
# DASHBOARD
# ==========================================================

dashboard = ttk.Frame(notebook)
notebook.add(dashboard, text="Dashboard")


tk.Label(
    dashboard,
    text="Internship Management Dashboard",
    font=("Arial", 20, "bold")
).pack(pady=25)


dashboard_frame = tk.Frame(dashboard)
dashboard_frame.pack(pady=20)


# Student card
student_card = tk.LabelFrame(
    dashboard_frame,
    text="Students",
    width=220,
    height=130
)
student_card.grid(row=0, column=0, padx=20, pady=20)
student_card.pack_propagate(False)

student_count = tk.Label(
    student_card,
    text="0",
    font=("Arial", 30, "bold")
)
student_count.pack(pady=25)


# Company card
company_card = tk.LabelFrame(
    dashboard_frame,
    text="Companies",
    width=220,
    height=130
)
company_card.grid(row=0, column=1, padx=20, pady=20)
company_card.pack_propagate(False)

company_count = tk.Label(
    company_card,
    text="0",
    font=("Arial", 30, "bold")
)
company_count.pack(pady=25)


# Internship card
internship_card = tk.LabelFrame(
    dashboard_frame,
    text="Internships",
    width=220,
    height=130
)
internship_card.grid(row=1, column=0, padx=20, pady=20)
internship_card.pack_propagate(False)

internship_count = tk.Label(
    internship_card,
    text="0",
    font=("Arial", 30, "bold")
)
internship_count.pack(pady=25)


# Application card
application_card = tk.LabelFrame(
    dashboard_frame,
    text="Applications",
    width=220,
    height=130
)
application_card.grid(row=1, column=1, padx=20, pady=20)
application_card.pack_propagate(False)

application_count = tk.Label(
    application_card,
    text="0",
    font=("Arial", 30, "bold")
)
application_count.pack(pady=25)


tk.Button(
    dashboard,
    text="Refresh Dashboard",
    command=update_dashboard,
    width=20
).pack(pady=20)


# ==========================================================
# STUDENT MANAGEMENT
# ==========================================================

student_tab = ttk.Frame(notebook)
notebook.add(student_tab, text="Students")


student_form = tk.LabelFrame(
    student_tab,
    text="Add Student",
    padx=15,
    pady=15
)
student_form.pack(fill="x", padx=15, pady=10)


tk.Label(student_form, text="Name").grid(row=0, column=0, padx=5, pady=5)
student_name = tk.Entry(student_form, width=20)
student_name.grid(row=0, column=1, padx=5, pady=5)

tk.Label(student_form, text="Roll No").grid(row=0, column=2, padx=5, pady=5)
student_roll = tk.Entry(student_form, width=20)
student_roll.grid(row=0, column=3, padx=5, pady=5)

tk.Label(student_form, text="Branch").grid(row=1, column=0, padx=5, pady=5)
student_branch = tk.Entry(student_form, width=20)
student_branch.grid(row=1, column=1, padx=5, pady=5)

tk.Label(student_form, text="Year").grid(row=1, column=2, padx=5, pady=5)
student_year = tk.Entry(student_form, width=20)
student_year.grid(row=1, column=3, padx=5, pady=5)

tk.Label(student_form, text="CGPA").grid(row=2, column=0, padx=5, pady=5)
student_cgpa = tk.Entry(student_form, width=20)
student_cgpa.grid(row=2, column=1, padx=5, pady=5)

tk.Label(student_form, text="Skills").grid(row=2, column=2, padx=5, pady=5)
student_skills = tk.Entry(student_form, width=20)
student_skills.grid(row=2, column=3, padx=5, pady=5)


tk.Button(
    student_form,
    text="Add Student",
    command=add_student,
    width=18
).grid(row=3, column=1, pady=10)


tk.Button(
    student_form,
    text="Clear",
    command=clear_student_fields,
    width=18
).grid(row=3, column=2, pady=10)


student_table = ttk.Treeview(
    student_tab,
    columns=("Name", "Roll", "Branch", "Year", "CGPA", "Skills"),
    show="headings"
)

for col in ("Name", "Roll", "Branch", "Year", "CGPA", "Skills"):
    student_table.heading(col, text=col)
    student_table.column(col, width=130)

student_table.pack(fill="both", expand=True, padx=15, pady=10)

show_students()


# ==========================================================
# COMPANY MANAGEMENT
# ==========================================================

company_tab = ttk.Frame(notebook)
notebook.add(company_tab, text="Companies")


company_form = tk.LabelFrame(
    company_tab,
    text="Add Company",
    padx=15,
    pady=15
)
company_form.pack(fill="x", padx=15, pady=10)


tk.Label(company_form, text="Company Name").grid(
    row=0, column=0, padx=5, pady=5
)
company_name = tk.Entry(company_form, width=25)
company_name.grid(row=0, column=1, padx=5, pady=5)

tk.Label(company_form, text="Location").grid(
    row=0, column=2, padx=5, pady=5
)
company_location = tk.Entry(company_form, width=25)
company_location.grid(row=0, column=3, padx=5, pady=5)

tk.Label(company_form, text="Role").grid(
    row=1, column=0, padx=5, pady=5
)
company_role = tk.Entry(company_form, width=25)
company_role.grid(row=1, column=1, padx=5, pady=5)

tk.Label(company_form, text="Stipend").grid(
    row=1, column=2, padx=5, pady=5
)
company_stipend = tk.Entry(company_form, width=25)
company_stipend.grid(row=1, column=3, padx=5, pady=5)


tk.Button(
    company_form,
    text="Add Company",
    command=add_company,
    width=18
).grid(row=2, column=1, pady=10)


tk.Button(
    company_form,
    text="Clear",
    command=clear_company_fields,
    width=18
).grid(row=2, column=2, pady=10)


company_table = ttk.Treeview(
    company_tab,
    columns=("Company", "Location", "Role", "Stipend"),
    show="headings"
)

for col in ("Company", "Location", "Role", "Stipend"):
    company_table.heading(col, text=col)
    company_table.column(col, width=180)

company_table.pack(fill="both", expand=True, padx=15, pady=10)

show_companies()


# ==========================================================
# INTERNSHIP MANAGEMENT
# ==========================================================

internship_tab = ttk.Frame(notebook)
notebook.add(internship_tab, text="Internships")


internship_form = tk.LabelFrame(
    internship_tab,
    text="Add Internship",
    padx=15,
    pady=15
)
internship_form.pack(fill="x", padx=15, pady=10)


tk.Label(internship_form, text="Company").grid(
    row=0, column=0, padx=5, pady=5
)
internship_company = tk.Entry(internship_form, width=25)
internship_company.grid(row=0, column=1, padx=5, pady=5)

tk.Label(internship_form, text="Role").grid(
    row=0, column=2, padx=5, pady=5
)
internship_role = tk.Entry(internship_form, width=25)
internship_role.grid(row=0, column=3, padx=5, pady=5)

tk.Label(internship_form, text="Duration").grid(
    row=1, column=0, padx=5, pady=5
)
internship_duration = tk.Entry(internship_form, width=25)
internship_duration.grid(row=1, column=1, padx=5, pady=5)

tk.Label(internship_form, text="Stipend").grid(
    row=1, column=2, padx=5, pady=5
)
internship_stipend = tk.Entry(internship_form, width=25)
internship_stipend.grid(row=1, column=3, padx=5, pady=5)


tk.Button(
    internship_form,
    text="Add Internship",
    command=add_internship,
    width=18
).grid(row=2, column=1, pady=10)


tk.Button(
    internship_form,
    text="Clear",
    command=clear_internship_fields,
    width=18
).grid(row=2, column=2, pady=10)


internship_table = ttk.Treeview(
    internship_tab,
    columns=("Company", "Role", "Duration", "Stipend"),
    show="headings"
)

for col in ("Company", "Role", "Duration", "Stipend"):
    internship_table.heading(col, text=col)
    internship_table.column(col, width=200)

internship_table.pack(fill="both", expand=True, padx=15, pady=10)

show_internships()


# ==========================================================
# APPLICATION MANAGEMENT
# ==========================================================

application_tab = ttk.Frame(notebook)
notebook.add(application_tab, text="Applications")


application_form = tk.LabelFrame(
    application_tab,
    text="Add Application",
    padx=15,
    pady=15
)
application_form.pack(fill="x", padx=15, pady=10)


tk.Label(application_form, text="Student Name").grid(
    row=0, column=0, padx=5, pady=5
)
application_student = tk.Entry(application_form, width=22)
application_student.grid(row=0, column=1, padx=5, pady=5)

tk.Label(application_form, text="Roll No").grid(
    row=0, column=2, padx=5, pady=5
)
application_roll = tk.Entry(application_form, width=22)
application_roll.grid(row=0, column=3, padx=5, pady=5)

tk.Label(application_form, text="Company").grid(
    row=1, column=0, padx=5, pady=5
)
application_company = tk.Entry(application_form, width=22)
application_company.grid(row=1, column=1, padx=5, pady=5)

tk.Label(application_form, text="Role").grid(
    row=1, column=2, padx=5, pady=5
)
application_role = tk.Entry(application_form, width=22)
application_role.grid(row=1, column=3, padx=5, pady=5)

tk.Label(application_form, text="Status").grid(
    row=2, column=0, padx=5, pady=5
)

application_status = ttk.Combobox(
    application_form,
    values=["Pending", "Selected", "Rejected"],
    state="readonly",
    width=19
)
application_status.set("Pending")
application_status.grid(row=2, column=1, padx=5, pady=5)


tk.Button(
    application_form,
    text="Add Application",
    command=add_application,
    width=18
).grid(row=3, column=1, pady=10)


tk.Button(
    application_form,
    text="Clear",
    command=clear_application_fields,
    width=18
).grid(row=3, column=2, pady=10)


application_table = ttk.Treeview(
    application_tab,
    columns=("Student", "Roll", "Company", "Role", "Status"),
    show="headings"
)

for col in ("Student", "Roll", "Company", "Role", "Status"):
    application_table.heading(col, text=col)
    application_table.column(col, width=160)

application_table.pack(fill="both", expand=True, padx=15, pady=10)

show_applications()


# -------------------- INITIAL DASHBOARD --------------------

update_dashboard()


# -------------------- RUN APPLICATION --------------------

root.mainloop()
