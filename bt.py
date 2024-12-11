from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import csv
from datetime import datetime
import pandas as pd


def save_to_csv():
    employee_data = {
        "Mã nhân viên": entry_employee_id.get(),
        "Tên nhân viên": entry_employee_name.get(),
        "Ngày sinh": date_of_birth_entry.get(),
        "Giới tính": "Nam" if gender_var.get() == 1 else "Nữ",
        "Đơn vị": department_combobox.get(),
        "Số CMND": id_card_entry.get(),
        "Ngày cấp CMND": id_issue_date_entry.get(),
        "Nơi cấp CMND": id_issue_place_entry.get(),
        "Chức danh": position_entry.get()
    }


    with open("employees.csv", mode="a", newline='', encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=employee_data.keys())
        if file.tell() == 0:  # Kiểm tra nếu file rỗng, thì ghi header
            writer.writeheader()
        writer.writerow(employee_data)

    messagebox.showinfo("Thông báo", "Lưu thông tin thành công!")
    clear_input_fields()


def check_birthday_today():
    try:
        today = datetime.now().strftime("%d/%m/%Y")
        birthday_employees = []
        with open("employees.csv", mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Ngày sinh'][:-5] == today[:-5]:  # So sánh ngày và tháng
                    birthday_employees.append(row)

        if birthday_employees:
            result = "Nhân viên có sinh nhật hôm nay:\n\n" + "\n".join([row['Tên nhân viên'] for row in birthday_employees])
        else:
            result = "Không có nhân viên nào sinh nhật hôm nay."

        messagebox.showinfo("Kết quả", result)
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "File dữ liệu chưa được tạo!")


def export_to_excel():
    try:
        df = pd.read_csv("employees.csv", encoding="utf-8")
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format="%d/%m/%Y")
        df.sort_values(by="Ngày sinh", ascending=True, inplace=True)
        output_file = "sorted_employees.xlsx"
        df.to_excel(output_file, index=False)  # Xuất ra file Excel
        messagebox.showinfo("Thông báo", f"Xuất danh sách thành công! File: {output_file}")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "File dữ liệu chưa được tạo!")


def clear_input_fields():
    entry_employee_id.delete(0, END)
    entry_employee_name.delete(0, END)
    date_of_birth_entry.set_date(datetime.now())
    gender_var.set(0)
    department_combobox.set("")
    id_card_entry.delete(0, END)
    id_issue_date_entry.set_date(datetime.now())
    id_issue_place_entry.delete(0, END)
    position_entry.delete(0, END)


window = Tk()
window.title("Thông tin nhân viên")
window.geometry("850x400")

# Tiêu đề
lbl = Label(window, text="Thông tin nhân viên", fg="black", font=("Timenewroman", 23))
lbl.grid(column=0, row=0, columnspan=4, pady=10, sticky="W")


chk_customer = Checkbutton(window, text="Là khách hàng")
chk_customer.grid(column=1, row=0, sticky="w")

chk_employee = Checkbutton(window, text="Là nhân viên")
chk_employee.grid(column=2, row=0)


label_employee_id = Label(window, text="Mã nhân viên", fg="black", font=("Arial", 10))
label_employee_id.grid(column=0, row=1, sticky="W")
entry_employee_id = Entry(window, width=30)
entry_employee_id.grid(column=0, row=2, padx=5, pady=5, sticky="W")


label_employee_name = Label(window, text="Tên nhân viên", fg="black", font=("Arial", 10))
label_employee_name.grid(column=1, row=1, sticky="W")
entry_employee_name = Entry(window, width=30, bd=2, relief="groove")
entry_employee_name.grid(column=1, row=2, padx=5, pady=5, sticky="w")


label_date_of_birth = Label(window, text="Ngày sinh", fg="black", font=("Arial", 10))
label_date_of_birth.grid(column=2, row=1, sticky="W")
date_of_birth_entry = DateEntry(window, width=20, foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
date_of_birth_entry.grid(column=2, row=2, sticky="W")


label_gender = Label(window, text="Giới tính", fg="black", font=("Arial", 10))
label_gender.grid(column=3, row=1, sticky="W")
gender_var = IntVar()
radio_male = Radiobutton(window, text="Nam", variable=gender_var, value=1)
radio_male.grid(row=2, column=3, padx=10, pady=5, sticky="W")
radio_female = Radiobutton(window, text="Nữ", variable=gender_var, value=2)
radio_female.grid(row=2, column=4, padx=10, pady=5, sticky="W")


label_department = Label(window, text="Đơn vị", fg="black", font=("Arial", 10))
label_department.grid(column=0, row=3, sticky="W")
department_combobox = StringVar()
departments = ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5", "Lớp 6"]
department_combobox = ttk.Combobox(window, textvariable=department_combobox, values=departments, width=27, font=("Arial", 12), state="readonly")
department_combobox.grid(row=4, column=0, padx=5, pady=5, sticky="W")


label_id_card = Label(window, text="Số CMND", fg="black", font=("Arial", 10))
label_id_card.grid(column=1, row=3, sticky="W")
id_card_entry = Entry(window, width=30)
id_card_entry.grid(column=1, row=4, sticky="W")


label_id_issue_date = Label(window, text="Ngày cấp CMND", fg="black", font=("Arial", 10))
label_id_issue_date.grid(column=2, row=3, sticky="W")
id_issue_date_entry = DateEntry(window, width=20, foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
id_issue_date_entry.grid(column=2, row=4, sticky="W")


label_position = Label(window, text="Chức danh", fg="black", font=("Arial", 10))
label_position.grid(column=0, row=5, sticky="W")
position_entry = Entry(window, width=40)
position_entry.grid(column=0, row=6, sticky="W")


label_id_issue_place = Label(window, text="Nơi cấp CMND", fg="black", font=("Arial", 10))
label_id_issue_place.grid(column=1, row=5, sticky="W")
id_issue_place_entry = Entry(window, width=40)
id_issue_place_entry.grid(column=1, row=6, sticky="W")


btn_save = Button(window, text="Lưu thông tin", command=save_to_csv, width=15, height=2)
btn_save.grid(row=7, column=0, padx=10, pady=20)

btn_check_birthday = Button(window, text="Sinh nhật hôm nay", command=check_birthday_today, width=20, height=2)
btn_check_birthday.grid(row=7, column=1, padx=10, pady=20)

btn_export = Button(window, text="Xuất danh sách", command=export_to_excel, width=20, height=2)
btn_export.grid(row=7, column=2, padx=10, pady=20)

# Chạy giao diện
window.mainloop()
