import tkinter as tk
from tkinter import messagebox
from unittest import result
from docxtpl import DocxTemplate
import psycopg2
import json
import os
from openpyxl import Workbook
# ================== TẠO HỒ SƠ ==================
def create_docs():
    try:
        conn = psycopg2.connect(
            host="db.cppilyhbusukcmrwpvfc.supabase.co",
            database="postgres",
            user="postgres",
            password="Reymysterio109"
        )
        cur = conn.cursor()

        # 🔥 Lấy toàn bộ người VANGNHA = TRUE
        cur.execute('SELECT * FROM public.population WHERE "VANGNHA" = TRUE;')
        rows = cur.fetchall()

        if not rows:
            messagebox.showinfo("Thông báo", "Không có người vắng nhà!")
            return

        # 🔥 Tạo folder output nếu chưa có
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)

        # 🔥 Lấy tên cột
        columns = [desc[0] for desc in cur.description]

        # 🔥 Gom nhóm theo SOHOK
        result = {}

        for row in rows:
            data = dict(zip(columns, row))
            sohkh = data.get("SOHOK")



            def format_quanhe(qh):
                if not qh:
                    return ""
                qh = qh.strip()

                if qh.upper() == "CH":
                    return "Chủ hộ"

                return qh.capitalize()  # viết hoa chữ cái đầu

            if not sohkh:
                continue

            person = {
                "HOTEN": data.get("HOTEN"),
                "NAMSINH": data.get("NAMSINH"),
                "GIOITINH": data.get("GIOITINH"),
                "CCCD": data.get("CCCD"),
                "QUANHE": format_quanhe(data.get("QUANHE")),
                "NOITHTRU": data.get("NOITHTRU"),
                "SDT": data.get("SDT"),
            }

            if sohkh not in result:
                result[sohkh] = []

            result[sohkh].append(person)

        # ================== LOG ==================
        print("\n===== KẾT QUẢ GROUP ALL VẮNG NHÀ =====")
        print(json.dumps(result, indent=4, ensure_ascii=False))

        # ================== TẠO FILE ==================
        for sohkh, people in result.items():

                    # ================== TẠO FILE EXCEL ==================
            wb = Workbook()
            ws = wb.active
            ws.title = "ThongKeVangNha"

            # Header
            ws.append(["SOHOK", "CCCD", "HOTEN", "NAMSINH", "QUANHE"])

            for sohkh, people in result.items():
                for person in people:
                    ws.append([
                        sohkh,
                        person.get("CCCD"),
                        person.get("HOTEN"),
                        person.get("NAMSINH"),
                        person.get("QUANHE"),
                    ])

            excel_path = os.path.join(output_dir, "thong_ke_vang_nha.xlsx")
            wb.save(excel_path)

            print(f"👉 Đang tạo file cho SOHOK: {sohkh}")

            doc = DocxTemplate("BIEN BAN XAC MINH XOA KHAU.docx")

            context = {
                "SOHOK": sohkh,
                "people": people,
                "NOITHTRU": people[0]["NOITHTRU"]   # 👈 lấy đại diện
            }

            safe_name = str(sohkh).replace("/", "_").replace("\\", "_")

            file_path = os.path.join(output_dir, f"{safe_name}.docx")

            doc.render(context)
            doc.save(file_path)

        messagebox.showinfo(
            "Thành công",
            f"Đã tạo {len(result)} file Word và 1 file Excel thống kê trong thư mục output"
        )
    except Exception as e:
        print("ERROR:", e)
        messagebox.showerror("Lỗi", str(e))


# ================== UI ==================
root = tk.Tk()
root.title("Tạo hồ sơ vắng nhà")

tk.Button(root, text="Tạo hồ sơ tất cả người vắng", command=create_docs).pack(pady=20)

root.mainloop()