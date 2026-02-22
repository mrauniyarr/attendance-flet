import flet as ft
from datetime import datetime
import openpyxl
import os
import shutil


# --- SMART STORAGE LOGIC ---
def get_db_path():
    path = os.getenv("FLET_APP_STORAGE_DATA")
    if not path:
        path = os.getcwd()
    return os.path.join(path, "Master_Attendance.xlsx")


def main(page: ft.Page):
    page.title = "Attendance System (Universal)"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 380
    page.window_height = 800
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.ADAPTIVE

    EXCEL_FILE = get_db_path()

    # --- UI CONTROLS ---
    roll_display = ft.TextField(
        label="Roll Number",
        read_only=True,
        text_align=ft.TextAlign.CENTER,
        width=250,
        text_size=30,
        color="blue",
        border_radius=15,
        bgcolor=ft.Colors.BLUE_50,
    )
    present_count_text = ft.Text(
        "Total Present: 0", size=20, weight="bold", color="green700"
    )

    selected_class = ft.Dropdown(
        label="Select Class",
        width=300,
        value="UG_1",
        options=[
            ft.dropdown.Option("UG_1"),
            ft.dropdown.Option("UG_2"),
            ft.dropdown.Option("UG_3"),
            ft.dropdown.Option("PG_1"),
            ft.dropdown.Option("PG_2"),
        ],
    )

    date_display = ft.TextField(
        label="Selected Date",
        value=datetime.now().strftime("%d-%m-%Y"),
        read_only=True,
        width=220,
    )

    # --- CORE LOGIC ---
    def update_present_count(e=None):
        count = 0
        if os.path.exists(EXCEL_FILE):
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
                if selected_class.value in wb.sheetnames:
                    ws = wb[selected_class.value]
                    current_date = date_display.value
                    date_col = None
                    for c in range(2, ws.max_column + 1):
                        if str(ws.cell(row=1, column=c).value) == current_date:
                            date_col = c
                            break
                    if date_col:
                        for r in range(2, ws.max_row + 1):
                            val = ws.cell(row=r, column=date_col).value
                            if val and str(val).upper() == "P":
                                count += 1
                wb.close()
            except:
                pass
        present_count_text.value = f"Total Present: {count}"
        page.update()

    def modify_attendance(action):
        if not roll_display.value:
            page.snack_bar = ft.SnackBar(ft.Text("Enter roll number!"), open=True)
            page.update()
            return
        target_roll = str(roll_display.value).strip()
        class_sheet = selected_class.value
        current_date = date_display.value
        try:
            wb = (
                openpyxl.load_workbook(EXCEL_FILE)
                if os.path.exists(EXCEL_FILE)
                else openpyxl.Workbook()
            )
            ws = (
                wb[class_sheet]
                if class_sheet in wb.sheetnames
                else wb.create_sheet(class_sheet)
            )
            if ws.cell(1, 1).value is None:
                ws.cell(1, 1).value = "Roll / Date"
            date_col = None
            for c in range(2, ws.max_column + 2):
                if str(ws.cell(row=1, column=c).value) == current_date:
                    date_col = c
                    break
                if ws.cell(row=1, column=c).value is None:
                    ws.cell(row=1, column=c).value = current_date
                    date_col = c
                    break
            all_rows = []
            for r in range(2, ws.max_row + 1):
                row_val = [
                    ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)
                ]
                if row_val and row_val[0] is not None:
                    all_rows.append(row_val)
            found = False
            for row in all_rows:
                if str(row[0]).strip() == target_roll:
                    row[date_col - 1] = "P" if action == "mark" else ""
                    found = True
                    break
            if not found:
                new_row = [None] * ws.max_column
                new_row[0] = target_roll
                new_row[date_col - 1] = "P" if action == "mark" else ""
                all_rows.append(new_row)
            all_rows.sort(
                key=lambda x: int(str(x[0]).strip()) if str(x[0]).isdigit() else 999
            )
            ws.delete_rows(2, ws.max_row + 10)
            for r_idx, values in enumerate(all_rows, start=2):
                for c_idx, val in enumerate(values, start=1):
                    ws.cell(row=r_idx, column=c_idx).value = val
            wb.save(EXCEL_FILE)
            update_present_count()
            page.snack_bar = ft.SnackBar(
                ft.Text(f"Roll {target_roll} Updated!"), open=True
            )
        except Exception as ex:
            page.snack_bar = ft.SnackBar(ft.Text(f"Error: {ex}"), open=True)
        roll_display.value = ""
        page.update()

    # --- RESET & DATE FUNCTIONS ---
    def on_date_change(e):
        date_display.value = date_picker.value.strftime("%d-%m-%Y")
        update_present_count()
        page.update()

    date_picker = ft.DatePicker(on_change=on_date_change)
    page.overlay.append(date_picker)

    def clear_date_data(e):
        if not os.path.exists(EXCEL_FILE):
            return
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if selected_class.value in wb.sheetnames:
                ws = wb[selected_class.value]
                current_date = date_display.value
                for c in range(2, ws.max_column + 1):
                    if str(ws.cell(row=1, column=c).value) == current_date:
                        for r in range(2, ws.max_row + 1):
                            ws.cell(row=r, column=c).value = None
                        break
                wb.save(EXCEL_FILE)
                update_present_count()
                page.snack_bar = ft.SnackBar(
                    ft.Text(f"Cleared {current_date}"), open=True, bgcolor="orange"
                )
        except:
            pass
        page.update()

    def reset_database(e):
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
            update_present_count()
            page.snack_bar = ft.SnackBar(
                ft.Text("Database Deleted!"), open=True, bgcolor="red"
            )
        page.update()

    def handle_export(e):
        if not os.path.exists(EXCEL_FILE):
            page.snack_bar = ft.SnackBar(ft.Text("No data to export!"), open=True)
            page.update()
            return

        try:
            if page.platform == ft.PagePlatform.WINDOWS:
                dest = os.path.join(
                    os.path.expanduser("~"), "Downloads", "Attendance_Export.xlsx"
                )
                shutil.copy2(EXCEL_FILE, dest)
                page.snack_bar = ft.SnackBar(
                    ft.Text("Saved to PC Downloads!"), open=True, bgcolor="blue"
                )
                page.update()
            else:
                # Tell the user it's working
                page.snack_bar = ft.SnackBar(
                    ft.Text("Preparing to share..."), open=True
                )
                page.update()

                # The file path is local to the phone in the APK
                page.share_binary(EXCEL_FILE, filename="Attendance_Report.xlsx")
        except Exception as ex:
            page.snack_bar = ft.SnackBar(ft.Text(f"Error: {ex}"), open=True)
            page.update()

    # --- UI LAYOUT ---
    def keypad_btn(val):
        return ft.FilledButton(
            content=ft.Text(str(val), size=20),
            width=75,
            height=75,
            on_click=lambda _: (
                setattr(roll_display, "value", roll_display.value + str(val))
                or page.update()
            ),
        )

    setup_view = ft.Column(
        [
            ft.Text("SESSION SETUP", size=20, weight="bold", color="blue700"),
            selected_class,
            ft.Container(height=20),
            ft.Row(
                [
                    date_display,
                    ft.IconButton(
                        ft.Icons.CALENDAR_MONTH,
                        on_click=lambda _: (
                            setattr(date_picker, "open", True) or page.update()
                        ),
                    ),
                ],
                alignment="center",
            ),
            ft.Container(height=10),
            ft.FilledButton(
                "Share_Excel",
                icon=ft.Icons.SHARE,
                on_click=handle_export,
                width=300,
                bgcolor="green700",
            ),
            ft.FilledButton(
                "Clear_Date",
                icon=ft.Icons.DELETE_SWEEP,
                on_click=clear_date_data,
                width=300,
                bgcolor="orange800",
            ),
            ft.FilledButton(
                "Reset_All",
                icon=ft.Icons.DANGEROUS,
                on_click=reset_database,
                width=300,
                bgcolor="red700",
            ),
            ft.Container(height=10),
            present_count_text,
        ],
        horizontal_alignment="center",
    )

    attendance_view = ft.Column(
        [
            ft.Text("MARKING", size=20, weight="bold", color="blue700"),
            roll_display,
            ft.Row([keypad_btn(1), keypad_btn(2), keypad_btn(3)], alignment="center"),
            ft.Row([keypad_btn(4), keypad_btn(5), keypad_btn(6)], alignment="center"),
            ft.Row([keypad_btn(7), keypad_btn(8), keypad_btn(9)], alignment="center"),
            ft.Row(
                [
                    ft.IconButton(
                        ft.Icons.DELETE,
                        on_click=lambda _: (
                            setattr(roll_display, "value", "") or page.update()
                        ),
                        icon_color="red",
                    ),
                    keypad_btn(0),
                    ft.IconButton(
                        ft.Icons.BACKSPACE,
                        on_click=lambda _: (
                            setattr(roll_display, "value", roll_display.value[:-1])
                            or page.update()
                        ),
                        icon_color="orange",
                    ),
                ],
                alignment="center",
            ),
            ft.Row(
                [
                    ft.FilledButton(
                        "Save",
                        on_click=lambda _: modify_attendance("mark"),
                        width=140,
                        bgcolor="green",
                    ),
                    ft.FilledButton(
                        "Remove",
                        on_click=lambda _: modify_attendance("remove"),
                        width=140,
                        bgcolor="red",
                    ),
                ],
                alignment="center",
            ),
        ],
        horizontal_alignment="center",
        visible=False,
    )

    def switch_nav(e):
        setup_view.visible = e.control.data == "s"
        attendance_view.visible = e.control.data == "m"
        page.update()

    page.add(
        ft.Row(
            [
                ft.TextButton("SETUP", data="s", on_click=switch_nav),
                ft.TextButton("MARKING", data="m", on_click=switch_nav),
            ],
            alignment="center",
        ),
        ft.Divider(),
        setup_view,
        attendance_view,
    )
    update_present_count()


if __name__ == "__main__":
    ft.run(main)
