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


async def main(page: ft.Page):  # ← MUST be async
    page.title = "Attendance System (Universal)"
    page.theme_mode = ft.ThemeMode.DARK
    page.window_width = 380
    page.window_height = 800
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.ADAPTIVE

    EXCEL_FILE = get_db_path()

    # Share service (official way)
    share_service = ft.Share()

    # --- UI CONTROLS (unchanged) ---
    roll_display = ft.TextField(
        label="Roll Number",
        read_only=True,
        text_align=ft.TextAlign.CENTER,
        width=250,
        text_size=30,
        # color="blue",
        border_radius=15,
        border_color="red",
    )
    present_count_text = ft.Text(
        "Total Present: 0", size=14, weight="bold", color="green700"
    )

    selected_class = ft.Dropdown(
        label="Select Class",
        width=300,
        value="UG_1",
        border_color="red",
        options=[
            ft.dropdown.Option("UG_1"),
            ft.dropdown.Option("UG_2"),
            ft.dropdown.Option("UG_3"),
            ft.dropdown.Option("PG_1"),
            ft.dropdown.Option("PG_2"),
        ],
    )

    # 1. Initialize the variable first
    date_display = ft.TextField(
        label="Selected Date",
        value=datetime.now().strftime("%d-%m-%Y"),
        read_only=True,
        width=300,
        suffix_icon=ft.Icons.CALENDAR_MONTH,
        border_color="red",
    )

    # 2. Define the functions that use that variable
    async def handle_date_change(e):
        nonlocal date_display  # <--- This is the magic line
        date_display.value = date_picker.value.strftime("%d-%m-%Y")
        update_present_count()
        page.update()

    def open_date_picker(e):
        date_picker.open = True
        page.update()

    # 3. Assign the function to the on_click now that it's defined
    date_display.on_click = open_date_picker

    # 4. Create the Picker
    date_picker = ft.DatePicker(
        on_change=handle_date_change,
        first_date=datetime(2024, 1, 1),
        last_date=datetime(2030, 12, 31),
    )
    page.overlay.append(date_picker)

    # --- CORE LOGIC (update_present_count & modify_attendance unchanged) ---
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

    # modify_attendance function remains exactly the same as your original
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
                if row_val[0] is not None:
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

    # --- FIXED EXPORT (this is the part that now works perfectly on Android) ---
    async def handle_export(e):
        if not os.path.exists(EXCEL_FILE):
            page.snack_bar = ft.SnackBar(ft.Text("No data to export!"), open=True)
            page.update()
            return

        try:
            if page.platform in (ft.PagePlatform.WINDOWS, ft.PagePlatform.LINUX):
                dest = os.path.join(
                    os.path.expanduser("~"),
                    f"Attendance_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                )
                shutil.copy2(EXCEL_FILE, dest)
                msg = "Saved to Downloads!"
                color = "blue"
            else:  # Android / iOS
                # Get app's safe temporary directory (this is the key for Android)
                sp = ft.StoragePaths()
                temp_dir = await sp.get_temporary_directory()

                export_name = (
                    f"Attendance_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )
                temp_file = os.path.join(temp_dir, export_name)

                shutil.copy2(EXCEL_FILE, temp_file)

                # Official Flet sharing (works on Android)
                result = await share_service.share_files(
                    [ft.ShareFile.from_path(temp_file)], text="Master Attendance Report"
                )
                msg = f"Share sheet opened! ({result.status})"
                color = "green"

            page.snack_bar = ft.SnackBar(ft.Text(msg), open=True, bgcolor=color)
        except Exception as ex:
            page.snack_bar = ft.SnackBar(
                ft.Text(f"Share error: {str(ex)[:80]}"), open=True, bgcolor="red"
            )
        page.update()

        # --- RESET ALL DATA LOGIC ---

    async def handle_reset_confirmed(e):
        confirm_dialog.open = False
        if os.path.exists(EXCEL_FILE):
            try:
                os.remove(EXCEL_FILE)
                update_present_count()  # This resets the counter to 0
                page.snack_bar = ft.SnackBar(
                    ft.Text("Database Deleted!"), bgcolor="red", open=True
                )
            except Exception as ex:
                page.snack_bar = ft.SnackBar(ft.Text(f"Error: {ex}"), open=True)
        page.update()

    confirm_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Confirm Reset"),
        content=ft.Text(
            "This will permanently delete the Excel file and all attendance data."
        ),
        actions=[
            ft.TextButton(
                "Yes, Delete",
                on_click=handle_reset_confirmed,
            ),
            ft.TextButton(
                "Cancel",
                on_click=lambda _: (
                    setattr(confirm_dialog, "open", False) or page.update()
                ),
            ),
        ],
    )
    page.overlay.append(confirm_dialog)

    def show_reset_dialog(e):
        confirm_dialog.open = True
        page.update()

    # --- UI LAYOUT (only button changed) ---
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

    # 2. Define the logic functions
    async def handle_date_change(e):
        # Now this can find date_display because it's defined above
        date_display.value = date_picker.value.strftime("%d-%m-%Y")
        update_present_count()
        page.update()

    def open_date_picker(e):
        date_picker.open = True
        page.update()

    # 3. Define the actual Picker
    date_picker = ft.DatePicker(
        on_change=handle_date_change,
        first_date=datetime(2024, 1, 1),
        last_date=datetime(2030, 12, 31),
    )
    page.overlay.append(date_picker)

    setup_view = ft.Column(
        [
            ft.Text("Configration:", size=20, weight="bold", color="blue"),
            ft.Container(height=10),
            date_display,  # <--- MAKE SURE THIS LINE IS HERE
            ft.Container(height=15),
            selected_class,
            ft.Container(height=15),
            ft.FilledButton(
                "Share_Excel",
                icon=ft.Icons.SHARE,
                on_click=handle_export,
                width=300,
            ),
            ft.Container(height=15),
            ft.FilledButton(
                "Reset_All",
                icon=ft.Icons.DELETE_FOREVER,
                on_click=show_reset_dialog,
                width=300,
                bgcolor="blue",
            ),
            present_count_text,
            # Developer
            ft.Divider(height=40, thickness=1),
            ft.Container(
                content=ft.Column(
                    [
                        # ft.Text(
                        #     "Developer Information",
                        #     size=16,
                        #     weight="bold",
                        #     color="blue_grey_700",
                        # ),
                        ft.Text(
                            "Developer: Dr Manish kumar Gupta", size=12, weight="w500"
                        ),
                        ft.Text("Asst. Professor,Faculty of Commerce", size=10),
                        ft.Text(
                            "SMM Town PG College, Ballia",
                            size=10,
                            italic=True,
                            color="w500",
                        ),
                    ],
                    spacing=5,
                    horizontal_alignment="center",
                ),
                padding=20,
                border=ft.border.all(1, ft.Colors.BLUE_GREY_100),
                border_radius=15,
                width=300,
            ),
            ft.Text("Version 1.0.0", size=4, color="grey400"),
        ],
        horizontal_alignment="center",
    )

    attendance_view = ft.Column(
        [  # your existing attendance_view unchanged
            ft.Text("MARKING", size=20, weight="bold"),
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
                    ),
                    keypad_btn(0),
                    ft.IconButton(
                        ft.Icons.BACKSPACE,
                        on_click=lambda _: (
                            setattr(roll_display, "value", roll_display.value[:-1])
                            or page.update()
                        ),
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


# --- RUN THE APP ---
ft.run(main)  # ← Use ft.run (current standard)

