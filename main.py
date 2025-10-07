# GOAL
# a time tracker app
# user to select the task from a task list drop down (data source for task list -> excel)
# has buttons -> start, pause, end, reset
# stores the start time, end time in Excel on click of end button
# resets timer on reset button click
import time
import customtkinter as ctk
import os
import pandas as pd
from PIL.ImageFile import ImageFile
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import datetime as dt
from enum import Enum
# for handling icons
# ImageTk for toolbar icon in about and manage task status windows
from PIL import Image, ImageDraw, ImageTk
# for handling file paths
import sys
# to create a separate thread for sys tray icon
import threading
# to create a sys tray icon and to create menu items for sys tray icon right click
import pystray
# platform detection
import platform

# -------------------
# Cross-platform flags
# -------------------
PLATFORM = platform.system()
IS_WINDOWS = PLATFORM == "Windows"
IS_MAC = PLATFORM == "Darwin"
IS_LINUX = PLATFORM == "Linux"

# ctypes.windll only exists on Windows; guard the import
try:
    if IS_WINDOWS:
        from ctypes import windll  # type: ignore
    else:
        windll = None  # type: ignore
except Exception:
    windll = None  # type: ignore


class TimerStatus(Enum):
    # to track the timer status for the buttons to work correctly
    RUNNING = 1
    PAUSED = 2
    STOPPED = 3

class TaskTimer:
    def __init__(self) -> None:
        # set the window theme to 'dark' mode
        ctk.set_appearance_mode("dark")
        # initialize the main window
        self.app = ctk.CTk()
        self.app_title = "Time Keeper"
        self.app.title(self.app_title)
        # disable window resizing
        self.app.resizable(False, False)

        # -------- window chrome (platform-specific) --------
        # On Windows we keep a borderless window and rely on the system tray menu.
        # On macOS we keep the normal titlebar so the app is always reachable from the Dock (no tray needed).
        if IS_WINDOWS:
            self.app.overrideredirect(True)
        else:
            self.app.overrideredirect(False)

        # Make sure our window actually comes to the front when launched (especially on macOS)
        # We define a helper and schedule it right after Tk finishes drawing.
        self.app.after(50, self._bring_to_front)

        # -----------assets-----------
        # Excel file to store the task list, time
        self.excel_file = "Time_Keeper.xlsx"
        # icon for the system tray (PNG/ICO both OK). Keep an ICO for Windows; PNG is safer elsewhere.
        self.app_icon = "app_icon.ico"
        # Sheet in the Excel file to store the task list
        self.excel_tasks_sheet = "Tasks"
        # name of the column storing tasks inside the Tasks sheet
        self.tasks_col_name = "Task"
        # Sheet in the Excel file to store the time for each task
        self.excel_time_sheet = "Time"
        # excel icon for excel_btn
        self.excel_btn_icon = "excel_btn_icon.png"
        self.task_active_status_symbol = "Active"

        # Tk iconphoto caches (avoid GC of PhotoImage)
        self._tk_icon_cache = {}
        self._apply_window_icon(self.app)

        # To track the work duration for the current date
        # We use .replace as only this works in _get_days_work_minutes() method
        self.current_date = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        self.days_work_minutes = self._get_days_work_minutes()

        # get the task list from the Excel if it exists, to populate task_list_menu combobox dropdown
        self.all_tasks_dict_list = None
        # this is only active tasks list
        self.task_list = self._get_task_list()
        # task selected from the task_list_menu combobox
        self.current_task = ""

        # track the timer status - running, paused, stopped
        self.is_timer_running: TimerStatus = TimerStatus.STOPPED
        # to display timer text inside the timer_display Entry
        self.timer_text = ctk.StringVar()
        # to track the number of seconds elapsed and to use to set the text for timer_display Entry via timer_text
        # self.seconds_elapsed_ui = 0
        self.task_start_time = None # to store the start time of the task
        self.task_end_time = None # to store the end time of the task
        self.segment_start_time = None # start of a segment, if paused
        self.segment_end_time = None # end of a segment, if paused
        self.work_seconds = 0
        # ----to track multi day task tracking variables----
        self.new_day_pause_start = None
        self.new_day_pause_seconds = 0
        self.work_seconds_logged = 0
        self.multiday_start_date = None # to group multiday tasks

        # to manage placeholder text in the notes_entry field
        # when is_placeholder_active is True, show PH text in the notes_entry filed
        self.is_placeholder_active = True

        # declared it here as these are used by multiple methods
        self.days_work_label = None
        self.task_list_menu = None
        self.status_label = None
        self.timer_display = None
        self.start_btn = None
        # Textbox for user to type in task notes
        self.notes_textbox = None
        # To be disabled when timer is running or paused
        self.manage_tasks_btn = None

        # to store the after() ID and to handle .after() calls overlaps i.e., to be used in .after_cancel()
        self.status_update_queue = None
        self.timer_running_queue = None

        # to handle drag and reposition of the app window (used on Windows where it's borderless)
        self.start_mouse_x_root = None
        self.start_mouse_y_root = None
        self.start_window_x_root = None
        self.start_window_y_root = None

        # to enable running and controlling the app from the system tray (Windows/Linux only)
        self.systray_icon = None
        if IS_WINDOWS or IS_LINUX:
            # create a separate thread for sys tray icon run so that mainloop() does not block this
            self.systray_thread = threading.Thread(target=self._initialize_systray_icon, daemon=True)
            # start the sys tray thread
            self.systray_thread.start()
        else:
            self.systray_thread = None

        # build the ui (widgets) of the app
        self._build_ui()

        # position app window in the bottom right corner of the screen
        self.position_window()

        # to track the system sleep/freeze/hang phases etc.
        self.last_ui_update_mono = time.monotonic()
        self.last_ui_update_time = dt.datetime.now()
        # perpetual loop to log last UI update time to detect system sleep/freeze/hang etc.
        self._schedule_update_timer()
        # to track if the end of the task is by user of by system so that the task_end_time and segment_end_time are set accordingly
        self.auto_end = False

        # start the loop to check if day has changed and update the day's duration display
        self.check_day_change_queue = self.app.after(60000, self._check_for_day_change_periodically)

        self.app.mainloop()


    def _get_days_work_minutes(self) -> int:
        """
        Get the total of days work minutes from the Excel when the app is opened
        If the Excel does not exist, reruns 0
        :return:
        """
        # 1. check if the Excel file exists
        if os.path.exists(self.excel_file) and os.path.getsize(self.excel_file) > 0:
            # check if the Time sheet exists and catch errors on read
            try:
                # 2. read the Time sheet in the Excel file
                time_df = pd.read_excel(self.excel_file, sheet_name=self.excel_time_sheet)

                # check if DF is not empty (e.g., only headers)
                if not time_df.empty:
                    # normalize dates to midnight to match self.current_date
                    if not pd.api.types.is_datetime64_any_dtype(time_df["Date"]):
                        time_df["Date"] = pd.to_datetime(time_df["Date"], errors="coerce")
                    days_time_df = time_df[ time_df["Date"].dt.normalize() == self.current_date ]
                    days_work_minutes_list = days_time_df["Work_Minutes"].fillna(0).tolist()
                    days_work_minutes = int(sum(days_work_minutes_list))
                    return days_work_minutes
            except pd.errors.ParserError as e:
                print(f"Error reading the file to get the Time list: {e}")
            except Exception as e:
                # catch any other unexpected errors
                print(f"An unexpected error occurred while getting the Time list: {e}")

        return 0


    def _check_for_day_change_periodically(self):
        """
        Checks if a new day has started every 1 minute, and if a new day is detected,
        resets the days work minutes to 0 (for UI display) and current date to new day's date
        triggers the previous day's data logging if the timer is not stopped i.e., paused or running
        """
        # get the calendar date and compare it with the date currently we are displaying the day's duration for
        if self.current_date != dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0):
            # a new day has started

            if self.is_timer_running != TimerStatus.STOPPED:
                # if the timer is running or is paused,
                # log the previous day's data to the Excel
                self._check_day_split_and_log()

            # update the UI to display the new day's work minutes, which would mostly be 0
            # _check_day_split_and_log() method also triggers _update_days_work_minutes_display() method
            # but only for current/new day's log and not for the previous day's log
            # so we call _update_days_work_minutes_display() here irrespective of timer running status
            self._update_days_work_minutes_display()

            print(f"{self.current_date=}, {self.days_work_minutes=}")

        # to ensure the check runs perpetually
        self.check_day_change_queue = self.app.after(60000, self._check_for_day_change_periodically)


    def _get_task_list(self):
        """
        Read tasks from the Excel file if it exists and is not empty
        Only includes tasks where Status='Active'
        :returns list: A list of tasks always starting with "<Add new task...>".
        Returns just ["<Add new task...>"] if the Excel file doesn't exist, is empty, or has invalid data.
        """

        tasks_list = []

        # 1. check if the file exists and is not empty
        if os.path.exists(self.excel_file) and os.path.getsize(self.excel_file) > 0:
            # to catch errors in reading the file
            try:
                # 2. read the Excel sheet into a DF
                tasks_df = pd.read_excel(self.excel_file, sheet_name=self.excel_tasks_sheet)
                # to check if a task exists on new task addition
                self.all_tasks_dict_list = tasks_df.to_dict(orient="records")

                # 3. check if DF is not empty (e.g., only headers)
                if not tasks_df.empty:
                    # 4. drop blanks in 'Tasks' column and filter out tasks with status != 'active'
                    active_tasks = tasks_df[
                        tasks_df[self.tasks_col_name].notna() &
                        (tasks_df["Status"].astype(str).str.lower() == self.task_active_status_symbol.lower())
                    ]

                    # 5. convert the tasks to a list
                    tasks_list = active_tasks[self.tasks_col_name].astype(str).to_list()
                    tasks_list.sort()

                else:
                    print(f"{self.excel_file} exists but contains no data.")
            except pd.errors.ParserError as e:
                print(f"Error reading the file to get the task list: {e}")
            except Exception as e:
                # catch any other unexpected errors
                print(f"An unexpected error occurred while getting the task list: {e}")
        else:
            print(f"{self.excel_file} doesn't exist or is empty.")

        return ["<Add new task...>"] + tasks_list


    def _list_menu_callback(self, choice):
        if choice == "<Add new task...>":
            self.task_list_menu.set("")
            self.current_task = ""
            self.task_list_menu.focus()
        else:
            self.current_task = choice
            self.app.focus()


    def _update_status_label(self, status: str, code: int):
        """
        Display status of a task for 3-seconds
        :param status: str
        :param code: int | 0 for success 1 for failure/error/warning
        """
        if self.status_update_queue is not None:
            self.app.after_cancel(self.status_update_queue)

        if code == 1:
            status = status + " :(" if status else status
            self.status_label.configure(text=status, text_color="#b54747")
        else:
            status = status + " :)" if status else status
            self.status_label.configure(text=status, text_color="#009933")

        if status:
            self.status_update_queue = self.app.after(3000,
                                                      lambda: self._update_status_label("", 1)
                                                      )


    def _append_data_to_excel(self, sheet_name, **kwargs) -> bool:
        """
        Appends new row of data to the specified sheet in the Excel file
        Creates the Excel file/new sheet if they don't exist
        Data is passed as keyword arguments and keywords become headers
        Example usage:
            _append_data_to_excel("Tasks", Task="Study Python", Status="Active", Added_On="2025-04-05 10:00")
            _append_data_to_excel("Time", Task="Study Python", Timestamp="2025-04-05 10:00", Notes="Great progress!")
        :param sheet_name: (str): Name of the sheet to append to
        :param kwargs: Each key becomes a column header, value becomes cell data
        :returns bool: True if successful, False otherwise
        """
        try:
            if os.path.exists(self.excel_file):
                try:
                    wb = load_workbook(self.excel_file)
                except (InvalidFileException, Exception) as e:
                    print(f"Error with existing file: {e}. Creating new file.")
                    wb = Workbook()
                    if len(wb.sheetnames) > 0:
                        for s in wb.sheetnames:
                            wb.remove(wb[s])
            else:
                wb = Workbook()
                if len(wb.sheetnames) > 0:
                    for s in wb.sheetnames:
                        wb.remove(wb[s])

            if sheet_name not in wb.sheetnames:
                sheet = wb.create_sheet(sheet_name)
                sheet.append(list(kwargs.keys()))
            else:
                sheet = wb[sheet_name]

            sheet.append( list(kwargs.values()) )

            wb.save(self.excel_file)
            wb.close()
            return True
        except Exception as e:
            print(f"Error on appending data to excel: {e}")
            return False


    def _show_placeholder(self):
        self.notes_textbox.insert("1.0", "Add notes")
        self.notes_textbox.configure(text_color="#7a848d")
        self.is_placeholder_active = True


    def _notes_focus_in(self, event):
        if self.is_placeholder_active:
            self.notes_textbox.delete("1.0", "end")
            self.notes_textbox.configure(text_color="#d2d9e0")
            self.is_placeholder_active = False


    def _notes_focus_out(self, event):
        if not self.notes_textbox.get("1.0", "end-1c").strip():
            self._show_placeholder()


    def _add_task_on_enter(self, event):
        new_task = self.task_list_menu.get().strip()

        if new_task:
            new_task = new_task[0].upper() + new_task[1:]
            if self.all_tasks_dict_list:
                does_task_exist = any(task_item[self.tasks_col_name].lower() == new_task.lower()
                                  for task_item in self.all_tasks_dict_list)
            else:
                does_task_exist = False

            if not does_task_exist:
                now = dt.datetime.now()
                write_status = self._append_data_to_excel(self.excel_tasks_sheet, Task=new_task,
                                                          Status=self.task_active_status_symbol,
                                                          Added_On=f"{now:%d-%b-%Y T%I:%M %p}")

                if write_status:
                    self.task_list = self._get_task_list()
                    self.task_list[1:] = sorted(self.task_list[1:])
                    self.task_list_menu.configure(values=self.task_list)
                    self.task_list_menu.set(new_task)
                    self.current_task = new_task
                    self._update_status_label("Added", 0)
                    self.app.focus()
                else:
                    self._update_status_label("Error", 1)
            else:
                self._update_status_label("Exists", 1)
        else:
            self._update_status_label("Empty", 1)


    def _update_timer_display(self):
        update_interval = 1.5

        if self.is_timer_running == TimerStatus.RUNNING:
            current_mono = time.monotonic()

            time_since_last_ui_update = current_mono - self.last_ui_update_mono

            if time_since_last_ui_update < update_interval:
                seconds_elapsed_ui = self.work_seconds + (dt.datetime.now() -
                                                          self.segment_start_time).total_seconds()
                seconds_elapsed_ui = int(seconds_elapsed_ui)
                hours_elapsed, remainder = divmod(seconds_elapsed_ui, 3600)
                minutes_elapsed, remainder = divmod(remainder, 60)
                self.timer_text.set(f"{hours_elapsed:02}:{minutes_elapsed:02}:{remainder:02}")
            else:
                self.segment_end_time = self.last_ui_update_time
                self.task_end_time = self.last_ui_update_time
                self.auto_end = True
                self._seconds_accumulator()
                self._end_timer()

        self.last_ui_update_time = dt.datetime.now()
        self.last_ui_update_mono = time.monotonic()
        self._schedule_update_timer()


    def _schedule_update_timer(self):
        self.timer_running_queue=  self.app.after(1000, self._update_timer_display)


    def _humanize_time(self, minutes):
        if minutes:
            hours, minutes = divmod(minutes, 60)
            return f"{hours:.0f}h {minutes:02.0f}m" if hours else f"{minutes:.0f}m"
        else:
            return 0


    def _seconds_accumulator(self):
        self.work_seconds += int((self.segment_end_time - self.segment_start_time).total_seconds())


    def _new_day_pause_seconds_accumulator(self, current_timestamp: dt.datetime) -> None:
        if self.new_day_pause_start:
            self.new_day_pause_seconds += int( (current_timestamp - self.new_day_pause_start).total_seconds())
            self.new_day_pause_start = None
        else:
            midnight_timestamp = dt.datetime.combine(current_timestamp, dt.time())
            self.new_day_pause_seconds += int( (current_timestamp - midnight_timestamp).total_seconds() )


    def _run_timer(self):
        if self.current_task:
            self.manage_tasks_btn.configure(command=lambda: ..., text_color="#353535")
            if self.is_timer_running != TimerStatus.RUNNING:
                if self.is_timer_running == TimerStatus.STOPPED:
                    self.task_start_time = dt.datetime.now()

                self.segment_start_time = dt.datetime.now()

                if self.is_timer_running == TimerStatus.PAUSED and self.task_start_time.date() != self.segment_start_time.date():
                    self._new_day_pause_seconds_accumulator(current_timestamp= self.segment_start_time)

                self.is_timer_running = TimerStatus.RUNNING
                self.start_btn.configure(text="⏸")
                self.task_list_menu.configure(state="disabled")
                self._update_status_label("Start", 0)
            else:
                self.segment_end_time = dt.datetime.now()
                if self.task_start_time.date() != self.segment_end_time.date():
                    self.new_day_pause_start = dt.datetime.now()
                self._seconds_accumulator()
                self.is_timer_running = TimerStatus.PAUSED
                self.start_btn.configure(text="▶")
                self._update_status_label("Pause", 0)
        else:
            self._update_status_label("Select", 1)
        self.app.focus()


    def _calculate_duration(self, task_start_time, task_end_time, work_seconds):
        task_start_trimmed = task_start_time.replace(second=0, microsecond=0)
        task_end_trimmed = task_end_time.replace(second=0, microsecond=0)

        total_task_seconds_trimmed = int( (task_end_trimmed - task_start_trimmed).total_seconds() )
        total_task_minutes = total_task_seconds_trimmed // 60

        total_actual_seconds = int( (task_end_time - task_start_time).total_seconds() )

        pause_seconds = max(0, total_actual_seconds - work_seconds)
        pause_minutes = pause_seconds // 60

        work_minutes = total_task_minutes - pause_minutes
        work_minutes = max(0, work_minutes)

        return work_minutes, pause_minutes, total_task_minutes


    def _update_days_work_minutes_display(self, current_task_work_minutes=0) -> None:
        if self.current_date == dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0):
            self.days_work_minutes += current_task_work_minutes
        else:
            self.current_date = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            self.days_work_minutes = self._get_days_work_minutes()

        days_work_minutes_formated = self._humanize_time(self.days_work_minutes)
        self.days_work_label.configure(text=f"Day: {days_work_minutes_formated}")


    def _log_data_to_excel(self, task_start_time, task_end_time, work_minutes, pause_minutes, total_task_minutes) -> bool:
        work_duration = self._humanize_time(work_minutes)
        pause_duration = self._humanize_time(pause_minutes)

        if self.is_placeholder_active:
            task_notes = ""
        else:
            task_notes = self.notes_textbox.get("1.0", "end-1c")

        log_status = self._append_data_to_excel(self.excel_time_sheet,
                                                Date=task_start_time.date(),
                                                Task=self.current_task,
                                                Work_Duration=work_duration,
                                                Notes=task_notes,
                                                Pause_Duration=pause_duration,
                                                Start_Time=f"{task_start_time:%I:%M %p}",
                                                End_Time=f"{task_end_time:%I:%M %p}",
                                                Work_Minutes=work_minutes,
                                                Pause_Minutes=pause_minutes,
                                                Total_Minutes=total_task_minutes,
                                                Multi_day_Start = f"{self.multiday_start_date}",
                                                )

        return log_status


    def _check_day_split_and_log(self) -> bool:
        current_timestamp = dt.datetime.now()
        if self.task_start_time.date() != current_timestamp.date():
            cumulative_work_seconds = self.work_seconds

            if self.is_timer_running == TimerStatus.RUNNING:
                cumulative_work_seconds += int( (current_timestamp - self.segment_start_time).total_seconds() )

            if self.is_timer_running == TimerStatus.PAUSED:
                self._new_day_pause_seconds_accumulator(current_timestamp=current_timestamp)

            midnight_timestamp = dt.datetime.combine(current_timestamp, dt.time())
            new_day_total_seconds = int( (current_timestamp - midnight_timestamp).total_seconds() )
            new_day_work_seconds = new_day_total_seconds - self.new_day_pause_seconds

            prev_day_work_seconds = cumulative_work_seconds - new_day_work_seconds

            self.multiday_start_date = f"{self.task_start_time.date()}"

            work_minutes, pause_minutes, total_task_minutes = self._calculate_duration(
                                                                task_start_time=self.task_start_time,
                                                                task_end_time=midnight_timestamp,
                                                                work_seconds=prev_day_work_seconds)

            prev_day_log_status = self._log_data_to_excel(task_start_time=self.task_start_time,
                                                          task_end_time=midnight_timestamp,
                                                          work_minutes=work_minutes,
                                                          pause_minutes=pause_minutes,
                                                          total_task_minutes=total_task_minutes)

            if prev_day_log_status:
                self.work_seconds_logged += prev_day_work_seconds
                self.task_start_time = midnight_timestamp
                self.new_day_pause_start = None
                self.new_day_pause_seconds = 0
            else:
                return False

        day_log_status = ""
        if self.is_timer_running == TimerStatus.STOPPED:
            current_day_work_seconds = self.work_seconds - self.work_seconds_logged

            work_minutes, pause_minutes, total_task_minutes = self._calculate_duration(
                                                                    task_start_time=self.task_start_time,
                                                                    task_end_time=self.task_end_time,
                                                                    work_seconds=current_day_work_seconds)

            day_log_status = self._log_data_to_excel(task_start_time=self.task_start_time,
                                                          task_end_time=self.task_end_time,
                                                          work_minutes=work_minutes,
                                                          pause_minutes=pause_minutes,
                                                          total_task_minutes=total_task_minutes)

            self._update_days_work_minutes_display(work_minutes)

        return day_log_status


    def _end_timer(self):
        if self.is_timer_running != TimerStatus.STOPPED:
            if not self.auto_end:
                self.task_end_time = dt.datetime.now()
                if self.is_timer_running != TimerStatus.PAUSED:
                    self.segment_end_time = dt.datetime.now()
                    self._seconds_accumulator()

            self.is_timer_running = TimerStatus.STOPPED
            log_status = self._check_day_split_and_log()

            if log_status:
                self._update_status_label("Saved", 0)
                self.auto_end = False
                self._reset_timer()
                if self.multiday_start_date: self.multiday_start_date = None
            else:
                self._update_status_label("Error", 1)
                self.is_timer_running = TimerStatus.PAUSED
                self.start_btn.configure(text="▶")


    def _reset_timer(self, status=""):
        self.is_timer_running = TimerStatus.STOPPED
        self.task_list_menu.configure(state="normal")
        self.current_task=""
        self.task_list_menu.set("")
        self.task_start_time = None
        self.task_end_time = None
        self.segment_start_time = None
        self.segment_end_time = None
        self.work_seconds = 0.0
        self.work_seconds_logged = 0
        self.timer_text.set("00:00:00")
        self.notes_textbox.delete("1.0", "end")
        self._show_placeholder()
        self.start_btn.configure(text="▶")
        self.manage_tasks_btn.configure(command=self._manage_task_status, text_color="#4a4a4a")
        self.app.focus()
        if status:
            self._update_status_label(status, 0)


    def _hide_app_window(self):
        """Minimise/hide (platform-specific)."""
        if IS_MAC:
            # iconify sends it to the Dock so users can bring it back easily
            self.app.iconify()
        else:
            # Withdraw keeps it out of taskbar – we rely on the tray on Win/Linux
            self.app.withdraw()


    def _start_drag(self, event):
        self.start_mouse_x_root = event.x_root
        self.start_mouse_y_root = event.y_root
        self.start_window_x_root = self.app.winfo_x()
        self.start_window_y_root = self.app.winfo_y()


    def _do_drag(self, event):
        deltax_root = event.x_root - self.start_mouse_x_root
        deltay_root = event.y_root - self.start_mouse_y_root
        window_new_x = self.start_window_x_root + deltax_root
        window_new_y = self.start_window_y_root + deltay_root
        self.app.geometry(f"+{window_new_x}+{window_new_y}")


    def _get_resource_path(self, file_name):
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS  # type: ignore
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, file_name)


    #---------icons & system tray [start]---------

    def _get_icon(self, icon_name) -> ImageFile:
        app_icon_path = self._get_resource_path(icon_name)
        try:
            if os.path.exists(app_icon_path):
                icon_image = Image.open(app_icon_path)
            else:
                # fallback: create a small generated icon
                icon_image = Image.new(mode="RGB", size=(36, 36), color="#05428b")
                drawer = ImageDraw.Draw(icon_image)
                drawer.text((10, 10), text="TK", fill="white")
            return icon_image
        except Exception:
            # final fallback
            icon_image = Image.new(mode="RGB", size=(36, 36), color="#05428b")
            drawer = ImageDraw.Draw(icon_image)
            drawer.text((10, 10), text="TK", fill="white")
            return icon_image

    def _apply_window_icon(self, window):
        """Use a cross-platform Tk icon (PhotoImage), keeps a ref to avoid GC."""
        try:
            pil_im = self._get_icon(self.app_icon)
            tk_im = ImageTk.PhotoImage(pil_im)
            window.iconphoto(True, tk_im)
            # keep a reference so it doesn't get garbage-collected
            self._tk_icon_cache[id(window)] = tk_im
        except Exception as e:
            print(f"Icon apply failed: {e}")

    def _initialize_systray_icon(self):
        """Initializes a python sys tray (pystray) icon. On macOS we skip using tray."""
        if IS_MAC:
            return
        try:
            menu_items = (
                pystray.MenuItem(f"Open {self.app_title}", self._show_app_window, default=True),
                pystray.MenuItem("Hide", self._hide_app_window),
                pystray.MenuItem("Quit", self._quit_app)
            )

            icon_image = self._get_icon(self.app_icon)
            self.systray_icon = pystray.Icon("time_keeper_widget", icon_image, f"{self.app_title} Widget", menu=menu_items)
            # detached is fine on Windows/Linux; on macOS we don't get here
            self.systray_icon.run_detached()
        except Exception as e:
            print(f"FATAL ERROR: System tray icon creation/run failed: {e}")
            # don't kill the app—gracefully continue without tray
            self.systray_icon = None

    def _show_app_window(self):
        is_app_visible = self.app.winfo_ismapped()
        if not is_app_visible:
            self.app.deiconify()
        self._bring_to_front()

    def _quit_app(self):
        self._end_timer()

        if self.status_update_queue:
            self.app.after_cancel(self.status_update_queue)

        if self.timer_running_queue:
            self.app.after_cancel(self.timer_running_queue)

        if self.check_day_change_queue:
            self.app.after_cancel(self.check_day_change_queue)

        if self.systray_icon:
            try:
                self.systray_icon.stop()
            except Exception:
                pass
            self.systray_icon = None

        self.app.destroy()
        sys.exit(0)

    # ---------icons & system tray [end]---------

    def _open_excel_file(self):
        if not os.path.exists(self.excel_file):
            self._update_status_label("Error", 1)
            return

        try:
            if sys.platform.startswith('win'):
                os.startfile(self.excel_file)
                self._update_status_label("Open", 0)
            elif sys.platform.startswith('darwin'):
                import subprocess
                subprocess.run(['open', self.excel_file], check=True)
                self._update_status_label("Opened", 0)
            elif sys.platform.startswith('linux'):
                import subprocess
                subprocess.run(['xdg-open', self.excel_file], check=True)
                self._update_status_label("Opened", 0)
            else:
                self._update_status_label("Error", 1)

        except (FileNotFoundError, Exception) as e:
            self._update_status_label("Error", 1)
            print(f"Error opening the file: {e}")


    def _open_about(self):
        about_window = ctk.CTkToplevel(self.app)
        about_window.title("About")
        about_window.resizable(False, False)

        # cross-platform icon
        self._apply_window_icon(about_window)

        about_window.grid_columnconfigure(1, weight=1)

        app_icon_image = ctk.CTkImage(light_image=self._get_icon("app_icon.ico"), dark_image=self._get_icon("app_icon.ico"), size=(48,48))
        app_icon_label = ctk.CTkLabel(about_window, text="", image=app_icon_image)
        app_icon_label.grid(row=1, column=1, padx=30, pady=(30, 5), sticky="we")


        app_name_label = ctk.CTkLabel(about_window, text=self.app_title, text_color="#6f7a83", font=ctk.CTkFont(weight="bold", size=16))
        app_name_label.grid(row=2, column=1, padx=30, sticky="we")

        description_text = "A lightweight & minimalistic app for time tracking.\nNo accounts, no fuss - just focus!"
        description_label = ctk.CTkLabel(about_window, text=description_text, text_color="#6f7a83",
                                      font=ctk.CTkFont(size=12), wraplength=190)
        description_label.grid(row=3, column=1, padx=20, pady=(12,0), sticky="we")

        brief_steps = ctk.CTkLabel(about_window, text="Add Task -> Track Time -> Log to Excel",
                                   font=ctk.CTkFont(size=12), text_color="#6f7a83")
        brief_steps.grid(row=4, column=1, sticky="we")

        author_label = ctk.CTkLabel(about_window, text="By Akshay (@akshay_r2 on X)", text_color="#6f7a83",
                                      font=("Segoe UI", 12, "bold"))
        author_label.grid(row=5, column=1, padx=30, pady=(12, 5), sticky="we")


    def _manage_task_status(self):
        manage_window = ctk.CTkToplevel(self.app)
        manage_window.title("Manage Tasks")
        manage_window.resizable(False, False)

        self._apply_window_icon(manage_window)

        def save_changes_to_task_status():
            active_tasks = [cb.cget("text") for cb in task_checkboxes if cb.get() == "on"]

            for task_item in self.all_tasks_dict_list:
                if task_item[self.tasks_col_name] in active_tasks:
                    task_item["Status"] = self.task_active_status_symbol
                else:
                    task_item["Status"] = ""

            save_error_message = ""
            try:
                wb = load_workbook(self.excel_file)
                if self.excel_tasks_sheet not in wb.sheetnames:
                    save_error_message= f"Error: {self.excel_tasks_sheet} sheet not found in '{self.excel_file}'. No update performed."
                    return

                ws = wb[self.excel_tasks_sheet]
                ws.delete_rows(2, ws.max_row - 1)

                for task_dict in self.all_tasks_dict_list:
                    row_values = [
                        task_dict.get(self.tasks_col_name, ""),
                        task_dict.get("Status", ""),
                        task_dict.get("Added_On", "")
                    ]
                    ws.append(row_values)

                wb.save(self.excel_file)

                self.task_list = self._get_task_list()
                self.task_list[1:] = sorted(self.task_list[1:])
                self.task_list_menu.configure(values=self.task_list)

                if self.is_timer_running == TimerStatus.STOPPED:
                    self.current_task = ""
                    self.task_list_menu.set("")
                manage_window.destroy()

            except (pd.errors.ParserError, FileNotFoundError, InvalidFileException, Exception) as err:
                print(err)
                save_error_message = "Error on save. Please try again :("
                save_error_label.configure(text=save_error_message)

        if self.all_tasks_dict_list:
            scrollable_frame = ctk.CTkScrollableFrame(manage_window, width=220, height=350)
            scrollable_frame.grid(row=1, column=1, padx=10, pady=(10,5))

            task_checkboxes = []

            sorted_tasks_dict_list = sorted(self.all_tasks_dict_list,
                                            key=lambda task_item: task_item[self.tasks_col_name])

            for item in sorted_tasks_dict_list:
                checked_state = "on" if item["Status"] == self.task_active_status_symbol else "off"
                checked_state_var = ctk.StringVar(value=checked_state)
                task_checkbox = ctk.CTkCheckBox(scrollable_frame, text=item[self.tasks_col_name],
                                                variable=checked_state_var, onvalue="on", offvalue="of",
                                                corner_radius=4, border_width=2, fg_color="#085bbe",
                                                hover_color="#05428b", font=("Segoe UI", 14))
                task_checkbox._text_label.configure(wraplength=280)
                task_checkboxes.append(task_checkbox)
                task_checkbox.grid(row=len(task_checkboxes), column=1, padx=10, pady=8, sticky="we")

            save_error_label = ctk.CTkLabel(manage_window, text="️", text_color="#afb4ba")
            save_error_label.grid(row=2, column=1, sticky="we")
            save_btn = ctk.CTkButton(manage_window, text="Save", fg_color="#085bbe", hover_color="#05428b",
                                     command=save_changes_to_task_status)
            save_btn.grid(row=3, column=1, pady=(0,12))
        else:
            manage_window.geometry("250x420")
            error_label = ctk.CTkLabel(manage_window, text_color="#4f575d", text="No Tasks Found :(", font=("Segoe UI", 16, "bold"))
            manage_window.grid_columnconfigure(1, weight=1)
            manage_window.grid_rowconfigure(1, weight=1)
            error_label.grid(row=1, column=1, sticky="nswe")


    def _build_ui(self):
        toolbar_frame = ctk.CTkFrame(self.app, height=30, fg_color="#2c2c2c", corner_radius=0)
        toolbar_frame.grid(row=1, column=1, columnspan=3, sticky="we")

        title_label = ctk.CTkLabel(toolbar_frame, text=self.app_title, font=ctk.CTkFont(size=12))
        title_label.grid(row=1, column=1, sticky="w", pady=(7,3), padx=10)

        days_work_minutes_formated = self._humanize_time(self.days_work_minutes)
        self.days_work_label = ctk.CTkLabel(toolbar_frame, text=f"Day: {days_work_minutes_formated}",
                                       text_color="#4f575d", font=("Segoe UI", 13, "bold"))
        self.days_work_label.grid(row=1, column=2, sticky="e")

        custom_minimize_btn = ctk.CTkButton(toolbar_frame, text="\u2013", fg_color="#343638",
                                            hover_color="#585a5c", width=40, height=20,
                                            font=("Segoe UI Symbol", 15),
                                            command=self._hide_app_window)
        custom_minimize_btn.grid(row=1, column=3,  padx=(10,5), pady=5)
        toolbar_frame.grid_columnconfigure(2, weight=1)

        # Only enable drag moving when the window is borderless (Windows)
        if IS_WINDOWS:
            for widget in (toolbar_frame, title_label):
                widget.bind("<Button-1>", self._start_drag)
                widget.bind("<B1-Motion>", self._do_drag)
            self.days_work_label.bind("<Button-1>", self._start_drag)
            self.days_work_label.bind("<B1-Motion>", self._do_drag)

        self.task_list_menu = ctk.CTkComboBox(self.app, values=self.task_list, command=self._list_menu_callback)
        self.task_list_menu.grid(row=2, column=1, padx=10, pady=(10, 0), sticky="ew", columnspan=3)
        self.task_list_menu.set("")
        self.task_list_menu.bind("<Return>", self._add_task_on_enter)

        hint_label = ctk.CTkLabel(self.app, text="Type new task & press Enter", font=("Segoe UI", 12, "bold"), height=5, text_color="#7a848d")
        hint_label.grid(row=3, column=1, columnspan=3, padx=10, sticky="w")

        self.status_label = ctk.CTkLabel(self.app, text="", font=("Segoe UI", 12, "bold"), height=5)
        self.status_label.grid(row=3, column=2, columnspan=3, sticky="e", padx=(0, 11))

        initial_timer_text= "00:00:00"
        self.timer_text.set(initial_timer_text)
        self.timer_display = ctk.CTkEntry(self.app, textvariable=self.timer_text, height=70,
                                     font=("Segoe UI Symbol", 40, "bold"), justify="center", state="disabled",
                                     text_color="#9e9e9e")
        self.timer_display.grid(row=4, column=1, columnspan=3, padx=10, pady=10, sticky="we")

        self.notes_textbox = ctk.CTkTextbox(self.app, text_color="#f2f2f2", height=62,
                                          font=("Segoe UI", 14), wrap="word",
                                          border_width=1, border_color="#4c5154")
        self.notes_textbox.grid(row=5, column=1, columnspan=3, sticky="we", padx=(10,8))

        self._show_placeholder()
        self.notes_textbox.bind("<FocusIn>", self._notes_focus_in)
        self.notes_textbox.bind("<FocusOut>", self._notes_focus_out)

        buttons_frame = ctk.CTkFrame(self.app, fg_color="transparent", height=30)
        buttons_frame.grid(row=6, column=1, columnspan=3, sticky="we", pady=(10,0))

        buttons_frame.grid_columnconfigure(3, weight=2)

        self.start_btn = ctk.CTkButton(buttons_frame, text="▶", command=self._run_timer, width=49,
                                       font=("Segoe UI Symbol", 16, "bold"), fg_color="#085bbe",
                                       hover_color="#05428b")
        self.start_btn.grid(padx=(10,8), row=1, column=1, sticky="we")

        end_btn = ctk.CTkButton(buttons_frame, text="⏹", width=49, fg_color="#085bbe",
                                font=("Segoe UI Symbol", 16, "bold"), hover_color="#05428b",
                                command=self._end_timer)
        end_btn.grid(row=1, column=2, sticky="we")

        reset_btn = ctk.CTkButton(buttons_frame, text="Reset", width=60, fg_color="#242424", hover_color="#414449",
                                  border_color="#414449", border_width=1,
                                  command=lambda: self._reset_timer("Reset"))
        reset_btn.grid(padx=(8,5), row=1, column=3, sticky="we")

        excel_btn_icon = ctk.CTkImage(light_image=self._get_icon(self.excel_btn_icon),
                                      dark_image=self._get_icon(self.excel_btn_icon), size=(19,19))
        open_excel_btn = ctk.CTkButton(buttons_frame, image=excel_btn_icon, fg_color="#242424",
                                       border_color="#414449", border_width=0, hover_color="#414449", width=1,
                                       text="", command=self._open_excel_file)
        open_excel_btn.grid(row=1, column=4, padx=(0,10), sticky="w")

        signature_label = ctk.CTkLabel(self.app, text="akshay;)", text_color="#262626",  height=5,
                                         font=ctk.CTkFont(size=8, weight="bold", slant="italic"))
        signature_label.grid(row=7, column=3, sticky="se")
        about_btn = ctk.CTkButton(self.app, text="About", height=1, fg_color="transparent", width=1,
                                         hover_color="#2c2c2c", text_color="#4a4a4a", font=ctk.CTkFont(weight="bold"), command=self._open_about)
        about_btn.grid(row=7, column=1, padx=(8,6), pady=3, sticky="we")

        self.manage_tasks_btn = ctk.CTkButton(self.app, text="Manage Tasks", height=1,
                                              fg_color="transparent", width=1, hover_color="#2c2c2c",
                                              text_color="#4a4a4a", font=ctk.CTkFont( weight="bold"),
                                              command=self._manage_task_status)
        self.manage_tasks_btn.grid(row=7, column=2, pady=3, sticky="we")

        quit_btn = ctk.CTkButton(self.app, text="Quit", height=1, fg_color="transparent", width=1,
                                         hover_color="#2c2c2c", text_color="#4a4a4a", font=ctk.CTkFont(weight="bold"), command=self._quit_app)
        quit_btn.grid(row=7, column=3, padx=(8,10), pady=3, sticky="we")


    def _get_dpi_scaling(self, hwnd):
        """
        Returns the DPI scaling factor for the given window handle on Windows.
        On macOS/Linux returns 1.0 (Tk handles scaling automatically).
        """
        try:
            if IS_WINDOWS and windll is not None:
                dpi = windll.user32.GetDpiForWindow(hwnd)
                return dpi / 96.0
        except Exception as e:
            print(f"Error getting DPI: {e}")
        return 1.0


    def _bring_to_front(self):
        """Bring the window to the front reliably on macOS/Windows without leaving it always-on-top."""
        try:
            # Raise above others briefly
            self.app.lift()
            self.app.attributes('-topmost', True)
            # Give the window input focus
            try:
                self.app.focus_force()
            except Exception:
                pass
            # After a short delay, release always-on-top so it behaves normally
            self.app.after(700, lambda: self.app.attributes('-topmost', False))
        except Exception as e:
            print(f"fronting failed: {e}")

    def _mac_work_area(self):
        """Return (vis_x, vis_y, vis_w, vis_h, dock_side) on macOS using AppKit if available.
        dock_side in {"bottom","left","right","unknown"}. Fallback returns None.
        """
        try:
            from AppKit import NSScreen  # type: ignore
            main = NSScreen.mainScreen()
            if not main:
                return None
            f = main.frame()
            v = main.visibleFrame()
            # Convert to ints
            frame = (int(f.origin.x), int(f.origin.y), int(f.size.width), int(f.size.height))
            vis = (int(v.origin.x), int(v.origin.y), int(v.size.width), int(v.size.height))
            fx, fy, fw, fh = frame
            vx, vy, vw, vh = vis
            # Detect dock side by differences between frame and visibleFrame
            dock_side = "unknown"
            # Bottom dock reduces visible height and raises visible origin.y
            if vh < fh and vy > fy and vw == fw and vx == fx:
                dock_side = "bottom"
            # Left dock reduces visible width and raises visible origin.x
            elif vw < fw and vx > fx and vh == fh and vy == fy:
                dock_side = "left"
            # Right dock reduces visible width but keeps origin.x
            elif vw < fw and vx == fx and vh == fh and vy == fy:
                dock_side = "right"
            return (vx, vy, vw, vh, dock_side)
        except Exception:
            return None

    def position_window(self):
        self.app.update_idletasks()

        right_margin = 5
        app_width = 240
        app_height = 280

        # Default margins
        bottom_margin = 85 if IS_WINDOWS else 40

        screen_width_logical = self.app.winfo_screenwidth()
        screen_height_logical = self.app.winfo_screenheight()

        if IS_MAC:
            wa = self._mac_work_area()
            if wa:
                vx, vy, vw, vh, dock_side = wa
                # Base padding away from Dock edge
                edge_pad = 16
                # Default placement: bottom-right of visible area
                x_logical = vx + vw - app_width - right_margin
                # For y, we need to convert from NSScreen's bottom-left origin to Tk's top-left
                # y_top = screen_height - (visible_bottom + app_height)
                y_bottom = vy + edge_pad
                y_logical = screen_height_logical - (y_bottom + app_height)

                if dock_side == "left":
                    # Keep extra space on the right, but ensure we don't collide with left Dock
                    x_logical = vx + vw - app_width - edge_pad
                elif dock_side == "right":
                    # Shift further left to avoid right Dock
                    x_logical = vx + vw - app_width - edge_pad
                elif dock_side == "bottom":
                    # Lift a little more from bottom
                    y_bottom = vy + edge_pad
                    y_logical = screen_height_logical - (y_bottom + app_height)
            else:
                # Fallback without AppKit: conservative margins
                bottom_margin = 100
                right_margin = 12
                x_logical = screen_width_logical - app_width - right_margin
                y_logical = screen_height_logical - app_height - bottom_margin

            x_physical = int(x_logical)
            y_physical = int(y_logical)
        else:
            # Windows/Linux
            hwnd = self.app.winfo_id()
            dpi_scale_factor = self._get_dpi_scaling(hwnd)

            x_logical = screen_width_logical - app_width - right_margin
            y_logical = screen_height_logical - app_height - bottom_margin

            if IS_WINDOWS:
                x_physical = int(x_logical * dpi_scale_factor)
                y_physical = int(y_logical * dpi_scale_factor)
            else:
                x_physical = int(x_logical)
                y_physical = int(y_logical)

        # Guard against negative placement (rare multi-monitor layouts)
        x_physical = max(0, x_physical)
        y_physical = max(0, y_physical)

        self.app.geometry(f"+{x_physical}+{y_physical}")


app = TaskTimer()
