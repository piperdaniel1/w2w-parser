from typing import List, Dict
import xlsxwriter as xl
from xlsxwriter.workbook import Format

#######################################
#
# BASE DATAYPE CLASS DEFINITIONS
#
#######################################
class TimeLength:
    def __init__(self, hours: int, minutes: int):
        self.hours = hours
        self.minutes = minutes

    def get_hours(self):
        return self.hours

    def get_minutes(self):
        return self.minutes

    # Returns strings like 3 hours 40 minutes
    def get_str1(self):
        return f"{self.hours} hours {self.minutes} minutes"

    def __eq__(self, __o: "TimeLength") -> bool:
        if type(__o) != type(self):
            return False

        return (__o.minutes == self.minutes) and (__o.hours == self.hours)

class DayTime:
    def __init__(self, hours: int, minutes: int):
        self.hours = hours
        self.minutes = minutes

    def get_hours(self):
        return self.hours

    def get_minutes(self):
        return self.minutes

    # get strings like 7:15am
    def get_str1(self):
        phours = self.hours
        suffix = "am"
        if phours > 12:
            phours -= 12
            suffix = "pm"
        elif phours == 0:
            suffix = "am"
            phours = 12
        elif phours == 12:
            suffix = "pm"

        pmins = self.minutes
        if pmins < 10:
            pmins = "0" + str(pmins)
        
        return f"{phours}:{pmins}{suffix}"

    def __eq__(self, __o: "DayTime") -> bool:
        if type(__o) != type(self):
            return False

        return (__o.minutes == self.minutes) and (__o.hours == self.hours)

    def __gt__(self, __o: "DayTime") -> bool:
        if self.hours > __o.hours:
            return True
        elif self.hours < __o.hours:
            return False
        else:
            if self.minutes > __o.minutes:
                return True
            else:
                return False

    def __ge__(self, __o: "DayTime") -> bool:
        if self.hours > __o.hours:
            return True
        elif self.hours < __o.hours:
            return False
        else:
            if self.minutes >= __o.minutes:
                return True
            else:
                return False

    def __lt__(self, __o: "DayTime") -> bool:
        if self.hours > __o.hours:
            return False
        elif self.hours < __o.hours:
            return True
        else:
            if self.minutes >= __o.minutes:
                return False
            else:
                return True

    def __le__(self, __o: "DayTime") -> bool:
        if self.hours > __o.hours:
            return False
        elif self.hours < __o.hours:
            return True
        else:
            if self.minutes > __o.minutes:
                return False
            else:
                return True

class Employee:
    def __init__(self, first_name: str, last_name: str):
        self.first_name = first_name
        self.last_name = last_name

    def get_first_name(self):
        return self.first_name

    def get_last_name(self):
        return self.last_name

    def get_full_name(self) -> str:
        return self.first_name + " " + self.last_name


class Shift:
    def __init__(self, employee: Employee, start_time: DayTime, \
                       end_time: DayTime, weekday: int):
        self.employee = employee
        self.start_time = start_time
        self.end_time = end_time
        self.weekday = weekday

    def get_start_time(self):
        return self.start_time

    def get_end_time(self):
        return self.end_time

    def get_length(self):
        return diff_between_times(self.start_time, self.end_time)

    def get_weekday(self):
        return self.weekday

    # Returns strings like 7:30am to 10:00am
    def get_str1(self):
        return f"{self.employee.get_first_name()} works from {self.start_time.get_str1()} to {self.end_time.get_str1()}"

    # Returns strings like 7:30am to 10:00am (2 hours 30 minutes)
    def get_str2(self):
        return self.get_str1() + " (" + self.get_length().get_str1() + ")"
    
    def __str__(self):
        return self.get_str1()

    def __repr__(self):
        return self.get_str1()

# Reprents a day of the week in the excel sheet.
# This class is needed to represent how the shifts will fit into the three 
# (or less/more) excel columns that that day has (because of overlapping shifts).
# Each shift in shift_list should have the same day of the week.
class OutputSlot:
    def __init__(self, shift_list: List[Shift]):
        for i in range(len(shift_list)-1):
            try:
                assert(shift_list[i].weekday == shift_list[i+1].weekday)
            except AssertionError:
                raise AssertionError(f"Invalid shift list provided to output slot: Shift #{i} ({shift_list[i]}) is on day {shift_list[i].weekday}, while the next shift is on day {shift_list[i+1].weekday}.")
                
        self.slots = self.__get_slot_lists(shift_list)

    def __get_slot_lists(self, shift_list : List[Shift]) -> List[List[Shift]]:
        curr_list : List[List[Shift]] = [[]]

        for shift in shift_list:
            for _, slot in enumerate(curr_list):
                if len(slot) == 0:
                    slot.append(shift)
                    break
                elif slot[-1].end_time <= shift.start_time:
                    slot.append(shift)
                    break

            if len(curr_list[-1]) != 0:
                curr_list.append([])

        return curr_list
    
    def get_slot(self, slot_index : int):
        return self.slots[slot_index]
    
    def get_slot_list(self):
        return self.slots[:-1]

# Represents an entire week in the excel sheet. In practice it partitions the shift_list into days of the week and passes it onto OutputSlots
class OutputWeek:
    def __init__(self, shift_list: List[Shift]) -> None:
        self.output_slots : List[OutputSlot] = []
        self.master_list = sorted(shift_list, key=lambda x: x.weekday)

        i = 0
        while i < len(self.master_list):
            curr_day = self.master_list[i].weekday
            curr_list = []
            
            while self.master_list[i].weekday == curr_day:
                curr_list.append(self.master_list[i])
                i += 1

                if i >= len(self.master_list):
                    break

            if len(curr_list) != 0:
                self.output_slots.append(OutputSlot(curr_list))

    def gen_xl_file(self, title : str):
        workbook = xl.Workbook(title.replace(" ", "-") + ".xlsx")
        # ABSOLUTE DEFAULT if no custom format is left
        fallback_format = workbook.add_format(properties={'bold': True, 
                                                       'font_color': 'yellow',
                                                       'bg_color': 'black',
                                                       'valign': 'top'})

        checkin_format = workbook.add_format(properties={'bold': True, 
                                                       'font_color': 'white',
                                                       'bg_color': 'pink',
                                                       'align': 'center'})

        header_format = workbook.add_format(properties={'bold': True, 
                                                        'font_size': 25,
                                                        'align': 'center',
                                                        'valign': 'center'})
        week_format = workbook.add_format(properties={'bold': True}) 

        # All formats to be used for each employee
        unused_formats = [
            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#70ad47',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#ff0404',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'black',
                                           'bg_color': '#ffd966',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'black',
                                           'bg_color': '#8ea9db',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#305496',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'black',
                                           'bg_color': '#ed7d31',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#7030a0',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#7b7b7b',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'black',
                                           'bg_color': '#99e4ff',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#833c0c',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#bf8f00',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#548235',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'black',
                                           'bg_color': '#fa63ff',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'white',
                                           'bg_color': '#9e2022',
                                           'valign': 'top'}),

            workbook.add_format(properties={'bold': True, 
                                           'font_color': 'yellow',
                                           'bg_color': 'black',
                                           'valign': 'top'}),
        ]

        # Store which employee gets which format
        reserved_fmt_dict : Dict[str, Format] = {}

        worksheet = workbook.add_worksheet(title)

        SHEET_WIDTH = self.get_total_slot_width()

        # WRITE THE HEADER
        worksheet.merge_range(0, 0, 0, SHEET_WIDTH, title, header_format)

        # WRITE THE COLUMN LABELS
        # Day COLUMN labels are written when shifts are written
        worksheet.write(1, 0, "Time")

        # WRITE THE TIME LABELS
        worksheet.write(2, 0, "7:30 AM")
        worksheet.write(3, 0, "8:00 AM")
        worksheet.write(4, 0, "9:00 AM")
        worksheet.write(5, 0, "10:00 AM")
        worksheet.write(6, 0, "11:00 AM")
        worksheet.write(7, 0, "12:00 PM")
        worksheet.write(8, 0, "1:00 PM")
        worksheet.write(9, 0, "2:00 PM")
        worksheet.write(10, 0, "3:00 PM")
        worksheet.write(11, 0, "4:00 PM")
        worksheet.write(12, 0, "5:00 PM")
        worksheet.write(13, 0, "6:00 PM")
        worksheet.write(14, 0, "7:00 PM")
        worksheet.write(15, 0, "8:00 PM")
        worksheet.write(16, 0, "9:00 PM")
        worksheet.write(17, 0, "10:00 PM")
        worksheet.write(18, 0, "11:00 PM")

        # ADD WEEKLY CHECKIN
        worksheet.merge_range(10, 1, 10, len(self.output_slots[0].get_slot_list()),
                "ALL PRESENT", checkin_format)

        # WRITE THE SHIFTS
        curr_col = 1
        open_time = DayTime(7, 0)

        for i, slot in enumerate(self.output_slots):
            worksheet.merge_range(1, curr_col,\
                    1, curr_col+len(slot.get_slot_list())-1, get_weekday_str(i), week_format)

            for shift_col in slot.get_slot_list():
                for shift in shift_col:
                    row = diff_between_times(shift.get_start_time(), \
                                             open_time).get_hours() + 2 

                    emp_format  = fallback_format
                    if shift.employee.get_full_name() in reserved_fmt_dict:
                        emp_format = reserved_fmt_dict[shift.employee.get_full_name()]
                    elif len(unused_formats) > 0:
                        new_fmt = unused_formats.pop(0)
                        reserved_fmt_dict[shift.employee.get_full_name()] = new_fmt
                        emp_format = new_fmt

                    shift_length = shift.get_length().get_hours()
                    if shift.get_length().get_minutes() > 0:
                        shift_length += 1

                    if shift_length > 1:
                        worksheet.merge_range(row, curr_col, row + shift_length-1, curr_col, shift.employee.first_name, emp_format)
                    else:
                        worksheet.write(row, curr_col, shift.employee.first_name, emp_format)
                curr_col += 1
        workbook.close()

    def get_total_slot_width(self):
        total = 0
        for slot in self.output_slots:
            total += len(slot.get_slot_list())
        return total
                
    def get_day(self, day_index : int):
        return self.output_slots[day_index]
    
    def get_days(self):
        return self.output_slots

        
#######################################
#
# HELPER FUNCTIONS TO WORK WITH THE DATATYPES
#
#######################################
def diff_between_times(t1: DayTime, t2: DayTime):
    if t1 < t2:
        hour_diff = t2.hours - t1.hours
        minute_diff = t2.minutes - t1.minutes
    else:
        hour_diff = t1.hours - t2.hours
        minute_diff = t1.minutes - t2.minutes

    if minute_diff < 0:
        hour_diff -= 1
    minute_diff = abs(minute_diff)

    return TimeLength(hour_diff, minute_diff)

# Accepts 12hr times like 8am, 8:00am, 12:30pm, 07:12pm etc
# Must include am or pm at the end.
def daytime_from_str1(time_str: str):
    period_of_day = time_str[-2:]
    rem_str = time_str[:-2]

    str_split = rem_str.split(":")
    if len(str_split) == 2:
        hours = int(str_split[0])
        mins = int(str_split[1])
    else:
        hours = int(str_split[0])
        mins = 0

    if period_of_day == "pm" and hours != 12:
        hours += 12

    if period_of_day == "am" and hours == 12:
        hours -= 12

    dt_to_ret = DayTime(hours, mins)

    return dt_to_ret

def get_weekday_str(weekday: int):
    weekday_list = ["Monday",
                    "Tuesday",
                    "Wednesday",
                    "Thursday",
                    "Friday",
                    "Saturday",
                    "Sunday"]
    
    return weekday_list[weekday]

def get_short_weekday_str(weekday: int):
    weekday_list = ["Mon",
                    "Tue",
                    "Wed",
                    "Thu",
                    "Fri",
                    "Sat",
                    "Sun"]
    
    return weekday_list[weekday]









