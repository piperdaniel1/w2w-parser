from datatypes import *
from typing import List
import sys

def parse_st_time(st_str : str):
    first_quote = st_str.index('"', 0)
    second_quote = st_str.index('"', first_quote + 1)

    time_substr = st_str[first_quote+1:second_quote]
    split_time = time_substr.split(" - ")
    start_time = split_time[0]
    end_time = split_time[1]

    start_time = daytime_from_str1(start_time) 
    end_time = daytime_from_str1(end_time) 

    return start_time, end_time

def parse_ss_emp(ss_str : str):
    comma_split = ss_str.split(",")
    emp_name = comma_split[2]
    emp_name = emp_name[1:-1]

    return emp_name

def classify_line(line : str):
    if "st(" in line:
        return "st"
    elif "ss(" in line:
        return "ss"
    elif "sdb()" in line:
        return "sdb"
    
    return "err"

def main():
    if len(sys.argv) != 3:
        print("Must provide input/output file as command arg: python3 parse.py input_file output_file (will auto append .xlsx extension)")
        return
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2]

    with open(input_file) as f:
        line = f.readlines()[0]
        lsplit = line.split(";")

    curr_stime = None
    curr_etime = None

    shifts: List[Shift] = []
    weekday = -1

    for split in lsplit:
        if classify_line(split) == "st":
            curr_stime, curr_etime = parse_st_time(split)
        elif classify_line(split) == "ss":
            emp_name = parse_ss_emp(split)
            name_list = emp_name.split(" ")
            first_name = name_list[0]
            last_name = name_list[1]

            if curr_stime == None or curr_etime == None:
                raise ValueError("Cannot create shift without time.")

            emp = Employee(first_name, last_name)

            if weekday == 0 and curr_stime.get_hours() == 15:
                pass
                # SKIP BECAUSE WEEKLY CHECKIN
            else:
                shifts.append(Shift(emp, curr_stime, curr_etime, weekday))
        elif classify_line(split) == "sdb":
            weekday += 1

    '''
    curr_weekday = None
    for shift in shifts:
        if curr_weekday != shift.get_weekday():
            curr_weekday = shift.get_weekday()
            print(f"\n=== {get_weekday_str(curr_weekday)} ===")

        print(shift.get_str1())
    '''

    out = OutputWeek(shifts)
    out.gen_xl_file(output_file)
    print("Complete.")

if __name__ == "__main__":
    main()
