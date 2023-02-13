from datatypes import *
from typing import List
import sys
import requests
import time
from datetime import date

def grab_text_from_w2w(getNext=False):
    print("Obtaining text from When to Work:")
    print(" > Logging in... ", end="", flush=True)

    url = ""
    while True:
        try:
            with open(".hashed_req") as f:
                lines = f.readlines()
                url = lines[0].strip("\n")
                payload = lines[1].strip("\n")
            break
        except FileNotFoundError:
            print("Error! You don't have a .hashed_req file, so I can't login to When To Work.")
            res = input("Do you want to create one now [Y/n]? ")

            if "y" in res.lower():
                url = input(" > Enter the URL of the login request (can be found by logging in while monitoring the network requests sent): ")
                payload = input(" > Enter the payload of the login request (can be found by logging in while monitoring the network requests sent): ")
                if url == "":
                    url = "https://whentowork.com/cgi-bin/w2w.dll/login"
                
                if payload == "":
                    print("Cannot use a null payload. Exiting...")
                    return None
                with open(".hashed_req", "w") as f:
                    f.write(url + "\n" + payload)
            else:
                return None

    resp = requests.post(url, data=payload)

    print("done")
    print(" > Accessing current schedule... ", end="", flush=True)

    session = resp.history[-1].headers["Location"].split("=")[1]
    full_url = resp.history[-1].headers["Location"]
    dll = full_url.split("/")[4]

    full_sched = f"https://www5.whentowork.com/cgi-bin/{dll}/empfullschedule?SID={session}&lmi="
    resp = requests.get(full_sched)

    print("done")

    if getNext:
        print(" > Accessing next schedule... ", end="", flush=True)
        full_sched += "&Week=Next"
        resp = requests.get(full_sched)
        print("done")

    print(" > Parsing response... ", end="", flush=True)
    time.sleep(0.5)
    main_text = resp.text.split("\n")
    for line in main_text:
        if "sdh(" in line:
            print("done")
            return line
    
    return None

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
    elif "sde()" in line:
        return "sde"
    elif "gl()" in line:
        return "gl"
    elif "weekly" in line.lower() \
        or "meeting" in line.lower() \
        or "checkin" in line.lower():
        return "meet"
    
    return "err"

def main():
    stime = time.time()
    if len(sys.argv) < 2:
        print("Usage: python parse.py <output_file> [--in <input_file>] [--next]")
        print("Tip: Add a .hashed_req file to the directory to use the w2w interface. \
            This file should contain the second when to work login request \
            (can be found by logging in while monitoring the network requests sent.")

    output_file = sys.argv[1]

    if output_file == "auto":
        WEEK_ONE_START = date(2023, 1, 7)
        today = date.today()

        offset = (today - WEEK_ONE_START).days

        weeks = offset // 7
        weeks += 1

        output_file = "Week " + str(weeks) + " Winter 2023 Schedule"
        pass

    lsplit = []

    if "--in" in sys.argv:
        print("Obtaining text from file:")
        print(" > Reading file... ", end="", flush=True)
        input_file = sys.argv[sys.argv.index("--in") + 1]
    
        try:
            with open(input_file) as f:
                line = f.readlines()[0]
                lsplit = line.split(";")
        except FileNotFoundError:
            print("Error: That input file was not found. Exiting...")
        print("done")
    else:
        line = grab_text_from_w2w("--next" in sys.argv)
        if line != None:
            lsplit = line.split(";")
        else:
            print("Error: Could not get line from when to work and no fallback input file provided. Exiting...")
            exit(1)

    curr_stime = None
    curr_etime = None

    shifts: List[Shift] = []
    weekday = 0

    is_weekly_meeting = False
    meeting: Dict[str, None | DayTime | int] = {
        "time": None,
        "day": None
    }

    emp_name = ""
    first_name = []
    last_name = []
    curr_stime = None
    curr_etime = None

    for i, split in enumerate(lsplit):
        # print(split)
        if i == len(lsplit) - 1:
            print("                                                         ", end="\r")
            print(f"Parsing input lines {i+1}/{len(lsplit)}... ", end="", flush=True)
        else:
            print(f"Parsing input lines {i+1}/{len(lsplit)}...       ", end="\r")
        
        time.sleep(0.005)
        

        if classify_line(split) == "st":
            curr_stime, curr_etime = parse_st_time(split)
        elif classify_line(split) == "ss" and is_weekly_meeting == False:
            emp_name = parse_ss_emp(split)
            name_list = emp_name.split(" ")
            first_name.append(name_list[0])
            last_name.append(name_list[1])

            if curr_stime == None or curr_etime == None:
                raise ValueError("Cannot create shift without time.")

        elif classify_line(split) == "sde":
            # add the last employee to the shift lift if there is one
            if curr_stime != None and curr_etime != None and len(first_name) != 0:
                for i in range(len(first_name)):
                    emp = Employee(first_name[i], last_name[i])
                    shifts.append(Shift(emp, curr_stime, curr_etime, weekday))
                # reset to prevent duplicates
                first_name = []
                last_name = []

            weekday += 1
        elif classify_line(split) == "meet":
            is_weekly_meeting = True
            if meeting["time"] == None:
                meeting["time"] = curr_stime
                meeting["day"] = weekday
            if len(first_name) > 0:
                first_name.pop()
                last_name.pop()
        elif classify_line(split) == "gl":
            if not is_weekly_meeting:
                # add the last employee to the shift list
                assert(curr_stime != None and curr_etime != None)
                assert(len(first_name) != 0)

                for i in range(len(first_name)):
                    emp = Employee(first_name[i], last_name[i])
                    shifts.append(Shift(emp, curr_stime, curr_etime, weekday))

                # reset to prevent duplicates
                first_name = []
                last_name = []
            
            is_weekly_meeting = False

    print("done")

    print("\nFormatting Excel file:")
    out = OutputWeek(shifts)
    out.gen_xl_file(output_file, meeting)

    final_time = round(time.time() - stime, 2)

    print(f"Complete. Finished file is '{output_file}.xlsx'. Took {final_time}s.")

if __name__ == "__main__":
    main()
