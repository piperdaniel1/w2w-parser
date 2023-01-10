from datatypes import *
from typing import List
import sys
import requests
import time
from datetime import date

def grab_text_from_w2w():
    url = ""
    try:
        with open(".hashed_req") as f:
            lines = f.readlines()
            url = lines[0].strip("\n")
            payload = lines[1].strip("\n")
    except FileNotFoundError:
        return None

    # contain payload in text/plain
    print("Sending request to " + url + " with payload " + payload)
    resp = requests.post(url, data=payload, headers={"Content-Type": "application/x-www-form-urlencoded"})

    # print(resp.text)
    # print(resp.headers)

    # print(resp.text)

    for i, elem in enumerate(resp.history):
        print("History " + str(i) + ": " + str(elem.headers))

    session = resp.history[-1].headers["Location"].split("=")[1]

    # print(resp.text)
    # print(resp.headers)
    # print(resp.history)
    # session = resp.history[0].headers["Location"].split("=")[1]
              #  "https://www5.whentowork.com/cgi-bin/w2wEE.dll/empfullschedule?SID=1592568667419D&lmi="
    full_sched = "https://www5.whentowork.com/cgi-bin/w2wEE.dll/empfullschedule?SID={sid}&lmi="

    print("Sending request with url: ", full_sched.format(sid=session))

    resp = requests.get(full_sched.format(sid=session),
            headers={"user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
                     "referer": f"https://www5.whentowork.com/cgi-bin/w2wEE.dll/home?SID={session}",
                     "cookie": "sd2=F99FC6D9EB75657F8DD412731B01B3233B8CC4B7; sd1=FDE683591CE86EDD8A2FCCC519FA217C4B0C7E87; _gcl_au=1.1.1595137208.1672171001; _ga=GA1.2.259010626.1672171002; _gid=GA1.2.742086556.1673302819",
                     "authority": f"www5.whentowork.com"})

    # print(resp.text)
    # print(resp.headers)

    main_text = resp.text.split("\n")

    for line in main_text:
        if "sdh(" in line:
            return line

    input("Press enter to show page...")
    
    for line in main_text:
        print(line)
    
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
    elif "sdb()" in line:
        return "sdb"
    
    return "err"

def main():
    if len(sys.argv) < 2:
        print("Usage: python parse.py <output_file> [--in <input_file>] [--meet <weekly meeting time>]")
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

        output_file = "Week " + str(weeks) + " Winter 2023 Schedule.xlsx"
        pass

    lsplit = []

    if "--in" in sys.argv:
        input_file = sys.argv[sys.argv.index("--in") + 1]
    
        try:
            with open(input_file) as f:
                line = f.readlines()[0]
                lsplit = line.split(";")
        except FileNotFoundError:
            print("Error: That input file was not found. Exiting...")
    else:
        line = grab_text_from_w2w()
        if line != None:
            lsplit = line.split(";")
        else:
            print("Error: Could not get line from when to work and no fallback input file provided. Exiting...")
            exit(1)

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

            shifts.append(Shift(emp, curr_stime, curr_etime, weekday))
        elif classify_line(split) == "sdb":
            weekday += 1

    out = OutputWeek(shifts)
    out.gen_xl_file(output_file)
    print("Complete.")

if __name__ == "__main__":
    main()
