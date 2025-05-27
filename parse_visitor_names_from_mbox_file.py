import mailbox
from email.utils import parsedate_to_datetime as parse_to_dt
from bs4 import BeautifulSoup
from datetime import datetime as dt, timezone as tz, timedelta as td
from zoneinfo import ZoneInfo
from time import sleep
import tkinter as tk
from tkinter import font
from playsound3 import playsound
import pickle, os

msg_id_file = 'msg_ids.pkl'
if os.path.exists(msg_id_file):
    with open(msg_id_file, 'rb') as f:
        already_processed_msgs = pickle.load(f)
else:
    already_processed_msgs = set()
    
inbox_file = 'imap_map_inbox_sample.txt'

recent_visitors = {}

dt_str_format = '%H:%M - %d %b'

def update_recent_visitors_dict():
    recent_visitors.clear()
    inbox = mailbox.mbox(inbox_file)
    now = dt.now(tz.utc)
    for email in inbox:
        if email.is_multipart():  # skip
            # print("Multipart email detected.")
            continue
        if email['from'] != "noreply@vistab.co.nz": # skip
            continue     
        if is_email_older_than_x_hours(email=email, hours=33): # skip
            continue
        if email['Message-ID'] in recent_visitors: # skip
            break
        
        
        msg_id = email['Message-ID']

        email_dt = parse_to_dt(email['date'])
        visitor_name = parse_visitor_name(email)
        nz_dt = utc_to_nz_dt(email_dt)

        msg_id = email['Message-ID']

        recent_visitors[msg_id] = {"visitor_name": visitor_name, "received_dt": nz_dt.strftime('%H:%M')}

        if msg_id not in already_processed_msgs:
            already_processed_msgs.add(msg_id)
            with open(msg_id_file, 'wb') as f:
                pickle.dump(already_processed_msgs, f)
            playsound("notification.mp3")


def is_email_older_than_x_hours(*,email, hours):
    email_dt = parse_to_dt(email['date'])
    now = dt.now(tz.utc)
    return now - email_dt > td(hours=hours)

def utc_to_nz_dt(utc_dt):
    return utc_dt.replace(tzinfo=tz.utc).astimezone(ZoneInfo("Pacific/Auckland"))

def parse_visitor_name(email):
    body = email.get_payload(decode=True).decode('utf-8')
    soup = BeautifulSoup(body, 'html.parser')

    # Find the visitor name based on text content
    text_lines = soup.get_text(separator="\n")
    for line in text_lines.splitlines():
        if 'Visitor Name:' in line:
            visitor_name = line.split('Visitor Name:')[1].strip()
            break
    else:
        visitor_name = None  # or raise an error/log a warning
    return visitor_name


root = tk.Tk()
root.title("")
title = tk.Label(root, text="Recent Arrivals", justify='center', font=("Arial", 18))
title.pack(padx=10, pady=10)
visitor_names_label = tk.Label(root, text="_", justify='left')
visitor_names_label.pack(pady=20)
def tkinter_main_loop():
    root.after(10, update_list)
    root.mainloop()

def update_list():
    print("Updating list...")
    update_recent_visitors_dict()
    visitor_names = ""
    for msg_id, msg in recent_visitors.items():
        visitor_names += (
            f"{msg["visitor_name"]}"
            f" "
            f"{msg["received_dt"]}"
            "\n"
            )

    visitor_names_label.config(text=visitor_names)
    root.after(1000, update_list)  # Re-run every 1000 ms (1 second)


"""
record-like object
    has properties:
        datetime
        visitor_name
        message-id
"""

tkinter_main_loop()


