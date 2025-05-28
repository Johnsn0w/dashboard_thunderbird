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
    play_notification = False
    recent_visitors.clear()
    inbox = mailbox.mbox(inbox_file)
    now = dt.now(tz.utc)
    for email in inbox:
        if email.is_multipart():  # skip
            # print("Multipart email detected.")
            continue
        if email['from'] != "noreply@vistab.co.nz": # skip
            continue     
        if is_email_older_than_x_hours(email=email, hours=60): # skip
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
            play_notification = True
            already_processed_msgs.add(msg_id)
            with open(msg_id_file, 'wb') as f:
                pickle.dump(already_processed_msgs, f)
    
    if play_notification:
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
root.title("Recent Arrivals")

# Configure the root window's grid to allow resizing
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

body_font  = font.Font(family="Arial", size=12)

title = tk.Label(root, text="Recent Arrivals", justify='center', font=body_font)
title.grid(row=0, column=0, padx=10, pady=10, sticky="n")

visitor_names_label = tk.Label(
    root,
    text="_",
    justify='left',
    anchor='nw',
    background='grey',
    font=body_font
)
visitor_names_label.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

def resize_font(event):
    # Determine the smaller dimension
    limiting_dimension = min(event.width, event.height)
    # Calculate new font size based on that dimension
    new_size = max(8, int(limiting_dimension / 10))  # Adjust the divisor as needed
    body_font.configure(size=new_size)


visitor_names_label.bind("<Configure>", resize_font)

def tkinter_main_loop():
    root.after(10, update_list)
    root.mainloop()

def update_list():
    print("Updating list...")
    update_recent_visitors_dict()
    visitors_list = '\n'.join([msg['visitor_name'] + " " + msg['received_dt'] for msg in recent_visitors.values()])
    visitor_names_label.config(text=visitors_list)
    root.after(1000, update_list)

tkinter_main_loop()





