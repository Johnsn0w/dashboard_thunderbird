import mailbox
from email.utils import parsedate_to_datetime as parse_to_dt
from bs4 import BeautifulSoup
from datetime import datetime as dt, timezone as tz, timedelta as td
from zoneinfo import ZoneInfo
from time import sleep
import tkinter as tk
from tkinter import font
from playsound3 import playsound
import pickle, os, sys

inbox_file = 'imap_map_inbox_sample.txt'
recent_visitors = {}

assert os.path.exists(inbox_file), f"assertion error, inbox file not found"

visitor_list_timeframe_in_hours = 100

msg_id_file = 'msg_ids.pkl'
if os.path.exists(msg_id_file): # load or create file tracking msg ids
    with open(msg_id_file, 'rb') as f:
        already_processed_msgs = pickle.load(f)
else:
    already_processed_msgs = set()
    
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
        if is_email_older_than_x_hours(email=email, hours=visitor_list_timeframe_in_hours): # skip
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
root.grid_propagate(False)

body_font  = font.Font(family="Arial", size=12)

title = tk.Label(root, text="Recent Arrivals", justify='center', font=body_font)
title.grid(row=0, column=0, padx=10, pady=10, sticky="n")


def resize_callback(event: tk.Event):
    widget = event.widget
    # filter out all widgets except root window
    if "." != widget.winfo_pathname(widget.winfo_id()): # skip
        return
    
    print("resizing event. path:", widget.winfo_pathname(widget.winfo_id()))
    
    body_font.configure(size=round(event.width * .1))
    

root.bind("<Configure>", resize_callback) 

def tkinter_main_loop():
    root.after(10, check_and_update_list)
    root.mainloop()

def check_and_update_list():
    print("Check for new msgs...")
    previous_visitors_dict = {**recent_visitors}
    update_recent_visitors_dict()
    i = 1
    if recent_visitors.keys() != previous_visitors_dict.keys(): # skip
        # visitors_list = '\n'.join([msg['visitor_name'] + " " + msg['received_dt'] for msg in recent_visitors.values()])
        print("change detected, updating list..")

        for i, msg in enumerate(recent_visitors.values()):
            i += 1
            visitor = tk.Label(
                root,
                text=msg['visitor_name'],
                justify='left',
                anchor='nw',
                background='grey',
                font=body_font
            )
            visitor.grid(row=i, column=0, padx=0, pady=0, sticky="nsew")

    root.after(1000, check_and_update_list)

tkinter_main_loop()





