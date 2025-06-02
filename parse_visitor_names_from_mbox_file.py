import mailbox
from email.utils import parsedate_to_datetime as parse_to_dt
from bs4 import BeautifulSoup
from datetime import datetime as dt, timezone as tz, timedelta as td
from zoneinfo import ZoneInfo
from time import sleep
from tkinter import font
from playsound3 import playsound
import pickle, os, sys, subprocess

import tkinter as tk
from pathlib import Path
import tkinter.ttk as ttk

saved_pos_file = Path("./temp/saved_pos.pkl")
msg_id_file = Path('./temp/msg_ids.pkl')
inbox_file = 'imap_map_inbox_sample.txt'
recent_visitors = {}

visitor_list_timeframe_in_hours = 200
font_geometry_size_ratio = .05
header_font_size = 1.2
body_font_size   = 1

if os.path.exists(msg_id_file): # load or create file tracking msg ids
    with open(msg_id_file, 'rb') as f:
        already_processed_msgs = pickle.load(f)
else:
    already_processed_msgs = set()

assert os.path.exists(inbox_file), f"assertion error, inbox file not found"

def reload_window(*, event=None, root):
    # some of te extra logic here is just to prevent vs-code closing all subprocesses after root subprocess is closed.
    processes = [] #
    python = sys.executable
    script = os.path.abspath(__file__)
    proc = subprocess.Popen([python, script])
    processes.append(proc) #
    root.destroy()

    for p in processes: #
        p.wait() #

def save_window_geometry(_geometry: str):
    with open(saved_pos_file, 'wb') as f:
        pickle.dump(_geometry, f)

def load_saved_position():
    global saved_pos
    if os.path.exists(saved_pos_file): # load or create file tracking msg ids
        with open(saved_pos_file, 'rb') as f:
            saved_pos = pickle.load(f)
    else:
        print("trace")
        saved_pos = "+0+0"
    return saved_pos

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
        nz_dt = nz_dt.strftime('%H:%M')
        nz_dt = nz_dt if nz_dt[0] != "0" else nz_dt[1:]

        msg_id = email['Message-ID']

        recent_visitors[msg_id] = {"visitor_name": visitor_name, "timestamp": nz_dt}

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

def utc_to_nz_dt(utc_dt) -> str:
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
root.title(" ")
# root.overrideredirect(True) 
# root.wm_attributes('-fullscreen', 'true') # fullscreen
root["bg"] = "#d5d2d2"
root.grid_propagate(False) # unsure on functionality
root.iconbitmap("blank.ico")

visitors_frame = tk.Frame(root, background=root["bg"])
visitors_frame.pack(anchor="center", padx=50, fill="x")

visitors_body_font  = font.Font(family="Courier New", size=12, weight="bold")
visitors_title_font = font.Font(family="Courier New", size=12, weight="bold")

style = ttk.Style()
style.configure("body.TLabel", 
                padding=3,
                foreground="black",
                font=visitors_body_font,
                relief="flat",
                background=root["bg"]
                )
style.configure("title.TLabel", 
                padding=3,
                foreground="black",
                font=visitors_title_font,
                relief="flat",
                background=root["bg"]
                )


title = ttk.Label(visitors_frame, text="Recent Arrivals", justify='center', style="title.TLabel")
title.grid(row=0, column=0, padx=10, pady=10, sticky="n", columnspan=2)


root.bind('<Control-r>', lambda event: reload_window(root=root) )
root.bind("<Configure>", lambda event: [resize_callback(event), save_window_geometry(root.geometry())])
root.bind("<Control-c>", lambda _: sys.exit())

root.geometry(load_saved_position())


def resize_callback(event: tk.Event):
    widget = event.widget

    if "." != widget.winfo_pathname(widget.winfo_id()): # skip non-root widgets
        return

    h = round(event.height * font_geometry_size_ratio)
    w = round(event.width  * font_geometry_size_ratio)
    font_size = min(h, w)

    title = round(font_size * header_font_size)
    body  = round(font_size * body_font_size)

    visitors_body_font  .configure(size=body)
    visitors_title_font.configure(size=title)

def tkinter_main_loop():
    root.after(10, check_and_update_list)
    root.mainloop()

def check_and_update_list():
    # print("Checking for new msgs...")
    previous_visitors_dict = {**recent_visitors}
    update_recent_visitors_dict()
    i = 0
    if recent_visitors.keys() != previous_visitors_dict.keys():
        print("change detected, updating list..")

        for widget in visitors_frame.winfo_children()[1:]:
            widget.destroy()


        for i, msg in enumerate(recent_visitors.values()):
            i += 1
            visitor = ttk.Label(
                visitors_frame,
                text=msg['visitor_name'],
                style="body.TLabel"
                )
            timestamp = ttk.Label(
                visitors_frame,
                text=msg['timestamp'],
                style="body.TLabel",
            )

            visitor.grid  (row=i, column=0, padx=0, pady=0, sticky="nsew")
            timestamp.grid(row=i, column=1, padx=0, pady=0, sticky="e")
    visitors_frame.grid_columnconfigure(0, weight=1)
    visitors_frame.grid_columnconfigure(1, weight=1)
    root.after(1000, check_and_update_list)

tkinter_main_loop()





