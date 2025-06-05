import mailbox
from email.utils import parsedate_to_datetime as parse_to_dt
from bs4 import BeautifulSoup
from datetime import datetime as dt, timezone as tz, timedelta as td
from zoneinfo import ZoneInfo
from time import sleep
from playsound3 import playsound
import pickle, os, sys, subprocess, json
from pathlib import Path
import pandas as pd
import math

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import font


def main():
    app = Application()
    app.mainloop()


class Application(tk.Tk): # instead of creating root instance, we are the root instance

    def __init__(self):
        super().__init__() # instantiate tkinter instance
        self["bg"] = "#d5d2d2"
        self.saved_pos_file = Path("./temp/saved_pos.pkl")
        self.bind('<Control-r>', lambda event: self.reload_window() )
        # self.bind("<Configure>", lambda event: [self.save_window_geometry(self.geometry())])
        self.bind("<Control-c>", lambda _: sys.exit())
        self.geometry(self.load_saved_position())
        self.title(" ")
        self.iconbitmap("blank.ico")

        self.v_frame = VisitorsFrame(self)
        self.v_frame.grid(row=0, column=0, sticky="nesw", padx=10, pady=10)
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure   (0,weight=1)
        
        # self.v_frame2 = VisitorsFrame(self)
        # self.v_frame2.grid(row=0, column=1, sticky="nesw", padx=10, pady=10)
        # self.grid_columnconfigure(1,weight=1)
        # self.grid_rowconfigure   (0,weight=1)
    
    def reload_window(self):
        # some of te extra logic here is just to prevent vs-code closing all subprocesses after root subprocess is closed.
        self.processes = [] #
        python = sys.executable
        script = os.path.abspath(__file__)
        proc = subprocess.Popen([python, script])
        self.processes.append(proc) #
        self.destroy()

        for p in self.processes: #
            p.wait() #

    def save_window_geometry(self, _geometry: str):
        with open(self.saved_pos_file, 'wb') as f:
            pickle.dump(_geometry, f)

    def load_saved_position(self):
        if os.path.exists(self.saved_pos_file): # load or create file tracking msg ids
            with open(self.saved_pos_file, 'rb') as f:
                saved_pos = pickle.load(f)
        else:
            saved_pos = "+0+0"
        return saved_pos

class VisitorsFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.visitor_list_timeframe_in_hours = 48
        self.font_geometry_size_ratio = .05
        self.header_font_size = 1.2
        self.body_font_size   = 1
        
        style = ttk.Style()
        style.configure("VisitorsFrame.TFrame", background=parent["bg"])
        self.configure(style="VisitorsFrame.TFrame")

        self.bind("<Configure>", lambda event: self.resize_callback(event))
        self.visitors_body_font  = font.Font(family="Courier New", size=12, weight="bold")
        self.visitors_title_font = font.Font(family="Courier New", size=12, weight="bold")
        
        style.configure("body.TLabel", 
                        padding=3,
                        foreground="black",
                        font=self.visitors_body_font,
                        relief="flat",
                        background=parent["bg"]
                        )
        style.configure("title.TLabel", 
                        padding=3,
                        foreground="black",
                        font=self.visitors_title_font,
                        relief="flat",
                        background=parent["bg"]
                        )


        self.recent_visitors = {}
        title = ttk.Label(self, text="Recent Arrivals", justify='center', style="title.TLabel")
        title.grid(row=0, column=0, padx=10, pady=10, sticky="n", columnspan=2)

        self.msg_id_file = Path('./temp/msg_ids.pkl')
        if os.path.exists(self.msg_id_file): # load or create file tracking msg ids
            with open(self.msg_id_file, 'rb') as f:
                self.already_processed_msgs = pickle.load(f)
        else:
            self.already_processed_msgs = set()

        self.after(10, self.check_and_update_list)

    def update_recent_visitors_dict(self):
        inbox_file = 'imap_map_inbox_sample.txt'
        assert os.path.exists(inbox_file), f"assertion error, inbox file not found"
        play_notification = False
        self.recent_visitors.clear()
        inbox = mailbox.mbox(inbox_file)
        now = dt.now(tz.utc)
        for email in inbox:
            if email.is_multipart():  # skip
                continue
            if email['from'] != "noreply@vistab.co.nz": # skip
                continue
            if self.is_email_older_than_x_hours(email=email, hours=self.visitor_list_timeframe_in_hours): # skip
                continue
            if email['Message-ID'] in self.recent_visitors: # skip
                break

            msg_id = email['Message-ID']

            email_dt = parse_to_dt(email['date'])
            visitor_name = self.parse_visitor_name(email)
            nz_dt = self.utc_to_nz_dt(email_dt)
            nz_dt = nz_dt.strftime('%H:%M')
            nz_dt = nz_dt if nz_dt[0] != "0" else nz_dt[1:]

            msg_id = email['Message-ID']

            self.recent_visitors[msg_id] = {"visitor_name": visitor_name, "timestamp": nz_dt}

            if msg_id not in self.already_processed_msgs:
                play_notification = True
                self.already_processed_msgs.add(msg_id)
                with open(self.msg_id_file, 'wb') as f:
                    pickle.dump(self.already_processed_msgs, f)

        if play_notification:
            playsound("notification.mp3")

    def check_and_update_list(self):
        # print("Checking for new msgs...")
        previous_visitors_dict = {**self.recent_visitors}
        self.update_recent_visitors_dict()
        i = 0
        if self.recent_visitors.keys() != previous_visitors_dict.keys():
            print("change detected, updating list..")
            
            for widget in self.winfo_children()[1:]:
                widget.destroy()

            for i, msg in enumerate(self.recent_visitors.values()):
                i += 1
                visitor = ttk.Label(
                    self,
                    text=msg['visitor_name'],
                    style="body.TLabel"
                    )
                timestamp = ttk.Label(
                    self,
                    text=msg['timestamp'],
                    style="body.TLabel",
                )

                visitor.grid  (row=i, column=0, padx=0, pady=0, sticky="nsew")
                timestamp.grid(row=i, column=1, padx=0, pady=0, sticky="e")
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.after(1000, self.check_and_update_list)

    def is_email_older_than_x_hours(self, email, hours):
        email_dt = parse_to_dt(email['date'])
        now = dt.now(tz.utc)
        return now - email_dt > td(hours=hours)

    def parse_visitor_name(self, email):
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

    def utc_to_nz_dt(self, utc_dt) -> str:
        return utc_dt.replace(tzinfo=tz.utc).astimezone(ZoneInfo("Pacific/Auckland"))
    
    def resize_callback(self, event: tk.Event):
        h = round(event.height * self.font_geometry_size_ratio)
        w = round(event.width  * self.font_geometry_size_ratio)
        print(f"{h=}")
        print(f"{w=}")
        # print(f"{}")
        font_size = min(h, w)

        title = round(font_size * self.header_font_size)
        body  = round(font_size * self.body_font_size)

        self.visitors_body_font .configure(size=body)
        self.visitors_title_font.configure(size=title)

class FeedbackFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.feedback_filepath = './feedback_stats.csv'
        self.feedback_df = self.load_csv()
        self.inbox_path = 'imap_map_inbox_sample.txt'

        self.update_data_from_inbox()
        self.feedback_stats = self.process_data_to_stats()
    
    def load_csv(self):
        if os.path.exists(self.feedback_filepath): # load or create file tracking msg ids
            feedback_stats_df = pd.read_csv(self.feedback_filepath)
        else:
            feedback_stats_df = pd.DataFrame(columns=['timestamp', 'satisfaction', 'learning_score_before', 'learning_score_after', 'message_id'])
        return feedback_stats_df

    def update_data_from_inbox(self) -> list | None:
        inbox = mailbox.mbox(self.inbox_path)
        df = None
        for email in inbox:
            if email.is_multipart():
                continue
            if "---powerautomate---" not in email["subject"]:
                continue
            if self.feedback_df['message_id'].isin([email['Message-ID']]).any():
                continue
            
            payload = email.get_payload()
            soup = BeautifulSoup(payload, 'html.parser')
            email_body = soup.get_text()
            
            data = email_body.split("---data_start---")[1]
            data = data.split("---data_end---")[0]
            data = json.loads(data)
            data['message_id'] = email['Message-ID']
            
            self.feedback_df.loc[len(self.feedback_df)] = data
            
            self.save_df_to_disk()

        print(self.feedback_df)
            
    def save_df_to_disk(self):
        self.feedback_df.to_csv(self.feedback_filepath)

    def process_data_to_stats(self):
        df = self.feedback_df

        # region ##### process_df_to_specified_date_range ################
        today_start = dt.today().replace(hour=0)
        today_end = dt.today().replace(hour=23)
        

        selected_date_range = "Last 7 days"
        date_ranges = {
            "Last 7 days": {
                "start": today_start - td(days=7),
                "end": today_end
            },
            "Last 30 days": {
                "start": today_start - td(days=30),
                "end": today_end
            },
            "Last 365 days": {
                "start": today_start - td(days=365),
                "end": today_end
            },
            "This week": {
                "start": today_start - td(days=today_start.weekday()),
                "end": today_end
            },
            "This month": {
                "start": today_start.replace(day=1),
                "end": today_end
            },
            "This year": {
                "start": today_start.replace(month=1, day=1),
                "end": today_end
            },
            "Previous week": {
                "start": today_start - td(days=today_start.weekday() + 7),
                "end": today_end - td(days=today_end.weekday() + 1)
            },
            "Previous month": {
                "start": (today_start.replace(day=1) - td(days=1)).replace(day=1),
                "end": today_end.replace(day=1) - td(days=1)
            },
            "Previous year": {
                "start": today_start.replace(month=1, day=1, year=today_start.year - 1),
                "end": today_end.replace(month=1, day=1) - td(days=1)
            }
        }

        # Convert Timestamp column to datetime
        df['datetime'] = pd.to_datetime(df['timestamp'], format='%Y-%m-%dT%H:%M:%S.%f')
        # Define date range
        start_date  = date_ranges[selected_date_range]['start']
        end_date    = date_ranges[selected_date_range]['end']
        print(f"Date Range: {selected_date_range}\nStart:      {start_date}\nEnd:        {end_date}")

        # Filter DataFrame based on selected date range
        df = df[(df['datetime'] >= start_date) & (df['datetime'] <= end_date)]

        # endregion ##### process_df_to_specified_date_range ###############

        ability_before_col = 'learning_score_before'
        df[ability_before_col] = pd.to_numeric(
            df[ability_before_col], errors='coerce')
        avg_ability_before = df[ability_before_col].mean()

        ability_after_col = 'learning_score_after'
        df[ability_after_col] = pd.to_numeric(
            df[ability_after_col], errors='coerce')
        avg_ability_after = df[ability_after_col].mean()

        learning_improvement = round(
            (avg_ability_after - avg_ability_before) * 10, 0)
        learning_improvement = learning_improvement if not math.isnan(learning_improvement) else 0

        # region get average satisfaction score as string `avg_satisfaction`
        satisfaction_col = 'satisfaction'
        df[satisfaction_col] = pd.to_numeric(df[satisfaction_col], errors='coerce')
        avg_satisfaction = str(round(df[satisfaction_col].mean(), 1))
        # print(f'avg satisfaction sore: {avg_satisfaction}')
        # endregion

        print(f"{avg_satisfaction=}")
        print(f"{float(avg_ability_before)=}")
        print(f"{float(avg_ability_after)=}")
        print(f"{float(learning_improvement)=}")


        record_count = str(df.shape[0])

        processed_data = {
        "avg_satisfaction": avg_satisfaction,
        "record_count": record_count,
        "learning_improvement": learning_improvement,
        "selected_date_range": selected_date_range
        }

        return processed_data            





main()
sys.exit()





