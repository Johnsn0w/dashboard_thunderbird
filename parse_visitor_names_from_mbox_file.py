import mailbox
from email.utils import parsedate_to_datetime as parse_to_dt
from bs4 import BeautifulSoup
from datetime import datetime as dt, timezone as tz, timedelta as td
from zoneinfo import ZoneInfo
from time import sleep

inbox = mailbox.mbox('imap_map_inbox_sample.txt')

dt_str_format = '%H:%M - %d %b'

"""notifications last"""
def main():
    now = dt.now(tz.utc)
    for email in inbox:
        if email.is_multipart():  # skip
            # print("Multipart email detected.")
            continue
        if email['from'] != "noreply@vistab.co.nz": # skip
            continue     
        if is_email_older_than_x_hours(email=email, hours=30): # skip
            continue

        email_dt = parse_to_dt(email['date'])
        visitor_name = parse_visitor_name(email)
        nz_dt = utc_to_nz_dt(email_dt)

        formatted_arrival_info = f"{visitor_name} - {nz_dt.strftime('%H:%M %d %b')}"

        print(formatted_arrival_info)

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


main()
