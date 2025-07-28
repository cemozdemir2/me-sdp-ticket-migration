#!/usr/bin/env python3
"""
GUI application to import an Excel file into ManageEngine ServiceDesk MSP via V3 API
using TECHNICIAN_KEY for authentication.
Shows a progress window with a progress bar and live log output, logs to console and app.log.
Excludes any Excel fields with value 'Not Assigned' from the JSON payload, and converts
'Created Time' field to Unix timestamp (milliseconds). Adds:
- Completed Time → udf_sline_301
- Resolved Time → udf_sline_302
Disables SSL verification (verify=False) for all requests.
"""
import os
import sys
import logging
import json
import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from datetime import datetime
from typing import Dict, Any, List

# Disable SSL warnings
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Logging configuration (log file and level can still be set via env vars)
LOG_FILE = os.getenv("LOG_FILE", "app.log")
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
error_log_file = os.getenv("ERROR_LOG_FILE", "error.log")

# Setup logger
logger = logging.getLogger(__name__)
logger.setLevel(LOG_LEVEL)
formatter = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")

ch = logging.StreamHandler()
ch.setLevel(LOG_LEVEL)
ch.setFormatter(formatter)
logger.addHandler(ch)

fh = logging.FileHandler(LOG_FILE)
fh.setLevel(LOG_LEVEL)
fh.setFormatter(formatter)
logger.addHandler(fh)

# Custom handler to write logs into Tkinter Text widget
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)

# Utility: read Excel into list of dicts
def read_excel(fpath: str) -> List[Dict[str, Any]]:
    logger.info("Loading Excel file: %s", fpath)
    df = pd.read_excel(fpath)
    logger.info("Loaded %d rows", len(df))
    return df.to_dict(orient="records")

# Utility: parse Excel date string to unix milliseconds
def parse_date_to_ms(date_str: str) -> str:
    try:
        dt = datetime.strptime(date_str.strip(), "%d/%m/%Y %I:%M %p")
        ms = int(dt.timestamp() * 1000)
        return str(ms)
    except Exception as e:
        logger.error("Failed to parse date '%s': %s", date_str, e)
        return None

# Utility: build API request body, excluding 'Not Assigned', adding Completed/Resolved timestamps
def build_request_body(record: Dict[str, Any]) -> Dict[str, Any]:
    field_map = {
        'Requester': 'requester',
        'Site': 'site',
        'Account': 'account',
        'Subject': 'subject',
        'Item': 'item',
        'Request Type': 'request_type',
        'Level': 'level',
        'Urgency': 'urgency',
        'Impact': 'impact',
        'Technician': 'technician',
        'Category': 'category',
        'Subcategory': 'subcategory',
        'Priority': 'priority',
        'Group': 'group',
        'Status': 'status',
        'Template': 'template',
        'Description': 'description',
        'Resolution': 'resolution',
    }
    date_fields = {
        'Created Time': 'created_time'
    }
    body: Dict[str, Any] = {}
    udf: Dict[str, Any] = {}
    # Standard and lookup fields
    for col, api in field_map.items():
        val = record.get(col)
        if pd.notna(val):
            txt = str(val).strip()
            if txt == 'Not Assigned':
                continue
            # text fields vs lookup
            if api in ['subject', 'description']:
                body[api] = txt
            elif api in ['resolution']:
                body[api] = {'content': txt}
            else:
                body[api] = {'name': txt}
    # Created Time as date field
    for col, api in date_fields.items():
        val = record.get(col)
        if pd.notna(val):
            txt = str(val).strip()
            if txt and txt != 'Not Assigned':
                ms = parse_date_to_ms(txt)
                if ms:
                    body[api] = {'value': ms}
    # Completed Time → udf_sline_301
    comp = record.get('Completed Time')
    if pd.notna(comp):
        txt = str(comp).strip()
        if txt and txt != 'Not Assigned':
            ms = parse_date_to_ms(txt)
            if ms:
                udf['udf_sline_301'] = ms
    # Resolved Time → udf_sline_302
    res = record.get('Resolved Time')
    if pd.notna(res):
        txt = str(res).strip()
        if txt and txt != 'Not Assigned':
            ms = parse_date_to_ms(txt)
            if ms:
                udf['udf_sline_302'] = ms
    # UDF fields from columns
    for col, val in record.items():
        # skip timestamps already handled
        if col.startswith('udf_') and col not in ['udf_sline_301', 'udf_sline_302'] and pd.notna(val):
            txt = str(val).strip()
            if txt == 'Not Assigned':
                continue
            # For picklist UDFs, wrap value as {'name': ...}
            if col.startswith('udf_pick_'):
                udf[col] = {'name': txt}
            else:
                udf[col] = txt
    if udf:
        body['udf_fields'] = udf
    return body

# Main GUI application class
class ImporterApp:
    def __init__(self, master):
        self.master = master
        master.title("ServiceDesk MSP Importer")
        frame = tk.Frame(master, padx=10, pady=10)
        frame.pack(fill="x")
        tk.Label(frame, text="ServiceDesk Domain (e.g. example.manageengine.com:8080):").pack(anchor="w")
        self.domain_entry = tk.Entry(frame)
        self.domain_entry.pack(fill="x", pady=(0,5))
        tk.Label(frame, text="Technician Key:").pack(anchor="w")
        self.key_entry = tk.Entry(frame, show="*")
        self.key_entry.pack(fill="x", pady=(0,10))
        tk.Button(frame, text="Browse Excel & Start Import", command=self.load_file).pack()

    def load_file(self):
        domain = self.domain_entry.get().strip()
        key = self.key_entry.get().strip()
        if not domain or not key:
            messagebox.showerror("Missing Info", "Please enter both Domain and Technician Key.")
            return
        self.base_url = f"https://{domain}/api/v3/requests"
        self.tech_key = key
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")])
        if not path:
            return
        try:
            self.records = read_excel(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel:\n{e}")
            logger.exception("Excel load error")
            return
        self.start_import(path)

    def create_request(self, sess: requests.Session, body: Dict[str, Any]) -> None:
        payload = json.dumps({"request": body})
        data = {'input_data': payload}
        headers = {'authtoken': self.tech_key}
        logger.info("Request URL: %s", self.base_url)
        logger.info("Request Headers: %s", headers)
        logger.info("Request Payload: %s", payload)
        resp = sess.post(self.base_url, headers=headers, data=data, verify=False)
        logger.info("Response Status: %s", resp.status_code)
        logger.info("Response Body: %s", resp.text)
        if resp.status_code == 400:
            try:
                err = resp.json().get('response_status', {})
                for msg in err.get('messages', []):
                    if msg.get('message','').startswith('Site-Group-Technician'):
                        logger.warning("Group validation failed, retrying without group field")
                        # retry without group
                        new_body = {k:v for k,v in body.items() if k != 'group'}
                        new_payload = json.dumps({"request": new_body})
                        new_data = {'input_data': new_payload}
                        resp2 = sess.post(self.base_url, headers=headers, data=new_data, verify=False)
                        logger.info("Retry Status: %s", resp2.status_code)
                        logger.info("Retry Body: %s", resp2.text)
                        if resp2.status_code == 400:
                            entry = {
                                'timestamp': datetime.utcnow().isoformat()+'Z',
                                'request_payload': json.loads(new_payload),
                                'response_status': resp2.status_code,
                                'response_body': resp2.text
                            }
                            with open(error_log_file,'a') as ef:
                                ef.write(json.dumps(entry)+"\n")
                            logger.error("Logged retry error to %s", error_log_file)
                        resp2.raise_for_status()
                        return
            except Exception as ex:
                logger.error("Error handling group retry: %s", ex)
        resp.raise_for_status()

    def start_import(self, path: str):
        total = len(self.records)
        win = tk.Toplevel(self.master)
        win.title("Importing...")
        win.geometry("600x400")
        tk.Label(win, text=f"Importing {os.path.basename(path)} ({total} rows)").pack(pady=(10,0))
        pb = ttk.Progressbar(win, length=580, mode='determinate', maximum=total)
        pb.pack(pady=5)
        status = tk.Label(win, text=f"0 of {total}")
        status.pack()
        log_text = tk.Text(win, height=15, width=80)
        log_text.pack(pady=(5,0))
        scrollbar = tk.Scrollbar(win, command=log_text.yview)
        log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        handler = TextHandler(log_text)
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        sess = requests.Session()
        sess.trust_env = False
        sess.verify = False
        successes, failures = 0, []
        for i, rec in enumerate(self.records, start=1):
            logger.info("Processing row %d/%d", i, total)
            body = build_request_body(rec)
            try:
                self.create_request(sess, body)
                successes += 1
            except Exception as e:
                logger.error("Row %d failed: %s", i, e)
                failures.append(i)
            pb['value'] = i
            status.config(text=f"{i} of {total}")
            win.update()
        summary = f"Import complete: {successes}/{total} succeeded."
        if failures:
            summary += f" Failed rows: {failures}"
        logger.info(summary)
        messagebox.showinfo("Done", summary)
        win.destroy()


def main():
    root = tk.Tk()
    ImporterApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
