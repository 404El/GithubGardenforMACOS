import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from canvasapi import Canvas
from datetime import datetime, timezone
import openpyxl
import urllib3
import re

# Silence Mac warnings
os.environ['TK_SILENCE_DEPRECATION'] = '1'
urllib3.disable_warnings(urllib3.exceptions.NotOpenSSLWarning)

class CanvasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Canvas Assignment Planner")
        self.root.geometry("500x600")
        self.root.configure(bg="#ffffff") # Clean white base

        # Force window to front
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(600, lambda: self.root.attributes("-topmost", False))

        # --- HEADER ---
        tk.Label(root, text="Canvas Planner", font=("Arial", 28, "bold"), bg="#ffffff", fg="#1d1d1f").pack(pady=(30, 20))

        # --- INPUT SECTION CONTAINER ---
        self.input_frame = tk.Frame(root, bg="#ffffff")
        self.input_frame.pack(padx=50, fill="x")

        # URL Input
        tk.Label(self.input_frame, text="1. Canvas URL", font=("Arial", 12, "bold"), bg="#ffffff", fg="#555555").pack(anchor="w")
        self.url_entry = tk.Entry(self.input_frame, font=("Arial", 14), 
                                  bg="#f1f1f1", fg="#000000", 
                                  insertbackground="black", relief="solid", bd=1)
        self.url_entry.insert(0, "https://vsu.instructure.com")
        self.url_entry.pack(pady=(5, 15), fill="x", ipady=8)

        # API Key Input
        tk.Label(self.input_frame, text="2. API Key", font=("Arial", 12, "bold"), bg="#ffffff", fg="#555555").pack(anchor="w")
        self.key_entry = tk.Entry(self.input_frame, font=("Arial", 14), 
                                  bg="#f1f1f1", fg="#000000", 
                                  show="*", insertbackground="black", relief="solid", bd=1)
        self.key_entry.pack(pady=(5, 15), fill="x", ipady=8)

        # --- SAVE PATH ---
        self.save_path = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        tk.Button(root, text="📂 Choose Save Folder", command=self.select_folder, 
                  font=("Arial", 11), fg="#007aff", bg="#ffffff", borderwidth=0).pack(pady=(10, 0))
        
        tk.Label(root, textvariable=self.save_path, font=("Arial", 10), bg="#ffffff", fg="gray").pack(pady=(0, 20))

        # --- GENERATE BUTTON ---
        # Note: Using a standard button to ensure Mac visibility
        self.btn = tk.Button(root, text="GENERATE DASHBOARD", 
                             command=self.run_export, 
                             bg="#007aff", fg="black", # 'black' text is safer for visibility on some Macs
                             font=("Arial", 14, "bold"),
                             padx=20, pady=15)
        self.btn.pack(pady=20)

        self.status_label = tk.Label(root, text="Ready", font=("Arial", 11), bg="#ffffff", fg="gray")
        self.status_label.pack()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder: self.save_path.set(folder)

    def run_export(self):
        url = self.url_entry.get().strip()
        key = self.key_entry.get().strip()
        if not key:
            messagebox.showerror("Error", "Please paste your API Key first!")
            return

        self.status_label.config(text="Syncing... Please wait.", fg="#007aff")
        self.root.update()

        try:
            canvas = Canvas(url, key)
            user = canvas.get_current_user()
            courses = user.get_courses(enrollment_state='active')
            
            data_rows = []
            now = datetime.now(timezone.utc)

            for course in courses:
                try:
                    c_name = getattr(course, 'name', 'Course')
                    if any(x in c_name for x in ["Orientation", "Testing"]): continue
                    
                    assignments = list(course.get_assignments())
                    for a in assignments:
                        due_at = getattr(a, 'due_at', None)
                        data_rows.append({
                            'Status': "✅ Done" if getattr(a, 'has_submitted_submissions', False) else "⏳ Pending",
                            'Course': c_name,
                            'Assignment': getattr(a, 'name', 'Unnamed'),
                            'Due Date': pd.to_datetime(due_at).strftime('%b %d') if due_at else "No Date"
                        })
                except: continue

            df = pd.DataFrame(data_rows)
            output_file = os.path.join(self.save_path.get(), "My_Canvas_Schedule.xlsx")
            df.to_excel(output_file, index=False)
            
            self.status_label.config(text="Success!", fg="green")
            os.system(f"open '{output_file}'")
        except Exception as e:
            messagebox.showerror("Error", f"Could not connect: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CanvasApp(root)
    root.mainloop()