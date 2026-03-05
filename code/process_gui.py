# Libraries
import os
import copy
import threading
import pandas as pd
import re
from datetime import datetime, timedelta
from itertools import compress
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ─────────────────────────────────────────────
#  GUI helpers
# ─────────────────────────────────────────────

class App(tk.Tk):
    """Single-window wizard that walks the user through all processing stages."""

    ACCENT  = "#011f5b"   # Penn Blue
    BG      = "#e1e4ec"   # Penn's lightest gray
    CARD    = "#ffffff"
    TEXT    = "#1c1c1c"
    SUBTEXT = "#5a5a5a"
    BORDER  = "#d4d4cc"
    RED     = "#c0392b"

    def __init__(self):
        super().__init__()
        self.title("LSEG Registration Processor")
        self.resizable(False, False)
        self.configure(bg=self.BG)

        # ── state ──────────────────────────────
        self.dat_ongoing_fname = tk.StringVar()
        self.dat_today_fnames  = []          # list of CSV paths
        self.dat_ongoing_snap  = None        # snapshot before user edits
        self.dat_ongoing_ref   = None        # live DataFrame (shared between stages)

        self._build_ui()
        self._show_stage(1)
        self.mainloop()

    # ── layout ────────────────────────────────

    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self, bg=self.ACCENT, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="LSEG Registration Processor",
                 bg=self.ACCENT, fg="white",
                 font=("Georgia", 15, "bold")).pack(side="left", padx=20, pady=0)
        self.stage_lbl = tk.Label(hdr, text="", bg=self.ACCENT, fg="#a8d5c2",
                                  font=("Courier New", 11))
        self.stage_lbl.pack(side="right", padx=20)

        # Content card
        self.card = tk.Frame(self, bg=self.CARD, bd=0,
                             highlightthickness=1,
                             highlightbackground=self.BORDER)
        self.card.pack(padx=30, pady=24, fill="both", expand=True)

        # Stage 1 – file selection
        self.frame_s1 = self._build_stage1()
        # Stage 2 – waiting for user edits
        self.frame_s2 = self._build_stage2()
        # Stage 3 – done
        self.frame_s3 = self._build_stage3()
        # Processing overlay
        self.frame_proc = self._build_processing()

    # ── Stage 1 ──────────────────────────────

    def _build_stage1(self):
        f = tk.Frame(self.card, bg=self.CARD)

        tk.Label(f, text="Step 1 — Select Input Files",
                 bg=self.CARD, fg=self.TEXT,
                 font=("Georgia", 13, "bold")).grid(row=0, column=0, columnspan=3,
                                                    sticky="w", padx=24, pady=(20, 4))
        tk.Label(f, text="Provide the ongoing accounts file and one or more "
                          "registration request CSVs for today.",
                 bg=self.CARD, fg=self.SUBTEXT,
                 font=("Helvetica", 10), wraplength=480).grid(
                     row=1, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 16))

        # AccountstoCheck
        tk.Label(f, text="Accounts to Check ongoing file (.xlsx)",
                 bg=self.CARD, fg=self.TEXT,
                 font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", padx=24)
        ongoing_entry = tk.Entry(f, textvariable=self.dat_ongoing_fname,
                                 width=42, state="readonly",
                                 readonlybackground="#eeeee8",
                                 relief="flat", bd=1,
                                 highlightthickness=1,
                                 highlightbackground=self.BORDER,
                                 font=("Courier New", 9))
        ongoing_entry.grid(row=3, column=0, columnspan=2, padx=(24, 6), pady=(4, 14), sticky="w")
        tk.Button(f, text="Browse…", command=self._browse_ongoing,
                  bg=self.ACCENT, fg="white", relief="flat",
                  font=("Helvetica", 9, "bold"),
                  activebackground="#155a3d", activeforeground="white",
                  cursor="hand2", padx=10, pady=4).grid(row=3, column=2, padx=(0, 24), sticky="w")

        # CSV list
        tk.Label(f, text="Registration Request file(s) from today (.csv)",
                 bg=self.CARD, fg=self.TEXT,
                 font=("Helvetica", 10, "bold")).grid(row=4, column=0, columnspan=3,
                                                      sticky="w", padx=24)

        list_frame = tk.Frame(f, bg=self.CARD)
        list_frame.grid(row=5, column=0, columnspan=3, padx=24, pady=(4, 0), sticky="ew")

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.csv_listbox = tk.Listbox(list_frame, height=5, width=54,
                                      yscrollcommand=scrollbar.set,
                                      selectmode="single",
                                      font=("Courier New", 9),
                                      bg="#eeeee8", relief="flat",
                                      bd=1, highlightthickness=1,
                                      highlightbackground=self.BORDER,
                                      activestyle="none")
        scrollbar.config(command=self.csv_listbox.yview)
        self.csv_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_row = tk.Frame(f, bg=self.CARD)
        btn_row.grid(row=6, column=0, columnspan=3, padx=24, pady=(6, 16), sticky="w")
        tk.Button(btn_row, text="+ Add CSV", command=self._add_csv,
                  bg=self.ACCENT, fg="white", relief="flat",
                  font=("Helvetica", 9, "bold"),
                  activebackground="#155a3d", activeforeground="white",
                  cursor="hand2", padx=10, pady=4).pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="✕ Remove Selected", command=self._remove_csv,
                  bg="#e8e8e2", fg=self.TEXT, relief="flat",
                  font=("Helvetica", 9),
                  cursor="hand2", padx=10, pady=4).pack(side="left")

        # Divider
        tk.Frame(f, bg=self.BORDER, height=1).grid(row=7, column=0, columnspan=3,
                                                    sticky="ew", padx=24, pady=(4, 12))

        self.s1_run_btn = tk.Button(f, text="Process Files →",
                                    command=self._run_stage1,
                                    bg=self.ACCENT, fg="white", relief="flat",
                                    font=("Helvetica", 11, "bold"),
                                    activebackground="#155a3d", activeforeground="white",
                                    cursor="hand2", padx=18, pady=8,
                                    state="disabled")
        self.s1_run_btn.grid(row=8, column=0, columnspan=3, padx=24, pady=(0, 20), sticky="e")

        f.columnconfigure(0, weight=1)
        return f

    # ── Stage 2 ──────────────────────────────

    def _build_stage2(self):
        f = tk.Frame(self.card, bg=self.CARD)

        tk.Label(f, text="Step 2 — Review & Update the Accounts File",
                 bg=self.CARD, fg=self.TEXT,
                 font=("Georgia", 13, "bold")).pack(anchor="w", padx=24, pady=(20, 4))

        self.s2_desc = tk.Label(
            f,
            text="",
            bg=self.CARD, fg=self.SUBTEXT,
            font=("Helvetica", 10), wraplength=480, justify="left")
        self.s2_desc.pack(anchor="w", padx=24, pady=(0, 14))

        # Checklist hint
        hint = tk.Frame(f, bg="#edf7f2", bd=0,
                        highlightthickness=1, highlightbackground="#aad6c0")
        hint.pack(fill="x", padx=24, pady=(0, 16))
        tk.Label(hint,
                 text="📋  What to do now:",
                 bg="#edf7f2", fg=self.ACCENT,
                 font=("Helvetica", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        for step in (
            "1. Open the AccountstoCheck file in Excel.",
            "2. Look up new accounts and fill in the required fields.",
            "3. Review flagged accounts and remove them as needed.",
            "4. Save and close the file.",
            "5. Return here and click Continue.",
        ):
            tk.Label(hint, text=step, bg="#edf7f2", fg=self.TEXT,
                     font=("Helvetica", 10)).pack(anchor="w", padx=14)
        tk.Label(hint, text="", bg="#edf7f2").pack(pady=4)

        self.s2_warn = tk.Label(f, text="", bg=self.CARD, fg=self.RED,
                                font=("Helvetica", 9, "italic"))
        self.s2_warn.pack(anchor="w", padx=24)

        tk.Frame(f, bg=self.BORDER, height=1).pack(fill="x", padx=24, pady=(8, 12))

        tk.Button(f, text="Continue →",
                  command=self._run_stage2,
                  bg=self.ACCENT, fg="white", relief="flat",
                  font=("Helvetica", 11, "bold"),
                  activebackground="#155a3d", activeforeground="white",
                  cursor="hand2", padx=18, pady=8).pack(anchor="e", padx=24, pady=(0, 20))

        return f

    # ── Stage 3 ──────────────────────────────

    def _build_stage3(self):
        f = tk.Frame(self.card, bg=self.CARD)

        tk.Label(f, text="✔  Processing Complete",
                 bg=self.CARD, fg=self.ACCENT,
                 font=("Georgia", 14, "bold")).pack(anchor="w", padx=24, pady=(20, 4))

        self.s3_desc = tk.Label(
            f, text="",
            bg=self.CARD, fg=self.SUBTEXT,
            font=("Helvetica", 10), wraplength=480, justify="left")
        self.s3_desc.pack(anchor="w", padx=24, pady=(0, 16))

        notice = tk.Frame(f, bg="#edf7f2", bd=0,
                          highlightthickness=1, highlightbackground="#aad6c0")
        notice.pack(fill="x", padx=24, pady=(0, 20))
        tk.Label(notice,
                 text="The Accounts to Check file has been updated.\n"
                      "Recommended actions are noted in the Take_Action column.",
                 bg="#edf7f2", fg=self.TEXT,
                 font=("Helvetica", 10), justify="left").pack(padx=14, pady=12, anchor="w")

        tk.Frame(f, bg=self.BORDER, height=1).pack(fill="x", padx=24, pady=(0, 12))
        tk.Button(f, text="Close", command=self.destroy,
                  bg="#e8e8e2", fg=self.TEXT, relief="flat",
                  font=("Helvetica", 11, "bold"),
                  cursor="hand2", padx=18, pady=8).pack(anchor="e", padx=24, pady=(0, 20))

        return f

    # ── Processing overlay ────────────────────

    def _build_processing(self):
        f = tk.Frame(self.card, bg=self.CARD)
        tk.Label(f, text="", bg=self.CARD).pack(pady=30)
        self.spinner_lbl = tk.Label(f, text="⏳", bg=self.CARD,
                                    font=("Helvetica", 28))
        self.spinner_lbl.pack()
        self.proc_msg = tk.Label(f, text="Processing files, please wait…",
                                 bg=self.CARD, fg=self.SUBTEXT,
                                 font=("Helvetica", 11))
        self.proc_msg.pack(pady=(10, 30))
        return f

    # ── Navigation ────────────────────────────

    def _show_stage(self, n, msg=""):
        for fr in (self.frame_s1, self.frame_s2, self.frame_s3, self.frame_proc):
            fr.pack_forget()
        labels = {1: "Stage 1 of 3", 2: "Stage 2 of 3", 3: "Stage 3 of 3", 0: "Processing…"}
        self.stage_lbl.config(text=labels.get(n, ""))
        if n == 1:
            self.frame_s1.pack(fill="both", expand=True)
        elif n == 2:
            self.s2_desc.config(
                text=f"The AccountstoCheck file has been updated with today's new "
                     f"registration records.\n\nFile: {os.path.basename(self.dat_ongoing_fname.get())}")
            self.frame_s2.pack(fill="both", expand=True)
        elif n == 3:
            self.s3_desc.config(
                text=f"All records have been processed.\n\n"
                     f"File: {os.path.basename(self.dat_ongoing_fname.get())}")
            self.frame_s3.pack(fill="both", expand=True)
        elif n == 0:
            self.proc_msg.config(text=msg or "Processing files, please wait…")
            self.frame_proc.pack(fill="both", expand=True)
        self.update_idletasks()

    # ── File browser helpers ──────────────────

    def _browse_ongoing(self):
        path = filedialog.askopenfilename(
            title="Select Accounts to Check ongoing file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.dat_ongoing_fname.set(path)
            self._update_run_button()

    def _add_csv(self):
        paths = filedialog.askopenfilenames(
            title="Select Registration Request file(s) from today",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        for p in paths:
            if p not in self.dat_today_fnames:
                self.dat_today_fnames.append(p)
                self.csv_listbox.insert("end", os.path.basename(p))
        self._update_run_button()

    def _remove_csv(self):
        sel = self.csv_listbox.curselection()
        if sel:
            idx = sel[0]
            self.csv_listbox.delete(idx)
            self.dat_today_fnames.pop(idx)
        self._update_run_button()

    def _update_run_button(self):
        ready = self.dat_ongoing_fname.get() and len(self.dat_today_fnames) > 0
        self.s1_run_btn.config(state="normal" if ready else "disabled")

    # ── Stage 1 processing ────────────────────

    def _run_stage1(self):
        self._show_stage(0, "Processing registration request file(s). Please wait…")
        threading.Thread(target=self._stage1_worker, daemon=True).start()

    def _stage1_worker(self):
        try:
            dat_ongoing_fname = self.dat_ongoing_fname.get()
            dat_today_fname   = list(self.dat_today_fnames)   # copy

            dat_ongoing = pd.read_excel(dat_ongoing_fname, na_values=[], keep_default_na=False)
            file_date   = re.search(r'_(.*)\.', dat_today_fname[0]).group(1)
            dat_today   = pd.read_csv(dat_today_fname.pop(), na_values=[], keep_default_na=False)
            while len(dat_today_fname) > 0:
                dat_today = pd.concat(
                    [pd.read_csv(dat_today_fname.pop(), na_values=[], keep_default_na=False), dat_today],
                    sort=False, ignore_index=True).fillna("")

            dat_today["New_Record"] = file_date

            # Remove full duplicate records
            dat_today.drop(dat_today[dat_today.duplicated(keep="first")].index, inplace=True)
            dat_today.reset_index(inplace=True, drop=True)

            dat_today.drop(
                dat_today[pd.merge(dat_today, dat_ongoing,
                                   on=list(dat_today.columns),
                                   how="left", indicator=True).loc[:, "_merge"] == "both"].index,
                inplace=True)
            dat_today.reset_index(inplace=True, drop=True)

            # Flag potential issues
            dat_today["Email_local-part"] = dat_today["COMPANY EMAIL"].str.extract(r"(.+?(?=\@))")
            dat_today["New_Warning"] = ""

            mask = dat_today[dat_today["COMPANY EMAIL"].str.contains(
                r"^[a-zA-Z]+\.[a-zA-Z]+\.w[a-zA-Z]\d\d@[a-zA-Z]*\.*upenn\.edu$", na=False)].index
            dat_today.loc[mask, "New_Warning"] += "Alumni-style email. "

            mask = dat_today[dat_today.loc[:, ["FIRST NAME", "LAST NAME"]].duplicated(keep=False)].index
            dat_today.loc[mask, "New_Warning"] += f"Repeated name in {file_date} file. "

            mask = dat_today[pd.merge(
                dat_today,
                dat_ongoing.drop_duplicates(subset=["FIRST NAME", "LAST NAME"]),
                on=["FIRST NAME", "LAST NAME"], how="left", indicator=True
            ).loc[:, "_merge"] == "both"].index
            dat_today.loc[mask, "New_Warning"] += f"Name from {file_date} exists in old file. "

            mask = dat_today[dat_today.loc[:, ["Email_local-part"]].duplicated(keep=False)].index
            dat_today.loc[mask, "New_Warning"] += f"Repeated email prefix in {file_date} file. "

            mask = dat_today[pd.merge(
                dat_today,
                dat_ongoing.drop_duplicates(subset=["Email_local-part"]),
                on=["Email_local-part"], how="left", indicator=True
            ).loc[:, "_merge"] == "both"].index
            dat_today.loc[mask, "New_Warning"] += f"Email prefix from {file_date} exists in old file. "

            # Merge and save
            dat_today.sort_values(by=["LAST NAME", "FIRST NAME", "COMPANY EMAIL"],
                                  inplace=True, ignore_index=True)
            dat_ongoing = pd.concat([dat_ongoing, dat_today],
                                    sort=False, ignore_index=True).fillna("")
            dat_ongoing.to_excel(dat_ongoing_fname, sheet_name="in", index=False)

            # Snapshot for change-detection in stage 2
            self.dat_ongoing_ref  = dat_ongoing
            self.dat_ongoing_snap = copy.deepcopy(dat_ongoing)

            self.after(0, lambda: self._show_stage(2))

        except Exception as e:
            self.after(0, lambda exc=e: self._on_error(exc))

    # ── Stage 2 processing ────────────────────

    def _run_stage2(self):
        self.s2_warn.config(text="")
        dat_ongoing_fname = self.dat_ongoing_fname.get()

        # Check if file was actually changed since stage 1
        try:
            dat_reloaded = pd.read_excel(dat_ongoing_fname, na_values=[], keep_default_na=False)
        except Exception as e:
            self._on_error(e)
            return

        if dat_reloaded.equals(self.dat_ongoing_snap):
            self.s2_warn.config(
                text="⚠  The file does not appear to have been updated yet.\n"
                     "   Please fill in the required fields, save the file, then click Continue.")
            return

        self._show_stage(0, "Assessing actions, please wait…")
        threading.Thread(target=self._stage2_worker,
                         args=(dat_reloaded,), daemon=True).start()

    def _stage2_worker(self, dat_ongoing):
        try:
            dat_ongoing_fname = self.dat_ongoing_fname.get()
            mask_iterate = list(dat_ongoing.index)

            # Send follow-up
            mask = dat_ongoing[
                pd.to_datetime(dat_ongoing.loc[:, "Followup_DueDate"],
                               format="%Y-%m-%d", errors="coerce") <= datetime.today()].index
            dat_ongoing.loc[mask, "Email_Text"]  = "Optional placeholder text for multiple accounts email template."
            dat_ongoing.loc[mask, "Take_Action"] = dat_ongoing.loc[mask, "Take_Action"] + "Send follow-up email. "
            dat_ongoing.loc[mask, "Followup_DueDate"] = ""

            mask = dat_ongoing[dat_ongoing.loc[:, "Followup_DueDate"] == ""].index
            mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

            # Fix patron label
            mask = dat_ongoing[dat_ongoing.loc[:, "New_Record"] != ""].index
            mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

            for i in mask_iterate:
                gy    = dat_ongoing.loc[i, "Graduation Year"]
                notes = dat_ongoing.loc[i, "Notes"]
                label = dat_ongoing.loc[i, "LABEL"]
                if gy == "ALUM" and label != "Alumni":
                    dat_ongoing.loc[i, "LABEL"]       = "Alumni"
                    dat_ongoing.loc[i, "Take_Action"] += "Change LSEG Label to Alumni. "
                elif gy == "N/A" and bool(re.search(r"Staff", notes)) and label != "Staff":
                    dat_ongoing.loc[i, "LABEL"]       = "Staff"
                    dat_ongoing.loc[i, "Take_Action"] += "Change LSEG Label to Staff. "
                elif (isinstance(gy, int) or gy == "N/A") and \
                     (bool(re.search(r"PhD", notes)) or bool(re.search(r"Faculty", notes))) and \
                     label != "Faculty/PhD":
                    dat_ongoing.loc[i, "LABEL"]       = "Faculty/PhD"
                    dat_ongoing.loc[i, "Take_Action"] += "Change LSEG Label to Faculty/PhD. "
                elif (isinstance(gy, int) or gy == "Unknown") and \
                     not bool(re.search(r"PhD", notes)) and label != "Student":
                    dat_ongoing.loc[i, "LABEL"]       = "Student"
                    dat_ongoing.loc[i, "Take_Action"] += "Change LSEG Label to Student. "

            # Remove licenses (Alumni)
            mask  = dat_ongoing[dat_ongoing.loc[:, "Has Licenses in backend"] == "Yes"].index
            mask2 = dat_ongoing[dat_ongoing.loc[:, "LABEL"] == "Alumni"].index
            mask  = list(compress(mask2, [i in set(mask) for i in mask2]))
            mask  = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
            dat_ongoing.loc[mask, "Take_Action"] += "Unassign all licenses. "
            dat_ongoing.loc[mask, "New_Record"]   = ""

            mask = dat_ongoing[dat_ongoing.loc[:, "LABEL"] != "Alumni"].index
            mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

            # Create account
            mask = sorted(list(set(
                list(dat_ongoing[pd.to_numeric(dat_ongoing["Appears in backend"],
                                               errors="coerce") <= 0].index) +
                list(dat_ongoing[dat_ongoing.loc[:, "Appears in backend"] == "No"].index))))
            mask = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
            dat_ongoing.loc[mask, "Email_Text"]  = "Placeholder text: Your account is ready."
            dat_ongoing.loc[mask, "Take_Action"] += "Create an LSEG account and notify by email. "
            dat_ongoing.loc[mask, "New_Record"]   = ""

            mask = dat_ongoing[dat_ongoing.loc[:, "New_Record"] != ""].index
            mask_iterate = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))

            # De-duplicate accounts
            mask = dat_ongoing[pd.to_numeric(dat_ongoing["Appears in backend"],
                                             errors="coerce") > 1].index
            mask = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
            dat_ongoing.loc[mask, "Followup_DueDate"] = (datetime.today() + timedelta(days=7)).date()
            dat_ongoing.loc[mask, "Email_Text"]  = (
                f"Placeholder text: You created accounts under {dat_ongoing.loc[mask, 'COMPANY EMAIL']} and __. Which would you prefer to keep?")
            dat_ongoing.loc[mask, "Take_Action"] += "Patron has multiple LSEG accounts; ask which to keep. "
            dat_ongoing.loc[mask, "New_Record"]   = ""

            # Add licenses
            mask  = dat_ongoing[dat_ongoing.loc[:, "Has Licenses in backend"] == "No"].index
            mask2 = dat_ongoing[dat_ongoing.loc[:, "LABEL"] != "Alumni"].index
            mask  = list(compress(mask2, [i in set(mask) for i in mask2]))
            mask  = list(compress(mask_iterate, [i in set(mask) for i in mask_iterate]))
            dat_ongoing.loc[mask, "Take_Action"] += "Assign licenses. "

            # Clear new-record flags and save
            dat_ongoing.loc[:, "New_Record"] = ""
            dat_ongoing.sort_values(by=["Take_Action"], inplace=True, ignore_index=True)
            dat_ongoing.to_excel(dat_ongoing_fname, sheet_name="in", index=False)

            self.after(0, lambda: self._show_stage(3))

        except Exception as e:
            self.after(0, lambda exc=e: self._on_error(exc))

    # ── Error handler ─────────────────────────

    def _on_error(self, exc):
        self._show_stage(1)
        messagebox.showerror("Processing Error",
                             f"An error occurred:\n\n{exc}\n\n"
                             "Please check your input files and try again.")


# ─────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────

if __name__ == "__main__":
    App()
