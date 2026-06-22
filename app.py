import os, sys, threading, shutil, math

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_parser import extract_text_from_pdf, parse
from excel_builder import build_detailed_excel, build_all_summaries_excel

BG = "#F4F6F9"
SURFACE = "#FFFFFF"
BORDER = "#E2E6EA"
BORDER2 = "#C8D0DA"
NAVY = "#0D2137"
NAVY2 = "#122840"
NAVY_LIGHT = "#1A3A5C"
BLUE = "#1565C0"
BLUE2 = "#1976D2"
BLUE_LIGHT = "#42A5F5"
BLUE_PALE = "#E3F2FD"
BLUE_CHIP = "#BBDEFB"

SUCCESS = "#2E7D32"
SUCCESS2 = "#388E3C"
SUCCESS_BG = "#E8F5E9"

WARNING_FG = "#E65100"
ERROR_FG = "#C62828"

TEXT1 = "#1A1A2E"
TEXT2 = "#455A64"
TEXT3 = "#90A4AE"
WHITE = "#FFFFFF"

DISABLED_BG = "#CFD8DC"
DISABLED_FG = "#90A4AE"

LOG_BG = "#1E2A35"
LOG_FG = "#CBD5E1"
LOG_OK = "#4ADE80"
LOG_ERR = "#F87171"
LOG_INFO = "#60A5FA"
LOG_SUB = "#94A3B8"


class VSS_App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF ⇄ Excel Converter")
        self.geometry("900x740")
        self.resizable(True, True)
        self.minsize(740, 620)
        self.configure(bg=BG)
        try:
            dpi = self.winfo_fpixels("1i")
            self.tk.call("tk", "scaling", dpi / 72)
        except Exception:
            pass

        self._pdf_paths: list[str] = []
        self._result_files: dict[str, str] = {}
        self._dl_btns: dict[str, tk.Button] = {}
        self._prog_anim_id = None
        self._prog_val = 0.0

        self._build_ui()

    def _build_ui(self):
        self._build_header()
        scroll_host = tk.Frame(self, bg=BG)
        scroll_host.pack(fill="both", expand=True)
        self._body = tk.Frame(scroll_host, bg=BG)
        self._body.pack(fill="both", expand=True, padx=24, pady=(18, 18))
        self._build_file_card()
        self._build_convert_btn()
        self._build_progress_card()
        self._build_download_card()
        self._build_log_card()

    def _build_header(self):
        hdr = tk.Canvas(self, height=76, bg=NAVY, highlightthickness=0)
        hdr.pack(fill="x")

        def _draw_hdr(evt=None):
            hdr.delete("all")
            w = hdr.winfo_width()
            if w <= 1:
                self.after(30, _draw_hdr)
                return
            h = 76
            for y in range(h):
                t = y / h
                r1 = int(13  + t * 10)
                g1 = int(33  + t * 20)
                b1 = int(55  + t * 30)
                col = f"#{r1:02x}{g1:02x}{b1:02x}"
                hdr.create_line(0, y, w, y, fill=col)
            hdr.create_rectangle(0, 73, w, 76, fill=BLUE, outline="")
    

            hdr.create_text(30, 29, anchor="w",
                            text="PDF ➜ Excel Converter",
                            fill=WHITE, font=("Segoe UI", 17, "bold"))
       

        hdr.bind("<Configure>", _draw_hdr)
        self.after(50, _draw_hdr)

    def _card(self, title=None, pady=(0, 12)):
        wrap = tk.Frame(self._body, bg=BG)
        wrap.pack(fill="x", pady=pady)

        for shade, pad in [("#C5CBD4", 3), ("#D4DAE2", 2), ("#E0E5EB", 1)]:
            tk.Frame(wrap, bg=shade, height=1).pack(fill="x", padx=pad)
        card = tk.Frame(wrap, bg=SURFACE,
                        highlightbackground=BORDER2, highlightthickness=1)
        card.pack(fill="x")

        if title:
            tk.Frame(card, bg=BLUE, height=3).pack(fill="x")   # top accent bar
            hrow = tk.Frame(card, bg="#F8FAFC")
            hrow.pack(fill="x")
            tk.Label(hrow, text=title, bg="#F8FAFC", fg=BLUE,
                     font=("Segoe UI", 8, "bold"),
                     padx=16, pady=7).pack(side="left")
            tk.Frame(card, bg=BORDER, height=1).pack(fill="x")
        return card

    def _build_file_card(self):
        card = self._card("📁 PDF FAYLNI TANLANG")

        inner = tk.Frame(card, bg=SURFACE)
        inner.pack(fill="x", padx=18, pady=14)
        dz = tk.Frame(inner, bg="#FAFBFD",
                      highlightbackground=BORDER2, highlightthickness=1)
        dz.pack(fill="x")

        left = tk.Frame(dz, bg="#FAFBFD")
        left.pack(side="left", fill="x", expand=True, padx=14, pady=12)
        self._lbl_files = tk.Label(
            left, text="Hech qanday fayl tanlanmagan",
            bg="#FAFBFD", fg=TEXT3,
            font=("Segoe UI", 10), anchor="w")
        self._lbl_files.pack(anchor="w")
        self._lbl_count = tk.Label(
            left, text="PDF fayllarni yuklash uchun tugmani bosing",
            bg="#FAFBFD", fg=TEXT3,
            font=("Segoe UI", 8), anchor="w")
        self._lbl_count.pack(anchor="w", pady=(2, 0))

        self._btn_pick = self._btn(
            dz, "📂 Fayl Tanlash", self._pick_files,
            bg=BLUE, fg=WHITE, hov=BLUE2,
            font=("Segoe UI", 10, "bold"), px=20, py=12)
        self._btn_pick.pack(side="right", padx=14, pady=12)

        self._file_frame = tk.Frame(card, bg=SURFACE)
        self._file_frame.pack(fill="x", padx=18, pady=(0, 6))

    def _build_convert_btn(self):
        wrap = tk.Frame(self._body, bg=BG)
        wrap.pack(fill="x", pady=(0, 12))
        self._btn_convert = self._btn(
            wrap, "  📄 Excel yaratish", self._start_convert,
            bg=BLUE, fg=WHITE, hov=BLUE2,
            font=("Segoe UI", 12, "bold"), px=0, py=14,
            state="disabled", dis=DISABLED_BG)
        self._btn_convert.pack(fill="x")

    def _build_progress_card(self):
        card = self._card(pady=(0, 12))

        inner = tk.Frame(card, bg=SURFACE)
        inner.pack(fill="x", padx=18, pady=(12, 14))

        top = tk.Frame(inner, bg=SURFACE)
        top.pack(fill="x")
        self._prog_lbl = tk.Label(
            top, text="Ready for conversion", bg=SURFACE, fg=TEXT2,
            font=("Segoe UI", 9))
        self._prog_lbl.pack(side="left")
        self._prog_pct = tk.Label(
            top, text="", bg=SURFACE, fg=BLUE,
            font=("Segoe UI", 9, "bold"))
        self._prog_pct.pack(side="right")

        pb_frame = tk.Frame(inner, bg=BORDER, height=10)
        pb_frame.pack(fill="x", pady=(8, 0))
        self._pb = tk.Canvas(pb_frame, height=10, bg="#DDE3EA",
                             highlightthickness=0, bd=0)
        self._pb.pack(fill="x")
        self._pb.bind("<Configure>", self._redraw_pb)

    def _redraw_pb(self, _=None):
        c = self._pb; c.delete("all")
        w = c.winfo_width() or 1
        c.create_rectangle(0, 0, w, 10, fill="#DDE3EA", outline="")
        fw = int(w * self._prog_val)
        if fw > 0:
            for x in range(fw):
                ratio = x / max(fw, 1)
                r = int(0x15 + (0x42 - 0x15) * ratio)
                g = int(0x65 + (0xA5 - 0x65) * ratio)
                b = int(0xC0 + (0xF5 - 0xC0) * ratio)
                c.create_line(x, 0, x, 10, fill=f"#{r:02x}{g:02x}{b:02x}")
        c.create_rectangle(0, 0, fw, 4, fill="", stipple="gray25",
                            outline="") if fw > 0 else None

    def _anim_pb(self):
        self._prog_val = min(self._prog_val + 0.009, 0.90)
        self._redraw_pb()
        self._prog_anim_id = self.after(55, self._anim_pb)

    def _stop_pb(self, ok=True):
        if self._prog_anim_id:
            self.after_cancel(self._prog_anim_id)
            self._prog_anim_id = None
        self._prog_val = 1.0 if ok else 0.0
        self._redraw_pb()

    def _build_download_card(self):
        card = self._card("🗂️NATIJA", pady=(0, 12))
        inner = tk.Frame(card, bg=SURFACE)
        inner.pack(fill="x", padx=18, pady=(10, 14))

        specs = [
            ("detailed", "⇄", "Converted",
             "Barcha PDF ma'lumotlari - to'liq hisobot"),
            ("all_summaries", "∑", "Summary",
             "·TURON HUMO USD ·TURON SUMMARY USD ·TURON HUMO UZS ·TURON SUMMARY UZS - 4 sheet"),
        ]
        for key, icon, title, sub in specs:
            row = tk.Frame(inner, bg=SURFACE,
                           highlightbackground=BORDER, highlightthickness=1)
            row.pack(fill="x", pady=(0, 8))

            band = tk.Frame(row, bg=BORDER2, width=5)
            band.pack(side="left", fill="y")

            info = tk.Frame(row, bg=SURFACE)
            info.pack(side="left", fill="x", expand=True, padx=14, pady=12)
            tk.Label(info, text=f"{icon}  {title}", bg=SURFACE, fg=TEXT1,
                     font=("Segoe UI", 10, "bold")).pack(anchor="w")
            tk.Label(info, text=sub, bg=SURFACE, fg=TEXT3,
                     font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

            btn = self._btn(
                row, "Yuklab olish 🡻", lambda k=key: self._download(k),
                bg=DISABLED_BG, fg=DISABLED_FG, hov=BORDER2,
                font=("Segoe UI", 9, "bold"), px=16, py=10,
                state="disabled")
            btn.pack(side="right", padx=14, pady=12)
            self._dl_btns[key] = btn
            self._dl_bands = getattr(self, "_dl_bands", {})
            self._dl_bands[key] = band

    def _build_log_card(self):
        wrap = tk.Frame(self._body, bg=BG)
        wrap.pack(fill="both", expand=True)

        for shade, pad in [("#B0B8C4", 3), ("#C5CBD4", 2)]:
            tk.Frame(wrap, bg=shade, height=1).pack(fill="x", padx=pad)

        outer = tk.Frame(wrap, bg=LOG_BG,
                         highlightbackground="#0A1520",
                         highlightthickness=1)
        outer.pack(fill="both", expand=True)

        tk.Frame(outer, bg=BLUE, height=3).pack(fill="x")

        hrow = tk.Frame(outer, bg="#162230")
        hrow.pack(fill="x")
        tk.Label(hrow, text="  ◉ LOG", bg="#162230", fg=LOG_INFO,
                 font=("Consolas", 8, "bold"), pady=6).pack(side="left")
        tk.Button(hrow, text="✕ Tozalash", bg="#162230", fg=LOG_SUB,
                  font=("Segoe UI", 8), bd=0, cursor="hand2",
                  activebackground="#162230", activeforeground=LOG_ERR,
                  command=self._log_clear).pack(side="right", padx=10)

        tk.Frame(outer, bg="#0A1520", height=1).pack(fill="x")

        inner = tk.Frame(outer, bg=LOG_BG)
        inner.pack(fill="both", expand=True)

        self._log = tk.Text(
            inner, bg=LOG_BG, fg=LOG_FG,
            font=("Consolas", 9), bd=0, relief="flat",
            state="disabled", wrap="word",
            insertbackground=BLUE_LIGHT,
            selectbackground=NAVY_LIGHT,
            padx=12, pady=8)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Dark.Vertical.TScrollbar",
                        background="#253545", troughcolor=LOG_BG,
                        arrowcolor=LOG_SUB, bordercolor=LOG_BG,
                        lightcolor=LOG_BG, darkcolor=LOG_BG)
        sb = ttk.Scrollbar(inner, style="Dark.Vertical.TScrollbar",
                           command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True)

        self._log.tag_configure("ok",   foreground=LOG_OK,
                                font=("Consolas", 9, "bold"))
        self._log.tag_configure("err",  foreground=LOG_ERR,
                                font=("Consolas", 9, "bold"))
        self._log.tag_configure("info", foreground=LOG_INFO)
        self._log.tag_configure("sub",  foreground=LOG_SUB)

    def _btn(self, parent, text, cmd,
             bg, fg, hov, font, px=14, py=8,
             state="normal", dis=DISABLED_BG):
        b = tk.Button(
            parent, text=text, command=cmd,
            bg=bg if state == "normal" else dis,
            fg=fg if state == "normal" else DISABLED_FG,
            activebackground=hov, activeforeground=fg,
            disabledforeground=DISABLED_FG,
            font=font, bd=0, padx=px, pady=py,
            cursor="hand2" if state == "normal" else "arrow",
            state=state, relief="flat")
        b._nbg = bg; b._hbg = hov; b._dis = dis; b._fg = fg
        b.bind("<Enter>", lambda e: b.config(bg=b._hbg)
               if b["state"] == "normal" else None)
        b.bind("<Leave>", lambda e: b.config(bg=b._nbg)
               if b["state"] == "normal" else None)
        return b

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title="PDF fayllarni tanlang",
            filetypes=[("PDF", "*.pdf"), ("Barcha", "*.*")])
        if not paths:
            return
        self._pdf_paths = list(paths)
        self._refresh_chips()
        self._result_files.clear()
        self._reset_dl_btns()
        self._btn_convert.config(state="normal", bg=self._btn_convert._nbg,
                                  fg=WHITE, cursor="hand2")
        self._log_clear()
        self._log_write(f"✔ {len(self._pdf_paths)} ta PDF tanlandi.\n", "ok")

    def _refresh_chips(self):
        for w in self._file_frame.winfo_children():
            w.destroy()
        names = [os.path.basename(p) for p in self._pdf_paths]
        n = len(names)
        if not names:
            self._lbl_files.config(text="Hech qanday fayl tanlanmagan", fg=TEXT3)
            self._lbl_count.config(
                text="PDF fayllarni yuklash uchun tugmani bosing", fg=TEXT3)
            return
        self._lbl_files.config(text=f"{n} ta PDF fayl tanlandi", fg=TEXT1)
        self._lbl_count.config(
            text="Fayllar ro'yxati:", fg=TEXT2)

        row = tk.Frame(self._file_frame, bg=SURFACE)
        row.pack(fill="x", pady=(6, 8))
        for nm in names:
            chip = tk.Frame(row, bg=BLUE_PALE,
                            highlightbackground=BLUE_CHIP,
                            highlightthickness=1)
            chip.pack(side="left", padx=(0, 6), pady=2)
            tk.Label(chip, text=f"📄 {nm}", bg=BLUE_PALE, fg=BLUE,
                     font=("Segoe UI", 8), padx=8, pady=5).pack()

    def _start_convert(self):
        if not self._pdf_paths:
            return
        self._btn_convert.config(state="disabled", bg=DISABLED_BG,
                                  fg=DISABLED_FG, cursor="arrow")
        self._reset_dl_btns()
        self._log_clear()
        self._prog_val = 0.0
        self._prog_lbl.config(text="The process is ongoing...", fg=WARNING_FG)
        self._prog_pct.config(text="0%")
        self._anim_pb()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        import tempfile
        tmp = tempfile.mkdtemp(prefix="vss_")
        try:
            all_rows = []
            for i, path in enumerate(self._pdf_paths):
                self._log_write(
                    f"PDF o'qilmoqda ({i+1}/{len(self._pdf_paths)}): "
                    f"{os.path.basename(path)}\n", "info")
                rows = parse(extract_text_from_pdf(path))
                all_rows.extend(rows)
                self._log_write(f"     ✓  {len(rows)} qator topildi\n", "sub")

            self._log_write("Umumiy Excel yozilmoqda...\n", "info")
            det = os.path.join(tmp, "umumiy.xlsx")
            build_detailed_excel(all_rows, det)
            self._result_files["detailed"] = det
            self._log_write("✓  Umumiy Excel tayyor\n", "ok")

            self._log_write("Barcha summalar yozilmoqda...\n", "info")
            sm = os.path.join(tmp, "barcha_summalar.xlsx")
            build_all_summaries_excel(all_rows, sm)
            self._result_files["all_summaries"] = sm
            self._log_write("✓ Barcha summalar tayyor\n", "ok")

            self._log_write("\n✨ Barcha fayllar tayyor!\n", "ok")
            self.after(0, self._done_ok)
        except Exception:
            import traceback; tb = traceback.format_exc()
            self._log_write(f"\n⚠️ Xatolik:\n{tb}\n", "err")
            self.after(0, self._done_err, tb.splitlines()[-1])

    def _done_ok(self):
        self._stop_pb(ok=True)
        self._prog_lbl.config(text="✓ Muvaffaqiyatli yakunlandi!", fg=SUCCESS)
        self._prog_pct.config(text="100%")
        self._btn_convert.config(state="normal", bg=self._btn_convert._nbg,
                                  fg=WHITE, cursor="hand2")
        labels = {
            "detailed": "⇄ Umumiy Excel",
            "all_summaries": "∑ Barcha Summalar",
        }
        for key, btn in self._dl_btns.items():
            if key in self._result_files:
                btn.config(state="normal", text=labels[key],
                           bg=SUCCESS, fg=WHITE, cursor="hand2")
                btn._nbg = SUCCESS; btn._hbg = SUCCESS2
                if hasattr(self, "_dl_bands"):
                    self._dl_bands[key].config(bg=SUCCESS)

    def _done_err(self, msg):
        self._stop_pb(ok=False)
        self._prog_lbl.config(text="⚠︎ Xatolik yuz berdi", fg=ERROR_FG)
        self._prog_pct.config(text="")
        self._btn_convert.config(state="normal", bg=self._btn_convert._nbg,
                                  fg=WHITE, cursor="hand2")
        messagebox.showerror("Xatolik", msg)

    def _reset_dl_btns(self):
        for key, btn in self._dl_btns.items():
            btn.config(state="disabled", text="Yuklab olish 🡻",
                       bg=DISABLED_BG, fg=DISABLED_FG, cursor="arrow")
            btn._nbg = DISABLED_BG; btn._hbg = BORDER2
            if hasattr(self, "_dl_bands"):
                self._dl_bands[key].config(bg=BORDER2)


    def _download(self, key):
        src = self._result_files.get(key)
        if not src or not os.path.exists(src):
            messagebox.showwarning("Ogohlantirish", "Fayl topilmadi.")
            return
        nm = {"detailed": "umumiy.xlsx", "all_summaries": "barcha_summalar.xlsx"}
        dest = filedialog.asksaveasfilename(
            title="Saqlash", initialfile=nm.get(key, "output.xlsx"),
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not dest:
            return
        shutil.copy2(src, dest)
        self._log_write(f"💾 Saqlandi: {dest}\n", "ok")
        messagebox.showinfo("Muvaffaqiyat", f"Fayl saqlandi:\n{dest}")

    def _log_write(self, msg, tag=None):
        def _do():
            self._log.config(state="normal")
            self._log.insert("end", msg, tag or "")
            self._log.see("end")
            self._log.config(state="disabled")
        self.after(0, _do)

    def _log_clear(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")


if __name__ == "__main__":
    app = VSS_App()
    app.mainloop()