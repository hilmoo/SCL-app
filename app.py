import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog
import webbrowser

import pandas as pd

import matplotlib.pyplot as plt
from io import BytesIO
from PIL import Image

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

class App:
    def __init__(self, root):
        # setting title
        root.title("SCL 90")
        # setting window size
        width = 400
        height = 400
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = "%dx%d+%d+%d" % (
            width,
            height,
            (screenwidth - width) / 2,
            (screenheight - height) / 2,
        )
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_142 = tk.Label(root)
        ft = tkFont.Font(family="Times", size=12)
        GLabel_142["font"] = ft
        GLabel_142["fg"] = "#333333"
        GLabel_142["justify"] = "center"
        GLabel_142["text"] = "Output Directory"
        GLabel_142.place(x=0,y=120,width=400,height=30)

        GLabel_882 = tk.Label(root)
        ft = tkFont.Font(family="Times", size=12)
        GLabel_882["font"] = ft
        GLabel_882["fg"] = "#333333"
        GLabel_882["justify"] = "center"
        GLabel_882["text"] = "Spreadsheet form"
        GLabel_882.place(x=0,y=30,width=400,height=25)
        
        GLabel_526=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_526["font"] = ft
        GLabel_526["fg"] = "#333333"
        GLabel_526["justify"] = "center"
        GLabel_526["text"] = "Made by hilmoo"
        GLabel_526.place(x=0,y=320,width=400,height=25)

        GLabel_39=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_39["font"] = ft
        GLabel_39["fg"] = "#333333"
        GLabel_39["justify"] = "center"
        GLabel_39["text"] = "https://github.com/hilmoo"
        GLabel_39.place(x=0,y=350,width=400,height=25)
        GLabel_39.bind("<Button-1>", self.my_github)
        GLabel_39.bind("<Enter>", lambda e: GLabel_39.config(cursor="hand2"))

        ####################################

        GLineEdit_230 = tk.Entry(root)
        GLineEdit_230["borderwidth"] = "1px"
        ft = tkFont.Font(family="Times", size=10)
        GLineEdit_230["font"] = ft
        GLineEdit_230["fg"] = "#333333"
        GLineEdit_230["justify"] = "center"
        GLineEdit_230["text"] = "id spreadsheet"  ##################
        GLineEdit_230.place(x=50,y=70,width=220,height=30)

        GLineEdit_303 = tk.Entry(root)
        GLineEdit_303["borderwidth"] = "1px"
        ft = tkFont.Font(family="Times", size=10)
        GLineEdit_303["font"] = ft
        GLineEdit_303["fg"] = "#333333"
        GLineEdit_303["justify"] = "center"
        GLineEdit_303["text"] = "Spreadsheet Column"  ##################
        GLineEdit_303.place(x=300,y=70,width=50,height=30)

        GLineEdit_720 = tk.Entry(root)
        GLineEdit_720["borderwidth"] = "1px"
        ft = tkFont.Font(family="Times", size=10)
        GLineEdit_720["font"] = ft
        GLineEdit_720["fg"] = "#333333"
        GLineEdit_720["justify"] = "center"
        GLineEdit_720["text"] = "Output Path"  ##################
        GLineEdit_720.place(x=50,y=160,width=200,height=30)

        GButton_469 = tk.Button(root)
        GButton_469["bg"] = "#f0f0f0"
        ft = tkFont.Font(family="Times", size=10)
        GButton_469["font"] = ft
        GButton_469["fg"] = "#000000"
        GButton_469["justify"] = "center"
        GButton_469["text"] = "Browse"  # Output
        GButton_469.place(x=280,y=160,width=70,height=30)
        GButton_469["command"] = lambda: self.GButton_469_command(GLineEdit_720)

        GButton_252 = tk.Button(root)
        GButton_252["anchor"] = "center"
        GButton_252["bg"] = "#f0f0f0"
        ft = tkFont.Font(family="Times", size=10)
        GButton_252["font"] = ft
        GButton_252["fg"] = "#000000"
        GButton_252["justify"] = "center"
        GButton_252["text"] = "Run"  ##################
        GButton_252.place(x=130,y=240,width=150,height=50)
        GButton_252["command"] = lambda: self.GButton_252_command(
            GLineEdit_230, GLineEdit_303, GLineEdit_720
        )

    def GButton_469_command(self, entry_widget):  # browse output | gline 720
        output_path = filedialog.askdirectory()
        if output_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, output_path)

    def GButton_252_command(
        self, csv_JAWABAN, Col, output_path
    ):  # submit button
        sheet_id = csv_JAWABAN.get()
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        dj = pd.read_csv(url)
        Col = int(Col.get())

        csv_TSCORE = "https://raw.githubusercontent.com/hilmoo/SCL-app/main/Tscore.csv"
        DATA_DIRI = dj.iloc[Col, 0:7]
        JAWABAN = dj.iloc[Col, 7:98]

        DATA_GRAPH = self.DataMaker(
            JAWABAN, csv_TSCORE
        )

        IMG_GRAPH = self.GraphMaker(DATA_GRAPH)

        self.PdfMaker(DATA_DIRI, IMG_GRAPH, DATA_GRAPH, output_path.get())

    def DataMaker(self, JAWABAN, csv_TSCORE):
        df = pd.read_csv(csv_TSCORE)
        KUNCI = [
            5,
            2,
            3,
            5,
            1,
            6,
            9,
            8,
            3,
            3,
            7,
            5,
            4,
            1,
            1,
            9,
            2,
            8,
            10,
            1,
            6,
            1,
            2,
            7,
            4,
            1,
            5,
            3,
            1,
            1,
            1,
            1,
            2,
            6,
            9,
            6,
            6,
            3,
            2,
            5,
            6,
            5,
            8,
            10,
            3,
            3,
            4,
            5,
            5,
            4,
            3,
            5,
            5,
            1,
            3,
            5,
            2,
            5,
            10,
            10,
            6,
            9,
            7,
            10,
            3,
            10,
            7,
            8,
            6,
            4,
            1,
            2,
            6,
            7,
            4,
            8,
            9,
            2,
            1,
            2,
            7,
            4,
            8,
            9,
            9,
            2,
            9,
            9,
            10,
            9,
        ]

        class Raw:
            DEPRESI = (
                ANSIETAS
            ) = (
                OBSESIS_K
            ) = (
                PHOBIA
            ) = (
                SOMATISASI
            ) = SENSITIV_I = HOSTILITAS = PARANOID = PSIKOTIK = ADDITIONAL = 0
            for i in range(90):
                if KUNCI[i] == 1:
                    DEPRESI += JAWABAN.iloc[i]
                if KUNCI[i] == 2:
                    ANSIETAS += JAWABAN.iloc[i]
                if KUNCI[i] == 3:
                    OBSESIS_K += JAWABAN.iloc[i]
                if KUNCI[i] == 4:
                    PHOBIA += JAWABAN.iloc[i]
                if KUNCI[i] == 5:
                    SOMATISASI += JAWABAN.iloc[i]
                if KUNCI[i] == 6:
                    SENSITIV_I += JAWABAN.iloc[i]
                if KUNCI[i] == 7:
                    HOSTILITAS += JAWABAN.iloc[i]
                if KUNCI[i] == 8:
                    PARANOID += JAWABAN.iloc[i]
                if KUNCI[i] == 9:
                    PSIKOTIK += JAWABAN.iloc[i]
                if KUNCI[i] == 10:
                    ADDITIONAL += JAWABAN.iloc[i]

            TOTAL = (
                DEPRESI
                + ANSIETAS
                + OBSESIS_K
                + PHOBIA
                + SOMATISASI
                + SENSITIV_I
                + HOSTILITAS
                + PARANOID
                + PSIKOTIK
                + ADDITIONAL
            )

        class Graph:
            TOTAL = df.iloc[Raw.TOTAL, 1]
            DEPRESI = df.iloc[Raw.DEPRESI, 2]
            ANSIETAS = df.iloc[Raw.ANSIETAS, 3]
            OBSESIS_K = df.iloc[Raw.OBSESIS_K, 4]
            PHOBIA = df.iloc[Raw.PHOBIA, 5]
            SOMATISASI = df.iloc[Raw.SOMATISASI, 6]
            SENSITIV_I = df.iloc[Raw.SENSITIV_I, 7]
            HOSTILITAS = df.iloc[Raw.HOSTILITAS, 8]
            PARANOID = df.iloc[Raw.PARANOID, 9]
            PSIKOTIK = df.iloc[Raw.PSIKOTIK, 10]
            ADDITIONAL = df.iloc[Raw.ADDITIONAL, 11]

            AAA = [
                TOTAL,
                DEPRESI,
                ANSIETAS,
                OBSESIS_K,
                PHOBIA,
                SOMATISASI,
                SENSITIV_I,
                HOSTILITAS,
                PARANOID,
                PSIKOTIK,
                ADDITIONAL,
            ]

        return Graph.AAA

    def GraphMaker(self, data):
        x_labels = [
            "TOTAL",
            "DEPRESI",
            "ANSIETAS",
            "OBSESIS_K",
            "PHOBIA",
            "SOMATISASI",
            "SENSITIV_I",
            "HOSTILITAS",
            "PARANOID",
            "PSIKOTIK",
            "ADDITIONAL",
        ]
        
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(x_labels, data, marker="o", color="blue", label="Data Points")
        ax.axhline(y=61, color="red", linestyle="--", label="Normal (61)")
        ax.set_ylim(0, 160)
        ax.set_ylabel("T-Score")
        ax.set_title("Tscore Graph")
        ax.legend()

        buffer = BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)

        return buffer

    def PdfMaker(self, DATA_DIRI, IMG_GRAPH, DATA_GRAPH, output_directory):
        tanggal_buat = DATA_DIRI.iloc[0]
        nama = DATA_DIRI.iloc[1]
        tempat_lhr = DATA_DIRI.iloc[2]
        tgl_lhr = DATA_DIRI.iloc[3]
        pendidikan = DATA_DIRI.iloc[4]
        status = DATA_DIRI.iloc[5]
        pekerjaan = DATA_DIRI.iloc[6]
        
        output_filename = f"{tanggal_buat[:10].replace('/', '_')}_{nama}.pdf"
        output_path = os.path.join(output_directory, output_filename)
        
        pdf_canvas = canvas.Canvas(output_path, pagesize=letter)

        pdf_canvas.setFont("Helvetica-Bold", 16)
        pdf_canvas.drawCentredString(300, 750, "SCL 90")

        pdf_canvas.setFont("Helvetica", 12)
        pdf_canvas.drawString(50, 700, f"Tanggal")
        pdf_canvas.drawString(150, 700, f": {tanggal_buat}")
        pdf_canvas.drawString(50, 680, f"Nama")
        pdf_canvas.drawString(150, 680, f": {nama}")
        pdf_canvas.drawString(50, 660, f"TTL")
        pdf_canvas.drawString(150, 660, f": {tempat_lhr}, {tgl_lhr}")
        pdf_canvas.drawString(50, 640, f"Pendidikan")
        pdf_canvas.drawString(150, 640, f": {pendidikan}")
        pdf_canvas.drawString(50, 620, f"Status")
        pdf_canvas.drawString(150, 620, f": {status}")
        pdf_canvas.drawString(50, 600, f"Pekerjaan")
        pdf_canvas.drawString(150, 600, f": {pekerjaan}")
        
        IMAGE1 = Image.open(IMG_GRAPH)
        original_width, original_height = IMAGE1.size
        pdf_canvas.drawInlineImage(
            IMAGE1,
            50, 350,
            width=original_width * 0.4,
            height=original_height * 0.4
        )
        
        TOTAL, DEPRESI, ANSIETAS, OBSESIS_K, PHOBIA, SOMATISASI, SENSITIV_I, HOSTILITAS, PARANOID, PSIKOTIK, ADDITIONAL = DATA_GRAPH[:11]
        pdf_canvas.setFont("Helvetica", 12)
        pdf_canvas.drawString(50, 330, f"Keterangan : ")
        pdf_canvas.drawString(50, 310, f"Total")
        pdf_canvas.drawString(150, 310, f": {TOTAL}")
        pdf_canvas.drawString(50, 290, f"Depresi")
        pdf_canvas.drawString(150, 290, f": {DEPRESI}")
        pdf_canvas.drawString(50, 270, f"Ansietas")
        pdf_canvas.drawString(150, 270, f": {ANSIETAS}")
        pdf_canvas.drawString(50, 250, f"Obsesi K")
        pdf_canvas.drawString(150, 250, f": {OBSESIS_K}")
        pdf_canvas.drawString(50, 230, f"Phobia")
        pdf_canvas.drawString(150, 230, f": {PHOBIA}")
        pdf_canvas.drawString(50, 210, f"Somatisasi")
        pdf_canvas.drawString(150, 210, f": {SOMATISASI}")
        pdf_canvas.drawString(50, 190, f"Sensitiv I")
        pdf_canvas.drawString(150, 190, f": {SENSITIV_I}")
        pdf_canvas.drawString(50, 170, f"Hostilitas")
        pdf_canvas.drawString(150, 170, f": {HOSTILITAS}")
        pdf_canvas.drawString(50, 150, f"Paranoid")
        pdf_canvas.drawString(150, 150, f": {PARANOID}")
        pdf_canvas.drawString(50, 130, f"Psikotik")
        pdf_canvas.drawString(150, 130, f": {PSIKOTIK}")
        pdf_canvas.drawString(50, 110, f"Additional")
        pdf_canvas.drawString(150, 110, f": {ADDITIONAL}")
        
        pdf_canvas.save()

    def my_github(self, event):
        webbrowser.open("https://github.com/hilmoo")
    
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
