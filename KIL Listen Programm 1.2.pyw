from tkinter import *
import os.path, time
from datetime import datetime
import os
import glob
import pandas as pd

Listen_IDs = []

userhome = os.path.expanduser('~')
desktop = userhome + '/Desktop'

#Pfad von Programmpfad
# Mit Dateinamen
script = os.path.realpath(__file__)
# Ohne Dateiname

def pfad_von_KSTID_ausgeben(Suchbegriff):
    for file in glob.glob(os.path.dirname(script) + "/*" + Suchbegriff + "*"):
        return desktop + "\\KILerstellen\\" + os.path.basename(file)
     
# Pfad von KST ID Tabelle ausgeben
pfad_von_KSTID_ausgeben(".XLS")

# Excel als DF einlesen mit Excel
df = pd.read_excel(pfad_von_KSTID_ausgeben(".XLS"))

#dataframes mit den gefilterten KST IDs
Köln = df.loc[
    (df["LF_ID"] == 2) & (df["KST_ID"] != 37)]
Bayreuth = df.loc[
    (df["LF_ID"] == 3) & (df["KST_ID"] != 37)]
Berlin = df.loc[
    (df["LF_ID"] == 4) & (df["KST_ID"] != 37)]
Hamburg = df.loc[
    (df["LF_ID"] == 5)& (df["KST_ID"] != 37)]
Riedstadt = df.loc[
    (df["LF_ID"] == 6) & (df["KST_ID"] != 37)]
Ulm = df.loc[
    (df["LF_ID"] == 7) & (df["KST_ID"] != 37)]
Kempten = df.loc[
    (df["LF_ID"] == 9) & (df["KST_ID"] != 37)]
Hildesheim = df.loc[
    (df["LF_ID"] == 13) & (df["KST_ID"] != 37)]
Dresden = df.loc[
    (df["LF_ID"] == 14) & (df["KST_ID"] != 37)]
Bremen = df.loc[
    (df["LF_ID"] == 15) & (df["KST_ID"] != 37)]
Halle = df.loc[
    (df["LF_ID"] == 17) & (df["KST_ID"] != 37)]
Eichenau = df.loc[
    (df["LF_ID"] == 36) & (df["KST_ID"] != 37)]
Rostock = df.loc[
    (df["LF_ID"] == 44) & (df["KST_ID"] != 37)]
CF_Gastro = df.loc[
    (df["LF_ID"] == 19) & (df["KST_ID"] != 37)]

def Kst_zur_Liste(df, Liste):
    for Kst_Id in df["KST_ID"]: 
        Liste.append(Kst_Id)

Lager_Kostenstellen_ID_Bayreuth = []
Lager_Kostenstellen_ID_Berlin = []
Lager_Kostenstellen_ID_Bremen = []
Lager_Kostenstellen_ID_Dresden = []
Lager_Kostenstellen_ID_Eichenau = []
Lager_Kostenstellen_ID_Halle = []
Lager_Kostenstellen_ID_Hamburg = []
Lager_Kostenstellen_ID_Hildesheim = []
Lager_Kostenstellen_ID_Kempten = []
Lager_Kostenstellen_ID_Köln = []
Lager_Kostenstellen_ID_Riedstadt = []
Lager_Kostenstellen_ID_Rostock = []
Lager_Kostenstellen_ID_Ulm = []
Zensiert_OE_Zuordnung = []

Kst_zur_Liste(Bayreuth, Lager_Kostenstellen_ID_Bayreuth)
Kst_zur_Liste(Berlin, Lager_Kostenstellen_ID_Berlin)
Kst_zur_Liste(Bremen, Lager_Kostenstellen_ID_Bremen)
Kst_zur_Liste(Köln, Lager_Kostenstellen_ID_Köln)
Kst_zur_Liste(Dresden, Lager_Kostenstellen_ID_Dresden)
Kst_zur_Liste(Hamburg, Lager_Kostenstellen_ID_Hamburg)
Kst_zur_Liste(Riedstadt, Lager_Kostenstellen_ID_Riedstadt)
Kst_zur_Liste(Ulm, Lager_Kostenstellen_ID_Ulm)
Kst_zur_Liste(Kempten, Lager_Kostenstellen_ID_Kempten)
Kst_zur_Liste(Hildesheim, Lager_Kostenstellen_ID_Hildesheim)
Kst_zur_Liste(Halle, Lager_Kostenstellen_ID_Halle)
Kst_zur_Liste(Eichenau, Lager_Kostenstellen_ID_Eichenau)
Kst_zur_Liste(Rostock, Lager_Kostenstellen_ID_Rostock)
Kst_zur_Liste(Zensiert, Zensiert)

def Liefertage_erstellen(aa, bb, cc, dd, ee, ff, gg, hh, ii, jj, kk, ll, mm, nn):
    datum_now = datetime.now().strftime("%d.%m.%Y")

    a = KIL_KW_15_Bayreuth_Lieferung_14 = aa
    b = KIL_KW_15_Berlin_Lieferung_14 = bb
    c = KIL_KW_15_Bremen_Lieferung_14 = cc
    d = KIL_KW_15_Dresden_Lieferung_14 = dd
    n = KIL_KW_15_Eichenau_Lieferung_14 = ee
    f = KIL_KW_15_Halle_Lieferung_14 = ff
    g = KIL_KW_15_Hamburg_Lieferung_14 = gg
    h = KIL_KW_15_Hildesheim_Lieferung_14 = hh
    i = KIL_KW_15_Kempten_Lieferung_14 = ii
    j = KIL_KW_15_Köln_Lieferung_14 = jj
    k = KIL_KW_15_Riedstadt_Lieferung_14 = kk
    l = KIL_KW_15_Rostock_Lieferung_14 = ll
    m = KIL_KW_15_Ulm_Lieferung_14 = mm

    p = KIL_KW_16_Zensiert = nn

 
    
    Befehl = "insert into"

    def printer(Befehl):

        print("--Lager Bayreuth")
        for e in Lager_Kostenstellen_ID_Bayreuth:

            print(Befehl + " anf2kst select b_id =" + str(a) + " , kst_id=" + str(e))

        print("--Lager Berlin")
        for e in Lager_Kostenstellen_ID_Berlin:

            print(Befehl + " anf2kst select b_id =" + str(b) + " , kst_id=" + str(e))

        print("--Lager Bremen")
        for e in Lager_Kostenstellen_ID_Bremen:

            print(Befehl + " anf2kst select b_id =" + str(c) + " , kst_id=" + str(e))

        print("--Lager Dresden")
        for e in Lager_Kostenstellen_ID_Dresden:

            print(Befehl + " anf2kst select b_id =" + str(d) + " , kst_id=" + str(e))

        print("--Lager Eichenau")
        for e in Lager_Kostenstellen_ID_Eichenau:

            print(Befehl + " anf2kst select b_id =" + str(n) + " , kst_id=" + str(e))

        print("--Lager Halle")
        for e in Lager_Kostenstellen_ID_Halle:

            print(Befehl + " anf2kst select b_id =" + str(f) + " , kst_id=" + str(e))

        print("--Lager Hamburg")
        for e in Lager_Kostenstellen_ID_Hamburg:

            print(Befehl + " anf2kst select b_id =" + str(g) + " , kst_id=" + str(e))

        print("--Lager Hildesheim")
        for e in Lager_Kostenstellen_ID_Hildesheim:

            print(Befehl + " anf2kst select b_id =" + str(h) + " , kst_id=" + str(e))

        print("--Lager Kempten")
        for e in Lager_Kostenstellen_ID_Kempten:

            print(Befehl + " anf2kst select b_id =" + str(i) + " , kst_id=" + str(e))

        print("--Lager Köln")
        for e in Lager_Kostenstellen_ID_Köln:

            print(Befehl + " anf2kst select b_id =" + str(j) + " , kst_id=" + str(e))

        print("--Lager Riedstadt")
        for e in Lager_Kostenstellen_ID_Riedstadt:

            print(Befehl + " anf2kst select b_id =" + str(k) + " , kst_id=" + str(e))

        print("--Lager Rostock")
        for e in Lager_Kostenstellen_ID_Rostock:

            print(Befehl + " anf2kst select b_id =" + str(l) + " , kst_id=" + str(e))

        print("--Lager Ulm")
        for e in Lager_Kostenstellen_ID_Ulm:

            print(Befehl + " anf2kst select b_id =" + str(m) + " , kst_id=" + str(e))

        print("--ZENISERT")
        for e in ZENSIERT_OE_Zuordnung:

            print(Befehl + " anf2kst select b_id =" + str(p) + " , kst_id=" + str(e))


    userhome = os.path.expanduser('~')
    desktop = userhome + '/Desktop\\'

    global o
    o = (
        desktop
        + "KIL_Liste_Sichtbarkeiten_" + str(KW_V) + "_"
        + str(datum_now)
        + ".sql"
    )

    import sys

    class StdoutRedirection:
        """Standard output redirection context manager"""

        def __init__(self, path):
            self._path = path

        def __enter__(self):
            sys.stdout = open(self._path, mode="w")
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            sys.stdout.close()
            sys.stdout = sys.__stdout__

    with StdoutRedirection(o):
        printer(Befehl)

def b_id_bestätigen():
    Listen_IDs.append(entry_1.get())
    Listen_IDs.append(entry_3.get())
    Listen_IDs.append(entry_4.get())
    Listen_IDs.append(entry_5.get())
    Listen_IDs.append(entry_6.get())
    Listen_IDs.append(entry_7.get())
    Listen_IDs.append(entry_8.get())
    Listen_IDs.append(entry_9.get())
    Listen_IDs.append(entry_10.get())
    Listen_IDs.append(entry_11.get())
    Listen_IDs.append(entry_12.get())
    Listen_IDs.append(entry_13.get())
    Listen_IDs.append(entry_14.get())
    Listen_IDs.append(entry_15.get())
    global KW_V 
    KW_V = entry_20.get()


    string_to_display = entry_1.get()

    label_3 = Label(my_window)
    label_3["text"] = "B-IDs wurden angenommen"
    label_3.grid(row=14, column=1)

    label_4 = Label(my_window)
    label_4["text"] = Listen_IDs
    label_4.grid(row=14, column=2)

    print(Listen_IDs)

my_window = Tk()
my_window.title("KIL Listen Sichtbarkeiten erstellen")

label_1 = Label(my_window, text="Bestell ID Bayreuth:")
label_3 = Label(my_window, text="Bestell ID Berlin:")
label_4 = Label(my_window, text="Bestell ID Bremen:")
label_5 = Label(my_window, text="Bestell ID Dresden:")
label_6 = Label(my_window, text="Bestell ID Eichenau:")
label_7 = Label(my_window, text="Bestell ID Halle:")
label_8 = Label(my_window, text="Bestell ID Hamburg:")
label_9 = Label(my_window, text="Bestell ID Hildesheim:")
label_10 = Label(my_window, text="Bestell ID Kempten:")
label_11 = Label(my_window, text="Bestell ID Köln:")
label_12 = Label(my_window, text="Bestell ID Riedstadt:")
label_13 = Label(my_window, text="Bestell ID Rostock:")
label_14 = Label(my_window, text="Bestell ID Ulm:")
label_15 = Label(my_window, text="ZENSIERT‚")
label_20 = Label(my_window, text="KIL-KW:")

entry_1 = Entry(my_window)
entry_3 = Entry(my_window)
entry_4 = Entry(my_window)
entry_5 = Entry(my_window)
entry_6 = Entry(my_window)
entry_7 = Entry(my_window)
entry_8 = Entry(my_window)
entry_9 = Entry(my_window)
entry_10 = Entry(my_window)
entry_11 = Entry(my_window)
entry_12 = Entry(my_window)
entry_13 = Entry(my_window)
entry_14 = Entry(my_window)
entry_15 = Entry(my_window)
entry_20 = Entry(my_window)


def datei_erstellen(*Listen_Ids):

    Liefertage_erstellen(*Listen_IDs)

    name = "Liste wurde erstellt!"
    string_to_display = name
    label_2 = Label(my_window)
    label_2["text"] = string_to_display
    label_2.grid(row=15, column=1)

    os.startfile(o)


def löschen():
    del Listen_IDs[:]

    label_5 = Label(my_window)
    label_5["text"] = Listen_IDs
    label_5.grid(row=16, column=2)

button_1 = Button(my_window, text="Datei erstellen!", command=datei_erstellen)

button_2 = Button(my_window, text="B-Ids bestätigen!", command=b_id_bestätigen)

button_3 = Button(my_window, text="B-Ids löschen!", command=löschen)



label_1.grid(row=0, column=0)
entry_1.grid(row=0, column=1)


label_3.grid(row=1, column=0)
entry_3.grid(row=1, column=1)

label_4.grid(row=2, column=0)
entry_4.grid(row=2, column=1)

label_5.grid(row=3, column=0)
entry_5.grid(row=3, column=1)

label_6.grid(row=4, column=0)
entry_6.grid(row=4, column=1)

label_7.grid(row=5, column=0)
entry_7.grid(row=5, column=1)

label_8.grid(row=6, column=0)
entry_8.grid(row=6, column=1)

label_9.grid(row=7, column=0)
entry_9.grid(row=7, column=1)

label_10.grid(row=8, column=0)
entry_10.grid(row=8, column=1)

label_11.grid(row=9, column=0)
entry_11.grid(row=9, column=1)

label_12.grid(row=10, column=0)
entry_12.grid(row=10, column=1)

label_13.grid(row=11, column=0)
entry_13.grid(row=11, column=1)

label_14.grid(row=12, column=0)
entry_14.grid(row=12, column=1)

label_15.grid(row=13, column=0)
entry_15.grid(row=13, column=1)

label_20.grid(row=17, column=0)
entry_20.grid(row=17, column=1)


button_1.grid(row=15, column=0)
button_2.grid(row=14, column=0)
button_3.grid(row=16, column=0)

my_window.mainloop()
