import tkinter as tk
from tkinter import N, S, E, W, filedialog, StringVar
from tkinter.ttk import Notebook, Frame, Button, Label
import OneInch.TarzanBonderUI as tb
import OneInch.TarzanStakerUI as ts
import OneInch.TarzanCapAttachUI as tca
import OneInch.TarzanCoverlayerUI as tcl
import OneInch.TarzanEncapUI as te
import OneInch.HVFAUI as HV
import json
import atexit


class OneInch(tk.Frame):
    # open configuration file
    with open('config.json') as config_file:
        data = json.load(config_file)
    conf_file_location = data['file_location']['One_Inch']
    file_location = conf_file_location
    print("conf file location ######")
    print(conf_file_location)
    print("file location ######")
    print(file_location)

    def browseFile(self, strFileDir):
        global file_location
        print("browse file")
        try:
            file_location = filedialog.askopenfilename(title="Select file",
                                                       filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            strFileDir.set(file_location)
            print(file_location)
            # hp.one_Inch_file = file_location
        except Exception as e:
            print(e)

    def resetFile(self, strFileDir):
        # load from config file
        global file_location
        file_location = self.conf_file_location
        strFileDir.set(file_location)

    @staticmethod
    def exit_handler():
        print("One Inch is ending")
        with open('config.json', 'r') as f:
            config = json.load(f)
            # edit the data

        print("file location before exit ######")
        print(file_location)
        config['file_location']['One_Inch'] = file_location

        # write it back to the file
        with open('config.json', 'w') as f:
            json.dump(config, f)

    def __init__(self, root):
        tk.Frame.__init__(self, root)

        # string variable
        strFileDir = StringVar()
        strFileDir.set(self.file_location)

        lblFileDir = Label(root, textvariable=strFileDir)
        lblFileDir.grid(row=0, column=0, sticky=W)
        btnBrowse = Button(root, text="browse", command=lambda: self.browseFile(strFileDir)).grid(row=0, column=1,
                                                                                                  pady=5, padx=5,
                                                                                                  sticky=E)

        btnReset = Button(root, text="reset", command=lambda: self.resetFile(strFileDir)).grid(row=0, column=2, pady=5,
                                                                                               padx=5, sticky=E)

        # tabControl One Inch
        tabOneInch = Notebook(root)  # Create Tab Control
        tabOneInch.rowconfigure(0, weight=1)
        tabOneInch.columnconfigure(0, weight=1)

        tarzanBonder = Frame(tabOneInch)  # Create a tab
        tarzanBonder.columnconfigure(0, weight=1)
        tarzanBonder.columnconfigure(1, weight=1)
        tarzanBonder.rowconfigure(0, weight=1)
        tarzanBonder.rowconfigure(1, weight=1)
        tarzanBonder.rowconfigure(2, weight=1)
        tarzanBonder.rowconfigure(3, weight=1)

        tarzanStaker = Frame(tabOneInch)
        tarzanCapAttach = Frame(tabOneInch)
        tarzanCoverlayer = Frame(tabOneInch)
        tarzanEncap = Frame(tabOneInch)
        hvfa = Frame(tabOneInch)

        tabOneInch.add(tarzanBonder, text='Tarzan Bonder')  # Add the tab
        tabOneInch.add(tarzanStaker, text='Tarzan Staker')
        tabOneInch.add(tarzanCapAttach, text='Tarzan Cap Attach')
        tabOneInch.add(tarzanCoverlayer, text='Tarzan Coverlayer')
        tabOneInch.add(tarzanEncap, text='Tarzan Encap')
        tabOneInch.add(hvfa, text='HVFA')
        tabOneInch.grid(row=1, column=0, columnspan=3, sticky=(N, S, E, W))  # Pack to make visible

        tb.TarzanBonder(tarzanBonder, strFileDir)
        ts.TarzanStaker(tarzanStaker, strFileDir)
        tca.TarzanCapAttach(tarzanCapAttach, strFileDir)
        tcl.TarzanCoverlayer(tarzanCoverlayer, strFileDir)
        te.TarzanEncap(tarzanEncap, strFileDir)
        HV.HVFA(hvfa, strFileDir)

        atexit.register(self.exit_handler)
