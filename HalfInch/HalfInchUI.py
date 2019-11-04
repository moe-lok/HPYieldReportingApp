import tkinter as tk
from tkinter import N, S, E, W, filedialog, StringVar
from tkinter.ttk import Notebook, Frame, Label, Button
import HalfInch.RaptorBonderUI as rb
import HalfInch.RaptorStakerUI as rs
import HalfInch.RaptorEncapUI as re
import HalfInch.RaptorProtectCoverlayerUI as rpcl
import HalfInch.RaptorAttachCoverlayerUI as racl
import HalfInch.HVIEUI as Hvie
import json
import atexit


class HalfInch(tk.Frame):
    # open configuration file
    with open('config.json') as config_file:
        data = json.load(config_file)
    conf_file_location = data['file_location']['Half_Inch']
    file_location = conf_file_location
    print("file location ######")
    print(file_location)

    @staticmethod
    def browseFile(strFileDir):
        global file_location
        print("browse file")
        try:
            file_location = filedialog.askopenfilename(title="Select file",
                                                       filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            strFileDir.set(file_location)
        except Exception as e:
            print(e)

    def resetFile(self, strFileDir):
        # load from config file
        global file_location
        file_location = self.conf_file_location
        strFileDir.set(file_location)

    @staticmethod
    def exit_handler():
        print("Half Inch is ending")
        with open('config.json', 'r') as f:
            config = json.load(f)
            # edit the data
        print("file location before exit ######")
        print(file_location)
        config['file_location']['Half_Inch'] = file_location

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

        # tabControl Half Inch
        tabHalfInch = Notebook(root)  # Create Tab Control
        tabHalfInch.rowconfigure(0, weight=1)
        tabHalfInch.columnconfigure(0, weight=1)

        raptorBonder = Frame(tabHalfInch)  # Create a tab
        raptorStaker = Frame(tabHalfInch)
        raptorEncap = Frame(tabHalfInch)
        raptorProtectCoverlayer = Frame(tabHalfInch)
        raptorAttachCoverlayer = Frame(tabHalfInch)
        hvie = Frame(tabHalfInch)

        tabHalfInch.add(raptorBonder, text='Raptor Bonder')  # Add the tab
        tabHalfInch.add(raptorStaker, text='Raptor Staker')
        tabHalfInch.add(raptorEncap, text='Raptor Encap')
        tabHalfInch.add(raptorProtectCoverlayer, text='Raptor Protect Coverlayer')
        tabHalfInch.add(raptorAttachCoverlayer, text='Raptor Attach Coverlayer')
        tabHalfInch.add(hvie, text='HVIE')
        tabHalfInch.grid(row=1, column=0, columnspan=3, sticky=(N, S, E, W))  # Pack to make visible

        rb.RaptorBonder(raptorBonder, strFileDir)
        rs.RaptorStaker(raptorStaker, strFileDir)
        re.RaptorEncap(raptorEncap, strFileDir)
        rpcl.RaptorProtectCoverlayer(raptorProtectCoverlayer, strFileDir)
        racl.RaptorAttachCoverlayer(raptorAttachCoverlayer, strFileDir)
        Hvie.HVIE(hvie, strFileDir)

        atexit.register(self.exit_handler)
