import tkinter as tk
from tkinter import ttk, S, N, E, W, StringVar, messagebox
from tkinter.ttk import Frame, Label, OptionMenu
from OneInch import OneInchUI
from HalfInch import HalfInchUI
import json

LARGE_FONT = ("Verdana", 12)


class MainWindow(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.iconbitmap('hewlett-packard-logo-black-and-white.ico')
        self.geometry("1200x900")
        self.title("Daily Yield and Reject Entry")
        container = tk.Frame(self)

        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        print(MainFrame)

        for F in (MainFrame, PageOne, PageTwo):
            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky=(N, E, W, S))
            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)

        self.show_frame(MainFrame)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


file_selection = ""


# execute this function before closing application
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        # load the config file first
        with open('config.json', 'r') as f:
            config = json.load(f)
        # edit the data
        config['file_selection'] = file_selection

        # write it back to the file
        with open('config.json', 'w') as f:
            json.dump(config, f)
        app.destroy()


class MainFrame(tk.Frame):
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)  # this line makes the window resolution crisp

    # save config before closing

    @staticmethod
    def raise_frame(frame):
        frame.tkraise()

    def callback(self, optFile, lblTitle, oneInchFrame, halfInchFrame):
        print(optFile.get())
        global file_selection
        file_selection = optFile.get()

        if optFile.get() == "Half Inch":
            print("inside Half Inch")
            lblTitle["text"] = "Half Inch Daily Yield and Reject Form"
            self.raise_frame(halfInchFrame)
        else:
            print("inside One Inch")
            lblTitle["text"] = "One Inch Daily Yield and Reject Form"
            self.raise_frame(oneInchFrame)

    def __init__(self, parent, controller):
        # file selection
        OPTIONS = ["One Inch", "Half Inch", ]

        tk.Frame.__init__(self, parent)

        content = Frame(self, padding=(12, 12, 12, 12))
        content.grid(column=0, row=0, sticky=(N, S, E, W))
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)
        content.rowconfigure(1, weight=15)

        # content window
        lblTitle = Label(content, text="One Inch Daily Yield and Reject Form",
                         font=("arial", 16, "bold"))
        lblTitle.grid(row=0, column=0, sticky=N)

        halfInchFrame = Frame(content, borderwidth=5, relief="solid")
        halfInchFrame.grid(row=1, column=0, columnspan=2, sticky=(N, S, E, W))
        halfInchFrame.rowconfigure(0, weight=1)
        halfInchFrame.rowconfigure(1, weight=15)
        halfInchFrame.columnconfigure(0, weight=1)
        halfInchFrame.columnconfigure(1, weight=1)

        oneInchFrame = Frame(content, borderwidth=5, relief="solid")
        oneInchFrame.grid(row=1, column=0, columnspan=2, sticky=(N, S, E, W))
        oneInchFrame.rowconfigure(0, weight=1)
        oneInchFrame.rowconfigure(1, weight=15)
        oneInchFrame.columnconfigure(0, weight=1)
        oneInchFrame.columnconfigure(1, weight=1)

        # open configuration file
        with open('config.json') as config_file:
            data = json.load(config_file)

        optFile = StringVar()
        global file_selection
        file_selection = data['file_selection']
        print(file_selection)

        optFile.set(file_selection)

        # set selection based on config
        self.callback(optFile, lblTitle, oneInchFrame, halfInchFrame)

        # option Menu to select which file
        optionMenu = OptionMenu(content, optFile, file_selection, OPTIONS[0], OPTIONS[1],
                                command=lambda x: self.callback(optFile, lblTitle, oneInchFrame,
                                                                halfInchFrame))
        optionMenu.grid(row=0, column=1, pady=5, padx=5, sticky=(E, N))

        ''' One Inch Frame '''
        oneInchUI = OneInchUI.OneInch(oneInchFrame)
        oneInchUI.grid(row=1, column=0, columnspan=2, sticky=(N, S, E, W))
        oneInchUI.grid_rowconfigure(0, weight=1)
        oneInchUI.grid_columnconfigure(0, weight=1)

        ''' Half Inch Frame '''
        halfInchUI = HalfInchUI.HalfInch(halfInchFrame)
        halfInchUI.grid(row=1, column=0, columnspan=2, sticky=(N, S, E, W))
        halfInchUI.grid_rowconfigure(0, weight=1)
        halfInchUI.grid_columnconfigure(0, weight=1)


class PageOne(tk.Frame):
    ''' This serve for further expansion'''

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text="Page One", font=LARGE_FONT)
        label.pack(pady=10, padx=10)

        button1 = ttk.Button(self, text="back to Start Page",
                             command=lambda: controller.show_frame(MainFrame))
        button1.pack()

        button1 = ttk.Button(self, text="visit page 2",
                             command=lambda: controller.show_frame(PageTwo))
        button1.pack()


class PageTwo(tk.Frame):
    ''' This serve for further expansion'''

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text="Graph Page", font=LARGE_FONT)
        label.pack(pady=10, padx=10)

        button1 = ttk.Button(self, text="back to Start Page",
                             command=lambda: controller.show_frame(MainFrame))
        button1.pack()

        button1 = ttk.Button(self, text="Page One",
                             command=lambda: controller.show_frame(PageOne))
        button1.pack()


app = MainWindow()
app.protocol("WM_DELETE_WINDOW", lambda: on_closing())
app.mainloop()
