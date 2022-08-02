# Licensed to TCI, LLC
# Located in W132N10611 Germantown, Wisconsin
# This file is owned and copyrighted by TCI, LLC
#  Author: Jorge Jesus Jurado-Garcia
#  Title: Product Specialist Intern
#   Project Description: Marketing department MTE Product selection
#   Goal: create an easy to use GUI where the marketing department can
#          use for any price increases. Anyone with no background in CCS
#
#  Date of Creation: 12/29/2021
#  Rev:
import tkinter as tk
from tkinter import ttk

import SeleniumScrap as Sel
import file_lookup_imports as FM
import noti_and_checker as noti


"""
Function for Creating new window for new settings and configurations
"""


def New_Product_Window():
    # creates a Tk() object
    master = tk.Tk()

    master.title('TCI LLC - MTE Product Search')
    master.resizable(False, False)
    # ensure that a window is always at the top of the stacking order
    master.attributes('-topmost', 1)
    master.geometry("250x230")

    master.iconbitmap('./assets/tci_logo_Csx_icon.ico')
    master.grid()

    # Label for Product Type
    label = tk.Label(master, text="Product Type")
    label.grid(column=0, columnspan=1, row=0, sticky=tk.W )

    # Create a combobox for Product Type
    master.select_Product = tk.StringVar()
    Product_cb = ttk.Combobox(master, textvariable=master.select_Product)

    # setup the Product Names as values and states
    Product_cb['state'] = 'readonly'
    Product_cb['values'] = ["RL", "RLW", "DVT", "DVS", "SWG", "SWN", "SWGM", "MAP", "MAEP", "RFI"]

    Product_cb.grid(column=3, columnspan=2, row=0, sticky=tk.EW)

    # Setup for Voltage and and Enclosure
    Voltage = tk.Label(master, text="Voltage")
    Enclosure = tk.Label(master, text="Enclosure")
    Voltage.grid(column=0, columnspan=1, row=1, sticky=tk.W )
    Enclosure.grid(column=0, columnspan=1, row=2, sticky=tk.W )

    # Create a combobox for Voltage
    select_Voltage = tk.StringVar()
    Voltage_cb = ttk.Combobox(master, textvariable=select_Voltage)
    Voltage_cb['state'] = 'readonly'
    Voltage_cb['values'] = ["208", "240", "400", "480", "600", "690"]
    Voltage_cb.grid(column=3, columnspan=2, row=1, sticky=tk.EW)


    # Create a combobox for Enclosure
    select_Enclosure = tk.StringVar()
    Enclosure_CB = ttk.Combobox(master, textvariable=select_Enclosure)
    Enclosure_CB['state'] = 'readonly'
    Enclosure_CB['values'] = ["KIT", "Panel Mount", "Modular", "NEMA 1", "NEMA 1/2", "NEMA 3R"]
    Enclosure_CB.grid(column=3, columnspan=2, row=2, sticky=tk.EW)


    def Product_changed():
        """"Handle the product changed event"""
        print("you Selected", master.select_Product.get() )
        print(select_Voltage.get() )

    Product_cb.bind('<<ComboboxSelected>>', Product_changed() )


"""
This class created the main window for UI and Setups of buttons and entry points. 
It also assigns functions and actions to each specific functions when clicked.
"""


class MainFrame(ttk.Frame):
    run_excel: bool

    # Initialization
    def __init__(self, container):
        super().__init__(container)
        # field options
        options = {'padx': 5, 'pady': 5}

        # Name label
        ttk.Label(self, text="Name :").grid(column=0, row=0, sticky=tk.W, **options)

        # Name entry
        self.projectname = tk.StringVar()
        ttk.Entry(self, textvariable=self.projectname).grid(column=1, columnspan=2, row=0, sticky=tk.EW, **options)

        # file_name
        self.filename = str()

        # web driver location
        self.web_driver = str()

        # file_name
        self.new_excel_path = str()

        # Dump excel here
        ttk.Button(self, text="New Location",
                   command=self.new_excel_location).grid(row=1, column=0, sticky=tk.EW, **options)

        # web driver button
        ttk.Button(self, text='Web Driver',
                   command=self.select_webdriver).grid(column=1, columnspan=2, row=1, sticky=tk.EW, **options)

        # excel button
        ttk.Button(self, text='Insert Product Name',
                   command=self.select_excel).grid(column=0, row=2, columnspan=3, sticky=tk.EW, **options)

        self.run_excel = False

        # Create A checkbox Button
        self.btn = tk.Checkbutton(self, text='Use Imported Names', command=self.checkbox,
                                  variable=self.run_excel)
        # self.btn.configure(bg='#ffb3fe')
        self.btn.grid(row=3, column=0, columnspan=3, sticky=tk.EW)

        # start button
        self.button = ttk.Button(self, text="Start Scrap", command=self.start)
        self.button.grid(row=4, column=0, columnspan=3, sticky=tk.EW, **options)

        # Close Button
        tk.Button(self, text="Exit", command=self.close,
                  bg="Red", fg="White").grid(row=5, column=0, columnspan=3, sticky=tk.EW, **options)

        # add padding to the frame and show it
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def new_excel_location(self):
        self.new_excel_path = FM.folder_lookup()

    def start(self):
        """
        In this function we will first check that the excel can be read
        afterwards checking if the excel can be read we will then check the chrome driver being used
        """
        # Debugging Variables
        print("FileName:", self.projectname.get())
        print("new Excel Path", self.new_excel_path)
        print("webdriver", self.web_driver)
        print("Inserted Excel", self.filename)

        # If  Project Name, Webdriver or Excel is null then do not run and tell the user about these null values
        # Afterwards Check if Excel can be run and if the ChromeDriver is the Correct Model if Error exist then tell
        # let the user know by a windows notification
        if self.CheckSUM() is True:
            if self.run_excel is True:
                if self.filename == "":
                    "Tell the user that they have not loaded an excel"
                    noti.InsertExcel()
                    return
                else:
                    "Check if the excel can be read if so save the values and store the data into an array "
                    "afterwards run a different that checks the entries being inserted."
                    self.Product_array = FM.load_excel(self.filename)
                    if self.Product_array is None:
                        noti.DataNotFound()
                        return
                    else:
                        # Go ahead and Open Selenium and Run Excel Variables
                        print("Excel data was found read Variables from Excel")

                        # figure out which Product Line we will be Scraping from
                        # based on the Index of the Product
                        Sel.selenium(self.Product_array, self.web_driver, self.projectname, self.new_excel_path)



            else:
                """
                If wer are not reading from a raw excel allow the user to select which MTE Product Family to Read from
                and store the result to a a variable. Tell the User that they should select a Value that must be ran
                """
                print("Not Reading from Raw Excel allow the user to select from MTE Product Family to Read Combobox")
                # FUTURE New_Product_Window()
        else:
            # debugging
            print("CHeck Entry Function Failed ")


    def CheckSUM(self):
        check = True
        # Retrieve the of name, date, and MTE Selection
        if self.projectname.get() == '':
            FM.ProjectNameNull()
            check = False
        elif self.new_excel_path == '':
            FM.NewExcelLocationNull()
            check = False
        elif self.web_driver == '':
            FM.WebDriverNull()
            check = False
        if check:
            # go ahead and attempt to open selenium and see if so return true else return false
            # Open Webdriver from Chrome
            output = Sel.TestChromeDriver(self.web_driver)
            return output
        else:
            return check


    def select_excel(self):
        self.filename = FM.file_lookup()

    def select_webdriver(self):
        self.web_driver = FM.file_lookup()

    def close(self):
        self.quit()

    def checkbox(self):
        if self.run_excel is False:
            self.run_excel = True
        else:
            self.run_excel = False


"""
Main Object Creation - this class setups the title, Window Attributes,
 Width and height, and Logo for UI
"""


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('TCI LLC - MTE Product Search')
        self.resizable(False, False)
        # ensure that a window is always at the top of the stacking order
        self.attributes('-topmost', 1)

        window_width = 250
        window_height = 220

        # get screen dimension
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # find the center point
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)

        # create the screen on window console
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        # changing the Tinker Logo into the TCI logo instead for our development
        self.iconbitmap('./assets/tci_logo_Csx_icon.ico')
        frm = ttk.Frame(self, padding=1)
        frm.grid()


if __name__ == "__main__":
    app = App()
    MainFrame(app)
    app.mainloop()
