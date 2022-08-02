import tkinter as tk
from tkinter import ttk

read_info = """
# MTE_Product_Search Version 2
Python Scripted for Web scrapping Product 
information from MTE Corp
Here is the steps and procedures in order 
to get the code working on your own personal computer

# New Project Procedure
Clone project and download Pycharm IDE
            OR
1. Create a python project using any IDE you like.
2. Create two new directories inside your project. 
   Named "assets" and "chromedriver_win32" respectively. 
3. Download tci_logo_Csx_icon.ico, store 
   into the assets folder. 
4. Download Chrome and check your chrome version.
   Settings->About Chrome. 
5. Follow link https://chromedriver.chromium.org/downloads 
   download chromedriver version based on your current 
   Chrome Browser Version. 
6. Store the chromeDriver applications into 
   chromedriver_win32 folder. 
8. In your interpreter download the following Packages:
          
          * Selenium 4.10 
          * tkinterx 0.0.9
          * openpyxl 3.0.9
          * plyer
          * pandas
          

Note: To represent or use your own personal logo simply 
insert an ico file into assets and change this line to 
represent your file name:
         
    self.iconbitmap('./assets/tci_logo_Csx_icon.ico')

# GUI Procedure

# Output
New excel will be output in your selected excel locations. 
If no excel location was selected the file will be found 
inside your python project. This project will have the 
product name, description, and Price of each product. 

"""
ProductSelection = {
    'Type' :"",
    'Style':"",
    'Voltage':""
}


"""
Text Window function that creates a window with reading information on the overall app 
and what the app procedure and outputs 
"""
def text_window(str, info):
    win = tk.Toplevel()

    window_width = 500
    window_height = 350

    # get screen dimension
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # find the center point
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)

    # create the screen on window console
    win.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    win.title(str)
    win.resizable(False, False)
    win.iconbitmap('./assets/tci_logo_Csx_icon.ico')
    win.attributes('-topmost', 1)

    # create a text widget and specify size
    T = tk.Text(win, height=18, width=70)

    # create a label
    l = tk.Label(win, text=str)
    l.config(font=("Courier", 16))

    # create an exit button
    b_exit = tk.Button(win, text="Exit", command=win.destroy)

    # create a Scrollbar and associate it with txt
    scrollb = tk.Scrollbar(win, command=T.yview)
    l.pack(side=tk.TOP)
    scrollb.pack(side=tk.RIGHT)
    T.pack()
    T.insert(tk.END, info)
    b_exit.pack()




"""
New Window for product search that holds stores style,type,voltage for product information inside dictionary
that automatically updates whenever the user selects the information and clicks on the done button. 
"""
def Product_Window(str):
    # creates a Tk() object
    master = tk.Tk()

    master.title(str)
    master.resizable(False, False)
    # ensure that a window is always at the top of the stacking order
    master.attributes('-topmost', 1)
    master.geometry("330x250")

    master.iconbitmap('./assets/tci_logo_Csx_icon.ico')
    master.grid()

    # MTE product description IE LABEL
    master.lf = ttk.LabelFrame(text="Please select Product Type?")
    master.lf.grid(column=0, row=1, padx=20, pady=20)

    global selected_product
    selected_product = tk.StringVar()

    selections = (('RL Reactor', 'RL'),
                  ('RLW Reactor', 'RLW'),
                  ('DV E-Series Filter', 'DVT'),
                  ('DV Sentry Filter', 'DVS'),
                  ('Sinewave Guardian', 'SWG'),
                  ('Sinewave Nexus', 'SWN'),
                  ('High Freq. Sinewave Gaurdian', 'SWGM'),
                  ('Matrix AP Filters', 'MAP'),
                  ('Matrix E-Series Filters', 'MAEP'),
                  ('RFI EMI Filters', 'RF3'),
                  )
    bupdate = True
    grid_row = 3
    for selection in selections:
        # create a radio button
        master.radio = ttk.Radiobutton(master, text=selection[0], value=selection[1],
                                       variable=selected_product)
        if not bupdate:
            master.radio.grid(column=1, row=grid_row, ipadx=10, ipady=10)
            bupdate = True
        else:
            master.radio.grid(column=0, row=grid_row, ipadx=10, ipady=10)
            bupdate = False
            grid_row -= 1
        # grid column
        grid_row += 1

    # Close Button
    tk.Button(master, text="Done", command=Done,
              bg="Blue", fg="White").grid(row=10, column=0, columnspan=2, sticky=tk.EW,)

"""
Function that stores product information inside a dictionary 
"""
def Done():
    print("done with product information settings")
    ProductSelection['Type'] = selected_product.get()
    print(ProductSelection)
    print(selected_product.get())
