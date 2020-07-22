import pandas as pd
import math
from pathlib import Path
import tkinter as tk
from PIL import ImageTk, Image

def add_to_hash(table, sample):
    for index, part in enumerate(sample['Part']):
        table[part] = table.get(part, 0) + 1


def check_remove_list(table, checked, year_count):
    for index, value in enumerate(table.items()):
        if value[1] > (year_count - 1):
            filt = checked['Part'] == value[0]
            checked.drop(index=checked[filt].index, inplace=True)



def sample(sample_size, checked, letter_class, table):
    if math.ceil(sample_size) <= len(checked["Part"]):
        return math.ceil(sample_size)
    elif len(checked["Part"]) > 0:
        return len(checked["Part"])
    else:
        checked.iloc[:, :] = letter_class.copy()
        table.clear()
        return math.ceil(sample_size)


def out_put_text():
    output_label['text'] = "Your Weekly Counter is ready! \nThe files can be found in the 'Weekly Counter' folder."

def error_text():
    output_label['text'] = "There was an error. Check that file name is correct"


def generate_weekly_counter(file_name):
    
    #sets the file path and gets raw data
    try:
        PATH = file_name
        file = pd.ExcelFile(PATH)
        raw_data = file.parse('Sheet1')

        #groups into 4 classes 'A' 'B' 'C' 'D'
        grouped = raw_data.groupby('Classification')

        A_class = grouped.get_group("A")
        B_class = grouped.get_group("B")
        C_class = grouped.get_group("C")
        D_class = grouped.get_group("D")

        #initializes the hash checkeded lists

        #A Class items must be counted 12 times per year, so split into 4 groups
        A_length = len(A_class['Classification'])
        A_sample_size = A_length / 4
        A_hash = dict()
        A_checked = A_class.copy() #Run this block to reset A_check 

        #B class items must be counted 6 times per year, so split into 8 groups
        B_length = len(B_class['Classification'])
        B_sample_size = B_length / 8
        B_hash = dict()
        B_checked = B_class.copy()

        #C class items must be counted 3 times per year, so split into 16 groups
        C_length = len(C_class['Classification'])
        C_sample_size = C_length / 16
        C_hash = dict()
        C_checked = C_class.copy()

        #D class items must be counted 1 time per year, so split into 52 groups
        D_length = len(D_class['Classification'])
        D_sample_size = D_length / 52
        D_hash = dict()
        D_checked = D_class.copy()

        for i in range(52):

            full_List = pd.DataFrame()

            #A sample
            A_sample = A_checked.sample(n=sample(A_sample_size, A_checked, A_class, A_hash), replace=False)
            add_to_hash(A_hash, A_sample)
            check_remove_list(A_hash, A_checked, 12)

            #B sample
            B_sample = B_checked.sample(n=sample(B_sample_size, B_checked, B_class, B_hash), replace=False)
            add_to_hash(B_hash, B_sample)
            check_remove_list(B_hash, B_checked, 6)

            #C sample
            C_sample = C_checked.sample(n=sample(C_sample_size, C_checked, C_class, C_hash), replace=False)
            add_to_hash(C_hash, C_sample)
            check_remove_list(C_hash, C_checked, 3)

            #D sample
            D_sample = D_checked.sample(n=sample(D_sample_size, D_checked, D_class, D_hash), replace=False)
            add_to_hash(D_hash, D_sample)
            check_remove_list(D_hash, D_checked, 1)

            full_List = full_List.append([A_sample, B_sample, C_sample, D_sample], ignore_index=True)
            full_List.to_excel(Path(f"./Weekly Counter/Week_{i + 1}_Counter.xlsx"), index=False)

        out_put_text()
    except:
        error_text()





WIDTH = 800
HEIGHT = 500

#H-E-B Red color is #ee3324

root = tk.Tk()

root.title("Random Counter")

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH, bg='#fbf9f5')
canvas.pack()

#Image Label
#img = ImageTk.PhotoImage(Image.open('HEB_Logo.jpg'))
#logo_label = tk.Label(root, image=img)
#logo_label.place(relx=0.9, rely=0.05, relwidth=0.2, relheight=0.1, anchor='ne')

#Title Frame 
title_frame = tk.Frame(root)
title_frame.place(relx=0.5, rely=0.17,relwidth=0.75, relheight=0.05, anchor='n')

title_txt = "Weekly Random Counter Generator"

title = tk.Label(title_frame, text=title_txt, bg='#fbf9f5')
title.configure(font='Arial 28 underline')
title.place(relwidth=1, relheight=1)
#End of Title


#Description frame
desc_frame = tk.Frame(root, bg='#ee3324', bd=2)
desc_frame.place(relx=0.5, rely=0.3, relwidth=0.75, relheight=0.1, anchor='n')

description_text = "This will generate a random count of your parts in the store room and return a list of Excel documents. These documents can be found in the 'Weekly Counter' folder."
desc_label = tk.Message(desc_frame, text=description_text, bg='#fbf9f5', anchor='nw', width=575, padx=5, pady=10)
desc_label.configure(font='Arial 12')
desc_label.place(relwidth=1, relheight=1)
#End of Description frame


#Input and button Frame
button_frame = tk.Frame(root, bg='#fbf9f5')
button_frame.place(relx=0.5, rely=0.43, relwidth=0.75, relheight=0.1, anchor='n')

file_label = tk.Label(button_frame, text='Enter File Here:', anchor='center', bg='#fbf9f5')
file_label.configure(font='Arial 16 bold')
file_label.place(relx=0, rely=0.05, relwidth=0.2, relheight=0.9)

#File name entry
file_entry = tk.Entry(button_frame, text='Enter File Name Here', bg='#fbf9f5', bd=5)
file_entry.place(relx=0.2, rely=0.05, relwidth=0.5, relheight=0.9)

button = tk.Button(button_frame, text="Run", command=lambda: generate_weekly_counter(file_entry.get()))
button.place(relx=0.7, rely=0.05, relwidth=0.3, relheight=0.9)

output_frame = tk.Frame(root, bg='#ee3324', bd=2)
output_frame.place(relx=0.5, rely=0.65, relwidth=0.75, relheight=0.3, anchor='n')

output_label = tk.Label(output_frame, text=" ", bg='#fbf9f5')
output_label.place(relx=0.5, relwidth=1, relheight=1, anchor='n')


root.mainloop()

