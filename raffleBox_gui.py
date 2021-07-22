import sys
import tkinter as tk
from raffle_box import open_file
from raffle_box import raffle
from raffle_box import export_file
import openpyxl as xl
from tkinter import messagebox
from PIL import Image


# Create root window
root = tk.Tk()
root.iconbitmap('icon.ico')
root.geometry("750x550")
root.title("Welcome to RaffleBox!")
root.configure(bg="white")
imported_file = None
num_winners = None
winners = None
winners_text = None
load_success_msg = None
export_text = None
yes_export_button = None
no_export_button = None
image_file = "getting_result.gif"

info = Image.open(image_file)

frames = info.n_frames  # gives total number of frames that gif contains

# creating list of PhotoImage objects for each frames
im = [tk.PhotoImage(file=image_file, format=f"gif -index {i}") for i in range(frames)]

anim = None


# Create frame
class StartFrame(tk.Frame):
    def __init__(self):
        tk.Frame.__init__(self)
        self.configure(bg="white")
        self.grid_rowconfigure(0, weight=1, minsize=30)
        self.grid_rowconfigure(4, weight=1, minsize=20)
        self.grid_rowconfigure((1, 3), weight=2, minsize=70)
        self.grid_rowconfigure(2, weight=2, minsize=370)
        self.grid_columnconfigure(0, weight=2, minsize=50)
        self.grid_columnconfigure(4, weight=1, minsize=20)
        self.grid_columnconfigure((1, 2), weight=1, minsize=50)
        self.grid_columnconfigure(3, weight=2, minsize=50)
        self.grid()


class NewFrame(tk.Frame):
    def __init__(self):
        tk.Frame.__init__(self)
        self.grid_rowconfigure(0, weight=1, minsize=30)
        self.grid_rowconfigure(4, weight=1, minsize=30)
        self.grid_rowconfigure((1, 3), weight=2, minsize=70)
        self.grid_rowconfigure(2, weight=2, minsize=370)
        self.grid_columnconfigure(0, weight=1, minsize=50)
        self.grid_columnconfigure(4, weight=1, minsize=20)
        self.grid_columnconfigure((1, 2), weight=1, minsize=50)
        self.grid_columnconfigure(3, weight=2, minsize=50)
        self.configure(bg="white")


# Create controller to switch between frames
def controller(frame_to_forget, frame_to_show):
    frame_to_forget.grid_forget()
    frame_to_show.grid()


def animation(frame_counter):
    global gif_label
    raffle_button.grid_forget()
    back_button3.grid_forget()
    global anim
    im2 = im[frame_counter]
    gif_label.configure(image=im2)
    frame_counter += 1
    if frame_counter == frames:
        gif_label.after_cancel(anim)
        anim = None
        gif_label.configure(image='')
        controller(frame4, frame5)
    else:
        anim = gif_label.after(20, lambda: animation(frame_counter))


def import_xl():
    global imported_file
    global upload_file
    global load_success_msg
    imported_file = open_file()
    if type(imported_file) == xl.worksheet.worksheet.Worksheet:
        load_success_msg = tk.Label(frame2, text="The Excel file was successfully loaded. \n Please click Next", bg="white", font=("Rubik", 14))
        upload_file.grid_forget()
        load_success_msg.grid(row=1, column=1)
        return imported_file
    else:
        messagebox.showerror('Unsupported file',
                             message="Please upload a valid Excel file. \n Click on 'Browse files'")
        upload_button.config(state="normal")


def validate_user_entry():
    global input_box
    input_box_value = input_box.get()
    try:
        float(input_box_value)
    except ValueError:
        messagebox.showerror('Invalid input',
                                    message="Please, enter a  number to select the number of winners")
        input_box.delete(0, 'end')
    else:
        if float(input_box_value).is_integer() is False:
            messagebox.showerror('Invalid input',
                                 message="Please, enter a whole number to select the number of winners. Decimal numbers are not accepted")
            input_box.delete(0, 'end')
        elif int(input_box_value) <= 0:
            messagebox.showerror('Invalid input',
                                 message="Please, enter a number higher than 0 to select the number of winners")
            input_box.delete(0, 'end')
        elif imported_file.max_row < int(input_box_value):
            messagebox.showerror('Invalid input',
                                 message="The number of winners cannot be greater than the number of contestants")
            input_box.delete(0, 'end')
        elif input_box_value == "":
            messagebox.showerror('Invalid input',
                                 message="Please, do not forget to type the number of winners")
            input_box.delete(0, 'end')
        else:
            next_button3.config(command=controller(frame3, frame4))


def result(xl_sheet, num_picks_entry):
    global winners
    global winners_text
    global num_winners
    global export_text
    global yes_export_button
    global no_export_button
    num_winners = num_picks_entry.get()
    winners = raffle(xl_sheet, int(num_winners))
    if type(winners) == list:
        if len(winners) == 1:
            winners_text = \
                tk.Label(winners_frame, text="The winner is: \n\n" + ',\n'.join(winners), bg="white", font=("Rubik", 14), anchor='center')
        else:
            winners_text = \
                tk.Label(winners_frame, text="The winners are: \n\n" + ',\n'.join(winners), bg="white", font=("Rubik", 14), anchor='center')
    else:
        if len(winners) == 1:
            winners_text = \
                tk.Label(winners_frame,
                         text="The winner is: \n\n" + '\n'.join("{}: {}".format(key, winners[key]) for key, winners[key] in winners.items()),
                         bg="white", font=("Rubik", 14), anchor='center')
        else:
            winners_text = \
                tk.Label(winners_frame,
                         text="The winners are: \n\n" + '\n'.join("{}: {}".format(key, winners[key]) for key, winners[key] in winners.items()),
                         bg="white", font=("Rubik", 14), anchor='center')
    winners_text.grid(row=0, column=1, columnspan=3, padx=(20, 0))
    get_outcome.grid_forget()
    raffle_button.grid_forget()
    export_text = tk.Label(frame5, text="Would you like to export the result to an Excel file?", bg="white", font=("Rubik", 14))
    export_text.grid(row=2, column=2, padx=(30, 0), sticky='n', pady=(30, 0))
    yes_export_button = tk.Button(frame5, width=12, height=4, text="Yes", command=lambda: [export_xl(), controller(frame5, frame6)])
    yes_export_button.grid(row=2, column=2, sticky='e', padx=(30, 0))
    no_export_button = tk.Button(frame5, width=12, height=4, text="No", command=lambda: [controller(frame5, frame6)])
    no_export_button.grid(row=2, column=2, sticky='w', padx=(30, 0))
    return winners, winners_text


def export_xl():
    export_file(winners)


def clear_data():
    global imported_file
    global winners
    global num_winners
    imported_file = None
    winners = None
    num_winners = None


def set_new_iteration():
    global winners_text
    global load_success_msg
    global yes_export_button
    global no_export_button
    winners_text.destroy()
    load_success_msg.destroy()
    export_text.destroy()
    upload_file.grid(row=1, column=1)
    yes_export_button.grid_forget()
    no_export_button.grid_forget()
    raffle_button.grid(row=3, column=3)
    next_button3.grid(row=3, column=3)
    get_outcome.grid(row=1, column=1)
    back_button3.grid(row=3, column=2, padx=(0, 20))
    upload_button.config(state="normal")
    next_button2.config(state="disabled")
    winners_canvas.configure(yscrollcommand=scrollbar.set)
    input_box.delete(0, 'end')


# Create frame1 - welcome users and provide guidance on how to use app
frame1 = StartFrame()

# Frame 1: Welcome
welcome = tk.Label(frame1, text="Welcome to RaffleBox!", bg="white", font=("Rubik", 18, "bold")).\
    grid(column=1, row=1, columnspan=2)

# Frame 1: instructions
instructions = tk.Message(frame1, text="""
RaffleBox is intended to be used to raffle one or more people randomly from a list (placed in 1 or 2 Excel columns).
Step 1: prepare your Excel file. Create a sheet named 'list' (case sensitive) and add the list of people (or other item types you want to raffle).
You can add a list of names in column A and an attribute (for example the e-mail address) in column B.
Step 2: upload the Excel sheet.
Step 3: type the number of 'winners' you want to be picked.
Step 4: click 'Raffle' and see the result of the random pick.
Step 5: if you want, you can export the result to a new Excel file.
Finally, you can start a new raffle or close the app.
""", bg="white", font=("Rubik", 14), justify="left")
instructions.grid(column=1, row=2, columnspan=2)
next_button1 = tk.Button(frame1, text="Next", command=lambda: controller(frame1, frame2), width=15, height=3)
next_button1.grid(column=3, row=3)

# Create frame 2 - uses the function load_xl to load an Excel file
frame2 = NewFrame()
upload_file = tk.Label(frame2, text="Click 'Browse files' to upload an Excel file", bg="white", font=("Rubik", 14))
upload_file.grid(row=1, column=1)
upload_button = tk.Button(frame2, text='Browse files', width=12, height=4, command=lambda: [upload_button.config(
                                                                                                state="disabled"),import_xl(),
                                                                                            next_button2.config(
                                                                                                state="normal")])
upload_button.grid(row=1, column=2)
next_button2 = tk.Button(frame2, text="Next", command=lambda: controller(frame2, frame3), width=15, height=3)
next_button2.grid(row=3, column=3)
next_button2.config(state="disabled")
back_button1 = tk.Button(frame2, text="Go back", command=lambda: controller(frame2, frame1), width=15, height=3)
back_button1.grid(row=3, column=2, padx=(0, 20))

# Create frame 3 -  Input the desired number of winners and uses the function validate_user_entry to validate
# the user input
frame3 = NewFrame()
prompt_num_winners = tk.Label(frame3, text="How many winners should I pick?", bg="white", font=("Rubik", 14))
prompt_num_winners.grid(row=1, column=1)
input_box = tk.Entry(frame3, width=10)
input_box.grid(row=1, column=2)
next_button3 = tk.Button(frame3, text="Next", command=lambda: validate_user_entry(), width=15, height=3)
next_button3.grid(row=3, column=3)
back_button2 = tk.Button(frame3, text="Go back", command=lambda: controller(frame3, frame2), width=15, height=3)
back_button2.grid(row=3, column=2, padx=(0, 20))

# Create frame 4 - Start the Raffle by clicking the Raffle button
frame4 = NewFrame()
get_outcome = tk.Label(frame4, text="Click Raffle to select the winner/s", bg="white", font=("Rubik", 14))
get_outcome.grid(row=1, column=1)
raffle_button = tk.Button\
    (frame4, text="Raffle",
     command=lambda: [result(imported_file, input_box), animation(0)], width=15, height=3)
raffle_button.grid(row=3, column=3)
back_button3 = tk.Button(frame4, text="Go back", command=lambda: controller(frame4, frame3), width=15, height=3)
back_button3.grid(row=3, column=2, padx=(0, 20))
gif_label = tk.Label(frame4, image="", bg="white")
gif_label.grid(row=2, column=2, sticky='e', padx=60)

# Create frame 5 - Uses the functions result() and export_xl() to display the winners and ask if the user wants to
# export the result to an Excel file
frame5 = NewFrame()

# Create scrollable canvas in case the list of winners is too long
winners_canvas = tk.Canvas(frame5, bg='white', width=550)
winners_canvas.grid(row=0, column=0, rowspan=2, columnspan=4, padx=(24, 0))
scrollbar = tk.Scrollbar(frame5, orient='vertical', command=winners_canvas.yview)
scrollbar.grid(row=0, column=4, rowspan=2, sticky='ns')
winners_frame = tk.Frame(winners_canvas, bg='white')
winners_frame.bind('<Configure>', lambda e: winners_canvas.configure(scrollregion=winners_canvas.bbox('all')))
winners_canvas.create_window((0, 0), window=winners_frame, anchor="nw")
winners_canvas.configure(yscrollcommand=scrollbar.set)


# Create frame 6 - Thanks the user and asks if a new iteration should be started
frame6 = NewFrame()
goodbye_msg = tk.Label(frame6, text="Thank you for using RaffleBox!", bg="white", font=("Rubik", 14))
goodbye_msg.grid(row=2, column=3, columnspan=2, sticky='n', padx=(80, 0))
start_over_button = \
    tk.Button(frame6, width=15, height=4, text="Start a new Raffle!",
              command=lambda: [clear_data(), set_new_iteration(), controller(frame6, frame1)])
start_over_button.grid(column=3, row=2, sticky='e', padx=(80, 0))
quit_button = tk.Button(frame6, text="Quit RaffleBox", command=lambda: [root.destroy(), sys.exit()], width=15, height=3)
quit_button.grid(row=3, column=3, sticky='e', padx=(80, 0))

root.mainloop()
