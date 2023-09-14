# import tkinter as tk
# from tkinter import ttk
# from tkinter.messagebox import showinfo
# from calendar import month_name

# root = tk.Tk()

# # config the root window
# root.geometry('300x200')
# root.resizable(False, False)
# root.title('Combobox Widget')

# # label
# label = ttk.Label(text="Please select a month:")
# label.pack(fill=tk.X, padx=5, pady=5)

# # create a combobox
# selected_month = tk.StringVar()
# month_cb = ttk.Combobox(root, textvariable=selected_month)

# # get first 3 letters of every month name
# month_cb['values'] = [month_name[m][0:3] for m in range(1, 13)]

# # prevent typing a value
# month_cb['state'] = 'readonly'

# # place the widget
# month_cb.pack(fill=tk.X, padx=5, pady=5)


# # bind the selected value changes
# def month_changed(event):
#     """ handle the month changed event """
#     showinfo(
#         title='Result',
#         message=f'You selected {selected_month.get()}!'
#     )

# month_cb.bind('<<ComboboxSelected>>', month_changed)

# root.mainloop()


# text = "1/2in test text"

# parts = text.split("in")

# print(parts[0])


# list = []
# x = 0
# value = 10
# while x < 10:
#     x += 1
#     print(list)
#     list.append(value + 10)


from tkinter import *

# Create an instance of Tkinter frame
win = Tk()

# Define empty variables
var1 = IntVar()
var2 = IntVar()

# Function to display the input value
def display_input():
    print("Input for Python:", var1.get())
    print("Input for C++:", var2.get())

# Define a Checkbox
t1 = Checkbutton(win, text="Python", variable=var1, onvalue=1, offvalue=0)
t1.pack()
t2 = Checkbutton(win, text="C++", variable=var2, onvalue=1, offvalue=0)
t2.pack()



# Create a button to trigger the display_input function
button = Button(win, text="Get Values", command=display_input)
button.pack()

win.mainloop()