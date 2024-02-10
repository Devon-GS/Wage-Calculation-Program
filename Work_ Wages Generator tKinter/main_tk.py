from tkinter import *

root = Tk()


def myclick():
    pass

# Widget
setup_label = Label(root, text='SETUP',borderwidth=1, relief='solid')

mybutton = Button(root, text='Click Me', width=10 ,command=myclick, fg='blue', bg='red')

mybutton2 = Button(root, text='Click Me', width=10 ,command=myclick, fg='blue', bg='red')


# Bind Widget
setup_label.grid(row=0, column=0, columnspan=3 ,sticky=W+E ,pady=(0,10))
mybutton.grid(row=1, column=0, columnspan=1, padx=(0,10))
mybutton2.grid(row=1, column=1)

# ROOT WINDOW CONFIG
root.title('Wage Calculator')
# root.iconbitmap('icons/smoking.ico')
root.geometry('360x360')
# root.columnconfigure(0, weight=1)

# RUN WINDOW
root.mainloop()