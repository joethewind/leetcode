from functools import partial
import tkinter

root = tkinter.Tk()
Mybutton = partial(tkinter.Button,root,fg = 'white', bg = 'blue')
b1 = Mybutton(text = 'Button1')
b2 = Mybutton(text = 'Button2')
qb = Mybutton(text = 'QUIT', bg = 'red', command = root.quit())

b1.pack()
b2.pack()
qb.pack(fill = tkinter.X,expand = True)
root.title('PFAs!')
root.mainloop()


