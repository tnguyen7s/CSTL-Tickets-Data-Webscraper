from tkinter import Frame
from tkinter.constants import W

class BottomFrame:
    def __init__(self, master):
        self.frame = Frame(master)
        self.frame.grid(row=1, sticky=W)