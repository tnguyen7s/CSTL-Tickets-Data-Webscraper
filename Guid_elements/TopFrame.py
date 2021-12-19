import os
from tkinter import Label, PhotoImage

from Enums.GuidParameter import GuidParameter


class TopFrame:
    def __init__(self, master):
         # semo images
        self.semo_image = PhotoImage(file = GuidParameter.SEMO_LOGO_IMAGE_FILE)
        self.top_frame = Label(master, image=self.semo_image, bg=GuidParameter.PINK_RED)
        self.top_frame.grid(row=0, column=0, columnspan=5)
