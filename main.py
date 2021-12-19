# Import module
from tkinter import Button, Canvas, Frame, PhotoImage, Tk, ttk
from tkinter.constants import BOTH, CENTER, NSEW, NW, SOLID, TOP, W
from tkinter.ttk import Notebook
from tkinter.font import BOLD, Font

from selenium import webdriver

from Enums.GuidParameter import GuidParameter
from Guid_elements.BottomFrame import BottomFrame
from Guid_elements.LeftFrame import LeftFrame
from Guid_elements.MiddleFrame import MiddleFrame
from Guid_elements.TopFrame import TopFrame

def openbrowser():
    webdriver.Firefox().get("http://localhost:8080/cstlmonthlyticketreportview/home")

# The master root
root = Tk()
root.title("CSTL-WHD")


# The notebook(<root)
notebook = Notebook(root, padding=0)

# Style
style = ttk.Style()
style.theme_create("MyStyle", parent="alt", settings={
        "TNotebook": {"configure": {"tabmargins": [0,0,0,0], "background": GuidParameter.BLACK} },
        "TNotebook.Tab": {"configure": {"padding": [10, 10], "background": GuidParameter.PINK_RED, "foreground": GuidParameter.WHITE,  "font": ("Helvetica", 12), "shape": "rectangle"}}})
style.theme_use("MyStyle")

# Tabs (<notebook)
update_tab = Frame(notebook)
view_tab = Frame(notebook)
web_tab = Frame(notebook)

notebook.add(update_tab, text='Update')
notebook.add(view_tab, text = 'View')
notebook.add(web_tab, text="Web")
notebook.pack(expand = 1, fill = BOTH)

# UPDATE TAB:
update_tab_top_frame = TopFrame(update_tab)

update_tab_bottom_frame = BottomFrame(update_tab)

update_tab_left_frame = LeftFrame(root, update_tab_bottom_frame.frame, "update_tab")

update_tab_middle_frame = MiddleFrame(root, update_tab_bottom_frame.frame)

update_tab_left_frame.set_middle_frame(update_tab_middle_frame)

# VIEW TAB
view_tab_top_frame = TopFrame(view_tab)

view_tab_bottom_frame = BottomFrame(view_tab)

view_tab_left_frame = LeftFrame(root, view_tab_bottom_frame.frame, "view_tab")

view_tab_middle_frame = MiddleFrame(root, view_tab_bottom_frame.frame)

view_tab_left_frame.set_middle_frame(view_tab_middle_frame)

# WEB TAB
canvas = Canvas(web_tab, width=GuidParameter.SCREEN_WIDTH, height=GuidParameter.SCREEN_HEIGHT)

image =PhotoImage(file="images/semo_campus.png")
canvas.create_image(0, 0, image = image, anchor=NW)

font1 = Font(root, family="adobe-garamond-pro", size = 30, slant="italic", weight="bold")
canvas.create_text(GuidParameter.SCREEN_WIDTH/2, GuidParameter.SCREEN_HEIGHT/6, font=font1, text="Southeast Missouri State University", fill=GuidParameter.ORANGE)

font2 = Font(root, family="proxima-nova", size=42, weight="bold")
canvas.create_text(GuidParameter.SCREEN_WIDTH/2, GuidParameter.SCREEN_HEIGHT/6+80, font=font2, text="Center for Teaching and Learning", fill=GuidParameter.WHITE)

font3 = Font(root, family="arial black", size=32, weight="bold")
canvas.create_text(GuidParameter.SCREEN_WIDTH/2, GuidParameter.SCREEN_HEIGHT/6+170, font=font3, text="WHD MONTHLY TICKET REPORT", fill=GuidParameter.BEAUTIFUL_RED)

btn = Button(
	web_tab, 
	text = 'WEB VIEW',
	command= openbrowser,
	width=10,
	height=1,
	relief=SOLID,
	font=('arial', 22)
)

btn_canvas = canvas.create_window(
        GuidParameter.SCREEN_WIDTH/2-40, 
	GuidParameter.SCREEN_HEIGHT/6+250,
	anchor = "center",
	window = btn,
	)

canvas.grid(row=0, column=0)

root.mainloop()  



