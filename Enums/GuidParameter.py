import os


class GuidParameter:
    # screen
    SCREEN_HEIGHT = 770
    SCREEN_WIDTH = 1620
    # image
    SEMO_LOGO_IMAGE_FILE = "c:/Users/Tuyen/source/repos/CSTL_Ticket_Report_Database_Update/images/semo_logo.PNG"
    SEMO_CAMPUS_IMAGE_FILE = "c:/Users/Tuyen/source/repos/CSTL_Ticket_Report_Database_Update/images/semo_campus.png"
    CATEGORY_ICON_FILE = "c:/Users/Tuyen/source/repos/CSTL_Ticket_Report_Database_Update/images/category.png"

    # color
    PINK_RED = "#9d2235"
    BLACK = "black"
    WHITE = "white"
    GRAY1 = '#E8E8E8'
    GRAY2 = '#A9A9A9'
    RED = "red"
    BLUE = "#6ab0de"
    BEAUTIFUL_RED = "#FE3912"
    ORANGE = "#C8102E"
    

    # color used for .xlsx file
    GREEN ="eeffcc"
    YELLOW = "FFFF00"
    ORANGE_BROWN = "FF9900"
    BLUE_FOR_BAR = "6ab0de"
    LIGHT_BLUE_FOR_BG = "DAE3F3"
    GRAY = "595959"

    # Left_frame_pad_x
    LEFT_FRAME_PAX_X = [10, 10]

print(os.path.dirname(__file__) )