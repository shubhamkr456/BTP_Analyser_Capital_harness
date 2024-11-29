import pyperclip
import time
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.mouse import click
from pywinauto import findwindows, Desktop
import pyautogui
import keyboard
import gc
import pandas as pd



#--------variable declaration---------------------
WireID_paster= (703, 800)   # for WireID entering Purpose
add_button= (1717, 585)     # Add button for properties
edit_note= (1044, 517)      # delete the default- text 
apply=(1721, 834)           
length= (1307, 361)         # enter numerial
absolute= (1319, 338)       # select absolute
# Path where the analysed excel has been placed----
path = r"C:\Users\vu378e\Desktop\C17 App\Capital_Files"

###--------------------------------------------------
def convert_str_input(string: str):
    new= ""
    for c in str(string):
        if c in ['(', ')']:
            new+= ('{' + c + '}')
        else:
            new += c
        return new
            


# image location on screen using pyautogui for image recognition
def locate_and_click(image_path):
    res = pyautogui.locateCenterOnScreen(image_path)
    if res:
        pyautogui.moveTo(res)
        pyautogui.click()
    else:
        print(f"Image {image_path} not found on screen.")


def HarnessWinSelector():
    locate_and_click("harnessApp.jpg")


def HarnessWireSelector():
    locate_and_click("harnessNumber.png")


def HarnessADDSelector():
    pyautogui.moveTo(add_button)
    pyautogui.click()


def WireIDpaster():
    pyautogui.moveTo(WireID_paster)
    pyautogui.click()
    send_keys("{BACKSPACE}" * 30, with_spaces=True)
    time.sleep(0.1)
    send_keys("^v")  # Paste the clipboard content


def curmover(tupl):
    pyautogui.moveTo(tupl)
    time.sleep(0.5)
    pyautogui.click()


def keyboard_tab_press():
    while True:
        if keyboard.is_pressed("tab"):
            break


def extract_integer_from_notex(notex):
    # Check if 'WL' is in the string
    index = notex.find('WL')
    if index != -1:
        # Grab the next 2 characters after 'WL'
        next_two_chars = notex[index + 2:index + 4]
        try:
            # Convert to integer
            result = int(next_two_chars)
            return result
        except ValueError:
            # Handle cases where conversion to integer fails
            print(f"Could not convert '{next_two_chars}' to integer.")
            return None
    else:
        print("'WL' not found in the string.")
        return None



z = input("Enter File name with extension")
file = path + "\\" + z
# Main Execution
variable_value = "W1262-29-24"
df = pd.read_excel(
    file,
    dtype=str, sheet_name= "Notes&Length"
)

df.fillna("", inplace=True)

# Select the Harnes Window
time.sleep(8)
#HarnessWinSelector()


for index, row in df.iterrows():

    # Wire ID is selcted
    variable_value = row["WIRE ID"]
    pyperclip.copy(variable_value)
    note_ = row["note_add"]
    note_text = row["note_Text1"]
    note_1 = row["note_add2"]
    note_text1 = row["note_Text2"]
    length_up= row["Length_Update"]
            
#----------------------------------
    time.sleep(0.1)
    screen_width, screen_height = pyautogui.size()
    pyautogui.moveTo(screen_width // 2, screen_height // 3)
    pyautogui.click()
    ## Type the wireID
    WireIDpaster()
    
    if note_ !=  "":
        HarnessADDSelector()
        # This selects the type box -----edit_note
        curmover(edit_note)
        time.sleep(1.0)  # 1019, 521
        pyautogui.click()
        # Delete the selected text using the Backspace key
        send_keys("{BACKSPACE}" * 30, with_spaces=True)
        send_keys(note_ * 1)
        time.sleep(1.2)
        send_keys("{ENTER}", with_spaces=True)
        #note_text= convert_str_input(note_text) 
        send_keys(note_text * 1, with_spaces=True)
        print(note_)
        print(note_text)

    # note 2  entering
    if note_1 != "":

        HarnessADDSelector()
        curmover(edit_note)  # edit_note variable
        time.sleep(1)  # 1019, 521
        pyautogui.click()
        # Now delete it
        # Delete the selected text using the Backspace key
        send_keys("{BACKSPACE}" * 30, with_spaces=True)
        send_keys(note_1 * 1)
        time.sleep(1.2)
        send_keys("{ENTER}", with_spaces=True)
        #note_text1= convert_str_input(note_text1)
        send_keys(note_text1 * 1, with_spaces=True)

      
    if (length_up== "Different"):
        if any(code in variable_value for code in ("SH", "BL", "WH", "BA")):
        #if "SH" in variable_value or "BL" in variable_value or "WH" in variable_value or "BA" in variable_value:
            pass
        else:
            # Length Update
            curmover(length)
            send_keys("{BACKSPACE}" * 5, with_spaces=True)
            send_keys(len_val * 1, with_spaces=True)
            curmover(absolute) # absolute drop down
            send_keys("A" * 1, with_spaces=True)
                
    keyboard_tab_press()
    curmover(apply)  # apply

gc.collect()

