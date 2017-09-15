import os
import sys
import time
import webbrowser

# Script requires pywinauto (python -m pip install pywinauto)
from pywinauto import findwindows
from pywinauto.application import Application

# This script was tested/working with Python 2.7.12 and Outlook 2016

outlook_title = '.* - Outlook'
calendar_save_full_name = r'C:\calendar.ics'
ms_calendar_string = r'"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /select Outlook:Calendar'
calendar_window_title = 'Calendar - .* - Outlook'

#Application().start(r'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --new_window https://calendar.google.com/calendar/render#settings-calendars_9')

def connect_to_window(title_re):
    handle = findwindows.find_window(title_re = title_re, visible_only = False)
    return Application().connect(handle=handle)

#"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /select Outlook:Calendar

def save_calendar_ics():
    # Connect to Outlook
    outlook = connect_to_window(outlook_title)
    # print outlook.__dict__
    win = outlook.rctrl_renwnd32
    # print win.print_control_identifiers()
    #print win.dump_tree(depth = 2)

    #TODO: Assert no dialog boxes are open
    
    # ^:Ctrl, %:Alt, +:Shift
    # Switch to calendar view
    #win.type_keys('^2') # Ctrl-2 goes to Calendar
    # This is what I'd like to do, but doesn't work:
    # win.AwesomeBar.Calendar.click()
    # This isn't as good, but works:
    win.AwesomeBar.type_keys('^2')

    # Open the save calendar dialog box
    # This is not working:
    #win.MsoCommandBarDock.MsoCommandBar.File.SaveCalendar.click()
    # Again, ugly, but works:
    # win.MsoCommandBar.type_keys('%f') # File
    # win.MsoCommandBar.type_keys('c') # Save Calendar
    win.MsoDockTop.type_keys('%fc') # File -> Save Calendar

    #win.dump_tree(depth=1)
    #print outlook.SaveAsDialog.print_control_identifiers()

    # # Enter the filename to save
    # outlook.SaveAsDialog.ComboBox0.Edit.set_text(calendar_save_full_name)

    # As soon as I click more options, I lose the full path, so do this stuff
    # before entering filename

    outlook.SaveAsDialog.MoreOptionsButton.click()

    # Select the whole calendar
    #outlook.SaveAsDialog.type_keys('{DOWN 4}')
    # outlook.SaveAs.ComboBox.Select('Whole calendar')
    # This is a hack to wait for the secondary dialog box to appear
    # (unfortunately, the dialog shares a name with its parent)
    while True:
        try:
            outlook.SaveAs.ComboBox.Select('Whole calendar')
            break
        except ValueError:
            pass
    # Select limited details
    #outlook.SaveAsDialog.type_keys('{TAB}{DOWN}')
    outlook.SaveAs.ComboBox2.Select('Limited details')
    #outlook.SaveAs.ComboBox2.Select('Full details')
    # Click OK
    #outlook.SaveAsDialog.OkButton.click() < This clicks the Show button for some reason
    #outlook.SaveAsDialog.Ok.click()
    outlook.SaveAs.Ok.click()
    # Accept the limited details message prompt
    outlook.MicrosoftOutlookDialog.Yes.click()

    # Enter the filename to save
    outlook.SaveAsDialog.ComboBox0.Edit.set_text(calendar_save_full_name)

    # Save it!
    outlook.SaveAsDialog.Save.click()
    # Yes, overwrite the file
    outlook.Confirm.Yes.click()

    # Go back to email view
    win.AwesomeBar.type_keys('^1')

def save_calendar_ics2():
    # Start a new instance of Outlook with calendar
    outlook = Application().start(ms_calendar_string, timeout=5)
    # This doesn't work:
    #win = outlook.rctrl_renwnd32
    
    # Need to allow window to open before connecting
    time.sleep(1)
    
    # Not sure why I need to connect to the window here
    outlook = connect_to_window(calendar_window_title)
    
    win = outlook.rctrl_renwnd32

    win.MsoDockTop.type_keys('%fc') # File -> Save Calendar

    outlook.SaveAsDialog.MoreOptionsButton.click()

    # Select the whole calendar
    #outlook.SaveAsDialog.type_keys('{DOWN 4}')
    # outlook.SaveAs.ComboBox.Select('Whole calendar')
    # This is a hack to wait for the secondary dialog box to appear
    # (unfortunately, the dialog shares a name with its parent)
    while True:
        try:
            outlook.SaveAs.ComboBox.Select('Whole calendar')
            break
        except ValueError:
            pass
    # Select limited details
    #outlook.SaveAsDialog.type_keys('{TAB}{DOWN}')
    outlook.SaveAs.ComboBox2.Select('Limited details')
    # Click OK
    #outlook.SaveAsDialog.OkButton.click() < This clicks the Show button for some reason
    #outlook.SaveAsDialog.Ok.click()
    outlook.SaveAs.Ok.click()
    # Accept the limited details message prompt
    outlook.MicrosoftOutlookDialog.Yes.click()

    # Enter the filename to save
    outlook.SaveAsDialog.ComboBox0.Edit.set_text(calendar_save_full_name)

    # Save it!
    outlook.SaveAsDialog.Save.click()
    # Yes, overwrite the file
    outlook.Confirm.Yes.click()

    # Close the window
    win.close()

if __name__ == "__main__":
    save_calendar_ics2()
