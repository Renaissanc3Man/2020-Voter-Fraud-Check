import sys
sys.path.append(".")
from corona_library import *

#import pyscreenshot as ImageGrab
import pyautogui



import win32api, win32con
def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

#click(10,10)


from pynput.keyboard import Key, Controller

keyboard = Controller()


def type_string(mystring):
    for myletter in list(mystring):
        keyboard.press(myletter)
        keyboard.release(myletter)
        time.sleep(0.005)




michigan_voter_website = r'https://mvic.sos.state.mi.us/Voter/Index'







voter_df = pd.read_csv(r"C:\__YOUTUBE__\__POLITICS__\Voterfraud Check\voterfraud_list.csv",low_memory=False)
voter_df = voter_df[voter_df['Birth Year'] != 1900]
voter_df = voter_df[voter_df['Ballot Received Date'] != 'NONE']
voter_df = voter_df.sort_values(['Age today'],ascending=False)
voter_df = safe_reset_index(voter_df)


delaytime = 0.15

#click on chrome
click(10,10)

print('starting')
multiple_tabs = 21

for myrow1 in [0]:
    #it takes so long to load, do multiple at same time
    for mynextrow in range(0,multiple_tabs):
        myrow = mynextrow + myrow1

        #new tab
        with keyboard.pressed(Key.ctrl):
            keyboard.press("t")
            keyboard.release("t")

        time.sleep(delaytime)




        #click http bar
        click(900,50)
        type_string(michigan_voter_website)
        time.sleep(delaytime)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)

        time.sleep(0.5)


    time.sleep(2)
    #go back to first tab
    for mynextrow in range(1,multiple_tabs):
        with keyboard.pressed(Key.ctrl):
            with keyboard.pressed(Key.shift):
                keyboard.press(Key.tab)
                keyboard.release(Key.tab)
        time.sleep(2)


    time.sleep(30)

    for mynextrow in range(0,multiple_tabs):
        myrow = mynextrow + myrow1



        #click first name slot
        click(900,775)

        #move mouse out of way so it doesnt interfere w the dropdown menu
        win32api.SetCursorPos((1500,775))


        #get first and last names
        name = str(voter_df.iloc[myrow]['Name'])
        name_split = name.split(' ')

        last_name = name_split[0]

        #handle if has a middle name
        if len(name_split) > 2:
            first_name = name_split[1:].join(' ')
        else:
            first_name = name_split[1]

        type_string(first_name)

        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)

        type_string(last_name)
        time.sleep(delaytime)

        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)

        #use dropdown menu to select birthmonth
        keyboard.press(Key.space)
        keyboard.release(Key.space)
        time.sleep(delaytime)


        birth_month = int(voter_df.iloc[myrow]['Birth Month'])
        print('birth_month',birth_month)
        for monthdown in range(0,birth_month):
            keyboard.press(Key.down)
            keyboard.release(Key.down)
            time.sleep(delaytime)

        #tab to end
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)

        #tab to next
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)



        #enter birthyear
        type_string(str(voter_df.iloc[myrow]['Birth Year']))
        time.sleep(delaytime)

        #tab to next
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)

        #enter zipcode
        type_string(str(voter_df.iloc[myrow]['Zip Code']))
        time.sleep(delaytime)

        #tab to next
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)

        #tab to search
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(delaytime)


        #press enter to search
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)


        #take a screenshot and save it
        name = str(voter_df.iloc[myrow]['Name'])
        picture_filename = r'C:\__YOUTUBE__\__POLITICS__\Voterfraud Check\output_data\screen_shots\\' + str(myrow) + '_' + name + '_search.png'
        try:
            #im=ImageGrab.grab(bbox=(10,10,1800,1050))
            im = pyautogui.screenshot(picture_filename,region=(10,30, 1800,1050))
            #im.save(picture_filename)
        except:
            print('error taking picture')

        voter_df.loc[myrow, ('Voter Search Screenshot')] = picture_filename
        time.sleep(5)


        #dont go to next tab on last iteration
        if mynextrow != multiple_tabs - 1:
            #go to next tab
            with keyboard.pressed(Key.ctrl):
                keyboard.press(Key.tab)
                keyboard.release(Key.tab)
        time.sleep(1)







    #go back to first tab
    for mynextrow in range(1,multiple_tabs):
        with keyboard.pressed(Key.ctrl):
            with keyboard.pressed(Key.shift):
                keyboard.press(Key.tab)
                keyboard.release(Key.tab)
        time.sleep(5)

    for mynextrow in range(1,multiple_tabs):
        with keyboard.pressed(Key.ctrl):
            keyboard.press(Key.tab)
            keyboard.release(Key.tab)
        time.sleep(5)

    for mynextrow in range(1,multiple_tabs):
        with keyboard.pressed(Key.ctrl):
            with keyboard.pressed(Key.shift):
                keyboard.press(Key.tab)
                keyboard.release(Key.tab)
        time.sleep(5)



    print('loading pages')




    #wait for search results
    #time.sleep(60)

    for mynextrow in range(0,multiple_tabs):
        print(mynextrow)
        myrow = mynextrow + myrow1
        #go through loaded pages and take screenshots

##        #delay to manually check if loaded
##        click(-1052,10)
##        input("Press Enter to continue...")


        #scroll up
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)
        click(1843,76)



        #get ballot received date
        click(1843,76)

        time.sleep(1)


        #take a screenshot and save it
        name = str(voter_df.iloc[myrow]['Name'])
        picture_filename = r'C:\__YOUTUBE__\__POLITICS__\Voterfraud Check\output_data\screen_shots\\' + str(myrow) + '_' + name + '_result.png'
        try:
            #im=ImageGrab.grab(bbox=(10,10,1800,1050))
            im = pyautogui.screenshot(picture_filename,region=(10,30, 1800,1050))
            #im.save(picture_filename)
        except:
            print('error taking picture')

        voter_df.loc[myrow, ('Voter Result Screenshot')] = picture_filename

        #dont go to next tab on last iteration
        if mynextrow != multiple_tabs - 1:
            #go to next tab
            with keyboard.pressed(Key.ctrl):
                keyboard.press(Key.tab)
                keyboard.release(Key.tab)
        time.sleep(1)


print('finished')












