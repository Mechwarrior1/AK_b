###################### GLOBAL SETTINGS ##########################
auto_refill = True #change to true for auto refil

#################################################################

import subprocess
import sys

try:
    import numpy as np
except ImportError:
    print('Installing module: numpy')
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'numpy']) #install python modules
finally:
    import numpy as np

try:
    import cv2
except ImportError:
    print('Installing module: opencv-python')
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'opencv-python'])
finally:
    import cv2

try:
    from PIL import ImageGrab
except ImportError:
    print('Installing module: pillow')
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'pillow'])
finally:
    from PIL import ImageGrab

try:
    from matplotlib import pyplot as plt
except ImportError:
    print('Installing module: matplotlib')
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'matplotlib'])
finally:
    from matplotlib import pyplot as plt

try:
    import pywintypes
    import win32gui, win32api, win32con
except ImportError:
    print('Installing module: pywin32')
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'pywin32'])
finally:
    import pywintypes
    import win32gui, win32api, win32con

print('Modules imported')
import pathlib, os, time
path=pathlib.Path().absolute()

os.chdir(path)
print('Current path of file: ',path)

assets_stage = {
    '1_start_mission_blustartbutton_san':1,
    '1_start_mission_autooff':1,
    '1_start_mission_autoon':1,
    '1_start_mission_blustartbutton':1,
    '1_start_mission_refill':1,
    '2_start_mission_ms':2,
    '3_mission_ongoing_gear':3,
    '3_mission_ongoing_takeover':3,
    '4_mission_end_exp':4,
    '4_mission_end_lmd':4,
    }

# for i in assets_stage.keys(): #for when required to rebuild assets
#     template = plt.imread(f'{i}.png')
#     height, width, _ = template.shape
#     center = [int(width/2), int(height/2)]
#     # click_area = [int(center[0]*0.85),int(center[1]*0.85)]
#     width_click = (int(width*0.1),int(width*0.9))
#     height_click = (int(height*0.2),int(height*0.9))
#     assets_specs[i]=(width, height, center, width_click, height_click)
#     print(i,': ',width, ',',height,',', center,',', width_click, ',', height_click)

assets_specs={ #(width, height, center, width_click, height_click)
    '1_start_mission_blustartbutton_san': (274, 67, [137, 33], (27, 246), (13, 60)),
    '1_start_mission_autooff': (261, 47, [130, 23], (26, 234), (9, 42)),
    '1_start_mission_autoon': (269, 43, [134, 21], (26, 242), (8, 38)),
    '1_start_mission_blustartbutton': (275, 67, [137, 33], (27, 247), (13, 60)),
    '1_start_mission_refill': (87, 75, [43, 37], (8, 78), (15, 67)),
    '2_start_mission_ms': (158, 329, [79, 164], (15, 142), (65, 296)),
    '3_mission_ongoing_gear': (61, 48, [30, 24], (6, 54), (9, 43)),
    '3_mission_ongoing_takeover': (183, 48, [91, 24], (18, 164), (9, 43)),
    '4_mission_end_exp': (95, 66, [47, 33], (9, 85), (13, 59)),
    '4_mission_end_lmd': (73, 51, [36, 25], (7, 65), (10, 45))
}

DEFAULT_SIMILARITY = 0.85

def update_screen():
    printscreen_pil =  ImageGrab.grab()
    printscreen_numpy =   np.array(printscreen_pil,dtype='uint8')
    # frame = cv2.cvtColor(printscreen_numpy, cv2.COLOR_BGR2RGB)
    frame = cv2.cvtColor(printscreen_numpy, cv2.COLOR_BGR2GRAY)
    return frame

#plt.imshow(aa,cmap='gray')

def find(image, similarity=None): #uses non boxed update_screen
        template = cv2.cv2.imread(f'{image}.png')
        template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        match = cv2.matchTemplate(update_screen(), template, cv2.TM_CCOEFF_NORMED)
        _, match_acc, _, location,= cv2.minMaxLoc(match)
        if similarity==None:
            return match_acc, location
        else:
            return similarity<match_acc

class click_control:
    def __init__(self,similarity=DEFAULT_SIMILARITY):
        self.similarity=similarity
        self.scr_width, self.scr_height = self.getscreensize()
        self.nox_screen()
        self.p=np.random.uniform(0,0.5)
    
    def getscreensize(self):
        pos = win32gui.GetCursorInfo()
        win32api.SetCursorPos((10000,5000))
        _, _, (width, height) = win32gui.GetCursorInfo()
        win32api.SetCursorPos((pos[2][0],pos[2][1]))
        return width+1, height+1
    
    def click_xy(self,x,y):
        win32api.SetCursorPos((x,y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
        if np.random.choice([0,1],p=[0.85,0.15])==1:
            time.sleep(np.random.uniform(0.090,0.120))
        else:
            time.sleep(np.random.uniform(0.110,0.150)) #15% chance of clicking slower
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
    
    def drag_xy(self,move_x=100,move_y=100,speed=0.04): #up, down, left, right
        win32api.SetCursorPos((int(np.random.randint(700,900)),int(np.random.randint(400,500))))
        flags, hcursor, cpos = win32gui.GetCursorInfo()
        x=(move_x)/abs(move_x)
        y=(move_y)/abs(move_x)
        win32api.SetCursorPos(cpos)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        for i in range(0,int(abs(move_x)/5)):
            win32api.SetCursorPos((int(cpos[0]+x*5*i),int(cpos[1]+y*5*i)))
            time.sleep(np.random.uniform(speed/10,speed/10+speed/110))
        #win32api.SetCursorPos((int(cpos[0]+x*move+x*random.randint(0,5)),int(cpos[1]+y*move+y*random.randint(0,5))))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
    
    def click_ranbox(self,click_again=False): #randomly clicks on nox screen
        if click_again: #click again on same around location
            if np.random.randint(0,1) == 0:
                self.click_x += np.random.randint(-5,5)
                self.click_y += np.random.randint(-5,5)
        else:
            self.click_x = np.random.randint(self.top_x,self.bot_x)
            self.click_y = np.random.randint(self.top_y,self.bot_y)
        self.click_xy(self.click_x, self.click_y)
    
    def click_ranbox_alt(self,location, width_click, height_click, click_again=False): #default nox screen range
        # for alternate easy input
        x_loc, y_loc = location
        if click_again: #click again on same around location
            if np.random.choice([0,1],p=[0.85,0.15])==1: #15% chance of moving click area
                self.click_x += np.random.randint(-5,5)
                self.click_y += np.random.randint(-5,5)
        else:
            self.click_x = np.random.randint(width_click[0],width_click[1]) + x_loc
            self.click_y = np.random.randint(height_click[0],height_click[1]) + y_loc
        self.click_xy(self.click_x, self.click_y)
    
    def nox_screen(self):
        nox_1_info = find('nox_1')
        nox_2_info = find('nox_2')
        if nox_1_info[0]>0.8: #nox icon is found or not
            self.top_x, self.top_y = nox_1_info[1]
            self.top_x += 200
            self.top_y += 120
        else:
            print('possible error found, nox top icon is not found, screen clicking area is not cailbrated')
            self.top_x = 200
            self.top_y = 120
        if nox_2_info[0]>0.8:
            self.bot_x, self.bot_y = nox_2_info[1]
            self.bot_x -= 200
            self.bot_y -= 350
        else:
            print('possible error found, nox bottom screen switcher icon is not found, screen clicking area is not cailbrated')
            self.bot_x = 1600
            self.bot_y = 650
    
    def find_click(self, image, similarity=DEFAULT_SIMILARITY,double_clicks=True,assets_specs=assets_specs,click_again=False):
        # assets_specs = (width, height, center, width_click, height_click)
        match_acc, location = find(image)
        if match_acc > similarity: 
            width, height, center, width_click, height_click = assets_specs[image]
            self.click_ranbox_alt(location = location, width_click = width_click, height_click = height_click,click_again=click_again)
            if (np.random.choice([0,1],p=[0.3,0.7])==1) and double_clicks:
                time.sleep(np.random.uniform(0.085,0.110))
                self.click_ranbox_alt(location = location, width_click = width_click, height_click = height_click,click_again=True)
            return True
        else:
            return False

def refill_sanity():
    if auto_refill:
        refilled_=click_control1.find_click('1_start_mission_refill')
    else:
        ans = input('Refill? y/n: ')
        while True:
            if ans =='y':
                print('attempting to refill sanity') #1_start_mission_blustartbutton_san
                for _ in range(3):
                    refilled_= click_control1.find_click('1_start_mission_refill')
                    if refilled_:
                        print('Refilled sanity')
                    time.sleep(1)
            elif ans =='n':
                return 'stop'
            elif find('1_start_mission_autoon',0.8):
                break
            else:
                ans = input('Refill? please type only y/n: ')
            

def start_1_loop(N=10,afk='y',error_in=False):
    click_again=False
    #error_in means it skiped stages and entered after error finding
    print('Currently at stage 1, attempting to enter map')
    for i in range(N): #loop for clicking stage auto button
        check_autooff = click_control1.find_click('1_start_mission_autooff',click_again=click_again)
        click_again=True
        check_autoon = find('1_start_mission_autoon',0.8)
        if find('1_start_mission_autolock',0.9):
            print('Auto deploy seems to be locked, try another stage')
            ans=input('after going to another map which you can auto, please hit any "y" to restart operation or "n" to stop: ')
            while True:
                if ans =="y":
                    break
                elif ans =="n":
                    return 'stop'
        elif not(check_autooff) and (check_autoon):
            break
        elif find('1_start_mission_refill',0.92):
            refill_sanity()
        elif find('2_start_mission_ms',0.85) and (i > 2): #break if
            break
        time.sleep(np.random.uniform(0.4,0.6))
    else:
        print('auto buttons not found/ has error finding/ error clicking auto, check assets or resolution of game client NOX')
        return False
    click_again=False
    for i in range(N): #loop for clicking blue start button
        check_start1 = click_control1.find_click('1_start_mission_blustartbutton',0.85,double_clicks=False,click_again=click_again)
        click_again=True
        time.sleep(np.random.uniform(1.5,2.1))
        ## refill loop
        if find('1_start_mission_refill', 0.9):
            refill_sanity()
        elif not(check_start1):
            click_control1.find_click('1_start_mission_blustartbutton_san',0.85,double_clicks=False,click_again=click_again)
        elif find('2_start_mission_ms',0.85): #found the next stage
            return True
    else:
        print('bot was unable to find the blue start button')
        return False



def start_2_loop(N=15,afk='y',error_in=False):
    click_again=False
    #error_in means it skiped stages and entered after error finding
    print('Currently at stage 2, attempting to enter combat')
    for i in range(N):
        check_start2 = click_control1.find_click('2_start_mission_ms',click_again=click_again)
        click_again=True
        if find('3_mission_ongoing_takeover',0.8) or find('3_mission_ongoing_gear',0.8): #found the next stage
            return True
        time.sleep(np.random.uniform(0.4,0.6))
    else:
        print('bot was unable to find the the red misson start')
        return False

def start_3_loop(N=35,afk='y',error_in=False):
    #error_in means it skiped stages and entered after error finding
    print('Currently at stage 3, should be in combat or loading screen')
    click_again=False
    def wait_battle(initial_wait=True,N=35):
        if initial_wait:
            print('sleep for 1 min')
            time.sleep(60) #sleep for 1 min
        for i in range(N):
            check_ongoing = find('3_mission_ongoing_takeover', 0.8)
            if find('4_mission_end_exp', 0.85) or find('4_mission_end_lmd', 0.85): #found the next stage
                return True
            time.sleep(3) #check every 3 seconds
            if (i%10==0) and check_ongoing:
                print('Battle ongoing ',i)
    
    if not(error_in):
        wait_battle()
    # if find('3_mission_ongoing_takeover', 0.85):
    #     wait_battle(initial_wait=False,N=N)
    # if find('3_mission_ongoing_takeover', 0.85):
    #     wait_battle(initial_wait=False,N=N)
    wait_battle(initial_wait=False,N=N)
    wait_battle(initial_wait=False,N=N)
    if find('4_mission_end_exp', 0.85) or find('4_mission_end_lmd', 0.85):
        return True
    else:
        print('bot seems to have stalled while waiting for battle to finish')
        return False

def start_4_loop(N=15,afk='y',error_in=False): 
    #error_in means it skiped stages and entered after error finding
    click_again=False
    for i in range(N):
        check_ended = find('4_mission_end_exp', 0.85) or find('4_mission_end_lmd', 0.85)
        if check_ended: #only click when it ends
            afk_factor=np.random.choice([1,2,3],p=[0.9,0.075,0.025])
            if afk_factor==1 or afk=='n':
                afk_time=np.random.uniform(3,5)
            elif afk_factor==2:
                afk_time=np.random.uniform(15,35)
            elif afk_factor==3:
                afk_time=np.random.uniform(55,185)
            print('Currently at ending screen, *note that the bot does not check result')
            print(f'AFK timer: {afk_time:.0f} seconds,  *2.5% chance of 1.5 mins to 3 mins')
            time.sleep(afk_time)
            click_control1.click_ranbox(click_again=click_again)
            click_again=True
        if find('1_start_mission_blustartbutton', 0.85) or  find('1_start_mission_blustartbutton_san', 0.85): #found the next stage
            return True
        time.sleep(2) #check every 3 seconds
    else:
        return False

loops_stages={
    1:start_1_loop,
    2:start_2_loop,
    3:start_3_loop,
    4:start_4_loop
}

assets_stage = {
    '1_start_mission_blustartbutton_san':1,
    '1_start_mission_autooff':1,
    '1_start_mission_autoon':1,
    '1_start_mission_blustartbutton':1,
    '1_start_mission_refill':1,
    '2_start_mission_ms':2,
    '3_mission_ongoing_gear':3,
    '3_mission_ongoing_takeover':3,
    '4_mission_end_exp':4,
    '4_mission_end_lmd':4,
    }

def error_finder():
    for img in list(assets_specs.keys()):
        found_clue=find(img, 0.9)
        if found_clue:
            stage=assets_stage[img]
            return stage
    else:
        input('The bot has encounted a screen it has not seen before, terminating bot')
        return False

def runs(N,afk):
    runs_total=0
    error_in=False
    for i in range(N):
        stage=1
        while True:
            stage_result=loops_stages[stage](afk=afk,error_in=error_in)
            error_in=False
            if stage_result=='stop':
                print('stop loop, a stop signal was received')
                return 'stop'
            elif stage_result==True:
                stage += 1
            elif stage_result==False:
                error_in=True
                print('An error has been detected, do wait while the bot checks the environment')
                stage = error_finder()
                if stage==False:
                    return False
            if stage_result and (stage==5):
                runs_total += 1
                print(f'Run completed, total runs completed in this instance: {runs_total}')
                break


def main():
    print()
    print('###############################  New Instance  ##################################')
    ans = input('How many rounds to run? please key in a number, or "stop" to stop: ')
    while True:
        try:
            if ans=='stop':
                return False
            N = int(ans)
            break
        except:
            ans = input('There seems to be an error in your input, please key in the number of runs you want: ')
    global afk
    if afk==None:
        afk = input('Do you want any rng chance for slightly longer afk timings in between runs (*2.5% chance of max 3 mins) y/n: ')
        while True:
            if (afk =='y') or (afk=='n'):
                break
            else:
                afk = input('There seems to be an error in your input, please type y/n for chance of longer afk timings (max 3mins):  ')
    result = runs(N,afk)
    if result==False:
        return False
    return True


if __name__ =='__main__':
    try:
        global afk
        afk = None
        while True:
            click_control1=click_control()
            aa = main()
            if aa== False:
                break
        input('press any key to exit, or just close python')
    except:
        input('an error seems to have occurred')

