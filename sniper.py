import os
import sys
import time
import pytz
import mouse
import msvcrt
import keyboard
from PIL import Image
from random import random
from itertools import product
from datetime import datetime as dt



DEFAULT_CURSOR_POSITION = (739, 296)
""" The position which the mouse should be in order to click an item appearing in the TSM item region. """










def delay(duration: int) -> None:
    """
    Halts the program for the specified duration.

    Parameters
    -
        `duration`:   The duration,  in seconds,  to halt the program for.
    """
    from time import perf_counter
    start = perf_counter()
    while perf_counter() - start < duration:
        pass



def now() -> str:
    """
    Returns the current date and time formatted as a string,
    in the `US/Central` timezone, with the form `09-11-2022 09:47:52`.
    """
    return dt.now(pytz.timezone('US/Central')).strftime("%m-%d-%Y %H:%M:%S")



def seconds_since(dtime: dt) -> int:
    """
    Returns the number of seconds that have elapsed
    between now and the time of the given `datetime` object.

    Parameters
    -
        `dtime`:   The `datetime` object to measure elapsed time since.
    """
    return int((dt.now() - dtime).total_seconds())



def get_users_money(rtype = None) -> list | str:
    """
    Gets the user's money via user input.

    Done at the start of the program and at the end, to calculate how much money the bot spent.

    Parameters
    -
        `rtype`:   The type to return the money as.  Can be left as `None`, which will return a `list`
                         of the form `[gold, silver, copper]`,  or it can be given any value to return a `str`.
    """
    money = input(
    """
    Separating gold, silver, & copper with periods,
    enter your player's current money reserves: >>>  """
    ).split('.')
    money = [int(i.strip().rstrip()) for i in money]
    return money if rtype is None else f"{money[0]}g {money[1]}s {money[2]}c"















class Log:
    """
    A class for writing data to a log file.

    Parameters
    -
        `filename`:   The name of the log file to write to.
    """
    def __init__(self, filename: str):
        self.filename = filename
        self.file = open(filename, 'a')



    def write(self, text: str) -> None:
        """
        Writes the given text to the log file.

        Parameters
        -
            `text`:   The text to write to the log file.
        """
        if self.file.closed:
            self.file = open(self.filename, 'a')
            self.file.write(text)
            self.file.close()
        elif self.file.mode == 'w' or self.file.mode == 'r':
            self.file.close()
            self.file = open(self.filename, 'a')
            self.file.write(text)
            self.file.close()
        return
    


    def read(self) -> str:
        """
        Returns the contents of the log file.
        """
        if self.file.closed:
            self.file = open(self.filename, 'r')
            contents = self.file.read()
            self.file.close()
            return contents
        elif self.file.mode == 'w' or self.file.mode == 'a':
            self.file.close()
            self.file = open(self.filename, 'r')
            contents = self.file.read()
            self.file.close()
            return contents















class Cursor:
    """
    Class for manipulating, and getting information about, the cursor.
    """
    def __init__(self):
        self._position = mouse.get_position()
    


    @property
    def position(self) -> tuple:
        """
        The current postion of the cursor.
        """
        return self._position
    
    @position.setter
    def position(self, value: tuple) -> None:
        if not isinstance(value, tuple):
            raise TypeError(f">> 'position' must be of type <{tuple.__name__}>, not <{type(value).__name__}>")
        self._position = value
    


    def update_position(self) -> None:
        """
        Updates the `position` attribute of the cursor with its current location.
        """
        self.position = mouse.get_position()
        return
    


    def move_to(self, xy: tuple) -> None:
        """
        Moves the cursor to the given `xy` coordinate.

        Makes a call to `self.update_position()` to update the `cursor.position` attribute after the move.
        
        Parameters
        -
            `xy`:   The position to move the cursor to.
        """
        mouse.move(*xy)
        self.update_position()
        delay(0.01)
        return
    


    def left_click(self, xy: tuple = None) -> None:
        """
        Sends a left click.  Returns the cursor to the original position after the click,  if `xy` was given.
        
        Parameters
        -
            `xy`:   The position to click.  If not specified, the current position of the cursor is used.
        """
        if not xy:
            self.update_position()
            xy = self.position
            mouse.click(button='left')
        else:
            self.update_position()
            current_position = self.position
            mouse.move(*xy)
            mouse.click(button='left')
            mouse.move(*current_position)
        delay(0.01)
        return
    


    def reset_position(self) -> None:
        """
        Moves the cursor to the default position,
        as defined with `DEFAULT_CURSOR_POSITION`.
        """
        self.move_to(DEFAULT_CURSOR_POSITION)
        return















class Screen:
    """
    A class containing several methods for obtaining information from the screen.
    """
    def __init__(self):
        import win32api
        self.resolution = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
        self.center = self.resolution[0]//2, self.resolution[1]//2
        self.top_left = (0, 0)
        self.top_right = (self.resolution[0], 0)
        self.bottom_left = (0, self.resolution[1])
        self.bottom_right = (self.resolution[0], self.resolution[1])
    


    @staticmethod
    def pixel_color(x: int, y: int) -> tuple:
        """
        Gets the pixel color of the screen at the specified coordinates,  as a tuple of `(r, g, b)` values.
        
        Parameters
        -
            `x`:   The x coordinate to get the pixel color of.
            `y`:   The y coordinate to get the pixel color of.
        """
        from PIL import ImageGrab
        return ImageGrab.grab().getpixel((x, y))
    


    @staticmethod
    def region(top_left: tuple, bottom_right: tuple) -> Image.Image:
        """
        Gets an image of the region of the screen.

        Parameters
        -
            `top_left`:   The x, y coordinates of the top left corner of the region.
            `bottom_right`:   The x, y coordinates of the top right corner of the region.
        """
        from PIL import ImageGrab
        x1, y1, x2, y2 = top_left[0],top_left[1],bottom_right[0],bottom_right[1]
        return ImageGrab.grab(bbox=(x1, y1, x2, y2))



    @staticmethod    
    def extract_text(img: Image.Image) -> str:
        """
        Extracts text from an image, or region of the screen via `Screen.region()`.

        Parameters
        -
            `img`:   The image to extract text from.
        """
        import pytesseract
        pytesseract.pytesseract.tesseract_cmd = r'F:\Program Files\Tesseract-OCR\tesseract.exe'
        return pytesseract.image_to_string(img)















class Keyboard():
    """
    Class for executing keyboard actions.
    """
    @staticmethod
    def one():
        """
        Presses and releases the `1` key on the keyboard.

        Used to buyout items via the `/click TSMSniperBtn` macro,  located on action bar button 1 in-game.
        """
        keyboard.press('1')
        time.sleep(0.05*(1+random()))
        keyboard.release('1')
        time.sleep(0.05*(1+random()))
    


    @staticmethod
    def two():
        """
        Presses and releases the `2` key on the keyboard.

        Used to target the auctioneer via the `/target Eoch` macro,  located on action bar button 2 in-game.
        """
        keyboard.press('2')
        time.sleep(0.05*(1+random()))
        keyboard.release('2')
        time.sleep(0.05*(1+random()))
    


    @staticmethod
    def down():
        """
        Presses and releases the `Down Arrow` key on the keyboard.

        Used to interact with the auctioneer (i.e. open TSM) via the in-game `Interact with Target` keybind.
        """
        keyboard.press('down')
        time.sleep(0.05*(1+random()))
        keyboard.release('down')
        time.sleep(0.05*(1+random()))
    


    @staticmethod
    def b():
        """
        Presses and releases the `B` key on the keyboard.

        Used to open the player's bags, for the purpose of taking a screenshot of their gold.
        """
        keyboard.press('b')
        time.sleep(0.05*(1+random()))
        keyboard.release('b')
        time.sleep(0.05*(1+random()))
    


    @staticmethod
    def enter():
        """
        Presses and releases the `ENTER` key on the keyboard.
        """
        keyboard.press('enter')
        time.sleep(0.05*(1+random()))
        keyboard.release('enter')
        time.sleep(0.05*(1+random()))
    


    @staticmethod
    def camp():
        """
        Combines a series of key-presses and releases to execute the `/camp` command in chat.

        Sequence:  `ENTER` ➝ `/` ➝ `C` ➝ `A` ➝ `M` ➝ `P`.  Each key is separated by a short random delay.
        """
        Keyboard.enter()
        keyboard.press('/')
        time.sleep(0.05*(1+random()))
        keyboard.release('/')
        time.sleep(0.05*(1+random()))
        keyboard.press('c')
        time.sleep(0.05*(1+random()))
        keyboard.release('c')
        time.sleep(0.05*(1+random()))
        keyboard.press('a')
        time.sleep(0.05*(1+random()))
        keyboard.release('a')
        time.sleep(0.05*(1+random()))
        keyboard.press('m')
        time.sleep(0.05*(1+random()))
        keyboard.release('m')
        time.sleep(0.05*(1+random()))
        keyboard.press('p')
        time.sleep(0.05*(1+random()))
        keyboard.release('p')
        time.sleep(0.05*(1+random()))
    
    













class TSM:
    """
    Class for TSM-specific things.
    """
    def __init__(self):
        self._restart_scan_button = (1367, 233)
        self._run_buyout_sniper_button = (695, 264)
        self._item_region = self.update_item_region()
        self._scan_status = self.get_scan_status()
        self._item_region_avg_color = (0, 0, 0)
        self._scan_status_region  = self.update_scan_status_region()
    


    @property
    def restart_scan_button(self) -> tuple:
        """
        The button for restarting the buyout scan,
        located in the upper-right, shown when the sniper scan is running.
        """
        return self._restart_scan_button
    


    @property
    def run_buyout_sniper_button(self) -> tuple:
        """
        The button for starting the sniper scan,
        shown when the TSM window is first opened.
        """
        return self._run_buyout_sniper_button
    


    @property
    def scan_status_region(self) -> Image.Image:
        """
        The region on the TSM pane located at the bottom-center
        that displays the current status of the scan (`'Running Sniper Scan'`, `'Scan Paused'`, etc).
        """
        return self._scan_status_region
    
    @scan_status_region.setter
    def scan_status_region(self, img: Image.Image) -> None:
        if not isinstance(img, Image.Image):
            raise TypeError(f">> 'scan_status_region' must be of type <{Image.Image.__name__}>, not <{type(img).__name__}>")
        self._scan_status_region = img
    


    @property
    def item_region(self) -> Image.Image:
        """
        The region on the TSM pane that displays the items found by the sniper scan.
        """
        return self._item_region
    
    @item_region.setter
    def item_region(self, img: Image.Image) -> None:
        if not isinstance(img, Image.Image):
            raise TypeError(f">> 'item_region' must be of type <{Image.Image.__name__}>, not <{type(img).__name__}>")
        self._item_region = img
    


    @property
    def item_region_avg_color(self) -> tuple:
        """
        The average  red,  green,  blue  color components of `tsm.item_region`.
        
        This property is updated every time there is a call to `self.update_item_region_avg_color()`.
        """
        return self._item_region_avg_color
    
    @item_region_avg_color.setter
    def item_region_avg_color(self, value: tuple) -> None:
        if not isinstance(value, tuple):
            raise TypeError(f">> 'item_region_avg_color' must be of type <{tuple.__name__}>, not <{type(value).__name__}>")
        self._item_region_avg_color = value
    


    @property
    def scan_running(self) -> bool:
        """
        Whether or not the sniper scanner is currently running.  Does a fresh check every time this property is accessed.

        Execution time:  `130ms`
        """
        return True if self.get_scan_status() == 'running' else False
    


    def update_item_region(self) -> None:
        """
        Gets a new image of the item region via `Screen.region()`,  then updates `self.item_region`.

        Execution time:  `30ms`
        """
        self.item_region = Screen.region((453,286), (1420,357))     # width = 967
        return



    def update_scan_status_region(self) -> None:
        """
        Gets a new image of the scan status region via `Screen.region()`,  then updates `self.scan_status_region`.

        Execution time:  `30ms`
        """
        self.scan_status_region = Screen.region((505,808), (1205,833))
        return
    


    def get_scan_status(self) -> str:
        """
        Returns the current status of the sniper scan (in lowercase).  Can be any of:  `buy`, `confirming`, `finding`, `running`, `paused`, or `unknown`.

        Done by taking a screenshot of the scan status region via `self.update_scan_status_region()`,  then extracting the text via `Screen.extract_text()`.
        
        Execution time:  `130ms`
        """
        self.update_scan_status_region()
        raw = Screen.extract_text(self.scan_status_region)
        text = "".join([i for i in raw if i.isalpha() or i.isspace()]).lower().strip().rstrip()
        if "buy" in text: return "buy"
        elif "confirming" in text: return "confirming"
        elif "finding" in text: return "finding"
        elif "running" in text: return "running"
        elif "paused" in text: return "paused"
        else: return "unknown"
    


    def update_item_region_avg_color(self) -> None:
        """
        Gets the average red, green, and blue color components
        of `tsm.item_region`,  then updates `self.item_region_avg_color`.

        Execution time:  `80ms`
        """
        tsm.update_item_region()
        delay(0.01)
        r,g,b = [],[],[]
        for x,y in product(range(tsm.item_region.size[0]),range(tsm.item_region.size[1])):
            rgb = tsm.item_region.getpixel((x,y))
            r.append(rgb[0])
            g.append(rgb[1])
            b.append(rgb[2])
        r,g,b = sum(r)/len(r),sum(g)/len(g),sum(b)/len(b)
        self.item_region_avg_color = (r,g,b)
    


    def item_in_item_region(self) -> bool:
        """
        Checks whether or not there is an item in the item area of the TSM window
        (`self.item_region`), based on the average color of all pixels in the region.

        Execution time:  `80ms`
        """
        self.update_item_region_avg_color()
        if self.item_region_avg_color[0] > 1.0 and self.item_region_avg_color[1] > 1.0 and self.item_region_avg_color[2] > 1.0:
            return True
        else: return False
    


    def open_window(self) -> None:
        """
        Opens the main TSM auction window.

        Requires:  `/target Eoch` macro to be on action bar button 2,
        as well as the `Interact with Target` keybind set to `Down Arrow`.
        """
        Keyboard.two()
        delay(0.5)
        Keyboard.down()
        delay(2)
        return















tsm = TSM()
screen = Screen()
cursor = Cursor()





def relog() -> None:
    """
    Logs out to the character selection screen by entering `/camp` in chat
    via `Keyboard.camp()`, then re-enters the world via `Keyboard.enter()`.

    Note:  One 10-20 second delay follows the `/camp` command,
    and another follows the call to `Keyboard.enter()` that re-enters the world.
    """
    Keyboard.camp()
    delay(1*(1+random()))
    Keyboard.enter()
    delay(10*(1+random()))
    Keyboard.enter()
    delay(10*(1+random()))



def start_scan() -> None:
    """
    Starts the buyout sniper by calling
    `cursor.left_click()` at the `tsm.run_buyout_sniper_button` location.

    Note:  The cursor is returned to its initial position after calling,
    which should be the center of the screen (i.e. `screen.center`).
    """
    cursor.left_click(tsm.run_buyout_sniper_button)
    delay(1*(1+random()))
    return



def restart_scan() -> None:
    """
    Restarts the sniper scan by calling
    `cursor.left_click()` at the `tsm.restart_scan_button` location.

    Note:  The cursor is returned to its initial position after calling,
    which should be the center of the screen (i.e. `screen.center`).
    """
    cursor.left_click(tsm.restart_scan_button)
    delay(1*(1+random()))
    return



def buyout_item() -> None:
    """
    Initiates buyout of an item via the TSM buy macro (via `Keyboard.one()`).
    """
    Keyboard.one()
    return



def save_gold_img() -> None:
    """
    Saves an image of player gold shown at the bottom of the backpack to the `gold.png` file.
    """
    Screen.region((1770,965),(1902,982)).save("gold.png")
    return



def save_first_item_image() -> None:
    """
    Takes a screenshot of the item region (`tsm.item_region`) and saves it to the `items.png` file.

    Note:  This function is called only on the first item found after the scan starts, after which `add_to_item_image_list()` is used.
    """
    im1 = Image.open("title.png")
    im2 = Screen.region((523,286),(1420,306))
    width = max(im1.size[0],im2.size[0])
    height = im1.size[1] + im2.size[1]
    im = Image.new('RGB', (width, height))
    im.paste(im1, (0, 0))
    im.paste(im2, (0, im1.size[1]))
    im.save("items.png")
    return



def add_to_item_image_list() -> None:
    """
    Adds the last image of the item region (`tsm.item_region`)
    to the list of item images stored in the `items.png` file.
    """
    im1 = Image.open("items.png")
    im2 = Screen.region((523,286),(1420,306))
    width = max(im1.size[0],im2.size[0])
    height = im1.size[1] + im2.size[1]
    im = Image.new('RGB', (width, height))
    im.paste(im1, (0, 0))
    im.paste(im2, (0, im1.size[1]))
    im.save("items.png")
    return



def cleanup_items_image() -> None:
    """
    Cleans up the `items` image at the end of the program.
    Removes the `ilvl`, `Seller`, and `Bid (Item)` sections.
    """
    im = Image.open("items.png")
    im2 = im.crop((0,0,200,im.height))
    im3 = im.crop((316,0,439,im.height))
    im4 = im.crop((697,0,im.width,im.height))
    im5 = Image.new("RGB", (im2.width+im3.width+im4.width,im.height))
    im5.paste(im2, (0,0))
    im5.paste(im3, (im2.width,0))
    im5.paste(im4, (im2.width+im3.width,0))
    im5.save("items.png")
    return




















if __name__ == "__main__":
    log = Log("log.txt")
    log.file.close()
    init_money = get_users_money()
    init_gold  = init_money[0] + init_money[1]/100 + init_money[2]/10000
    log.write(f"Start time:   {now()}\n")
    log.file.close()
    log.write(f"Start money:  {init_money[0]}g {init_money[1]}s {init_money[2]}c\n\n")
    log.file.close()
    
    
    delay(5)
    running = True
    os.system("cls")
    program_start = dt.now()
    

    i = 0
    num_relogs = 0
    wait_time = int(300*(1+random()))
    start_time = dt.now()
    print('\n')
    print(f"[{now()}]  >> Starting the sniper scan...")
    print(f"[{now()}]  >> Set to run for {wait_time} seconds before relog.")
    print('-'*66 + '\n')
    


    while running:

        if keyboard.is_pressed('q'):                        # Quit the program when 'q' is pressed
            msvcrt.getch()
            running = False

        if seconds_since(start_time) > wait_time:           # Every 300-600 seconds
            num_relogs += 1
            start_time = dt.now()
            wait_time = int(300*(1+random()))
            print('\n')
            print('-'*66)
            print(f"[{now()}]  >> Relogging..."); relog()
            print(f"[{now()}]  >> Opening TSM..."); tsm.open_window()
            print(f"[{now()}]  >> Starting the sniper scan..."); start_scan()
            print(f"[{now()}]  >> Set to run for {wait_time} seconds before relog.")
            print('-'*66 + '\n')
        

        if tsm.scan_running:                                # If the scan is actively searching for items
            if tsm.item_in_item_region():                       # Upon detection of a new item
                i += 1
                if i == 1: save_first_item_image()
                else: add_to_item_image_list()
                print(f"[{now()}]  >> Item found (#{i}).\n"+
                      f"[{now()}]  >> Buying item...", end=" ")
                cursor.move_to(DEFAULT_CURSOR_POSITION)
                cursor.left_click()
                cursor.move_to(screen.center)

                while tsm.get_scan_status() != "buy":           # While the scan is finding the selected item
                    pass
                
                buyout_item()
                delay(0.3*(1+random()))
                print("successfully bought.")

                if tsm.get_scan_status() != "running":          # If the scan has not restarted its search
                    print(f"[{now()}]  >> More than 1 stack found for the previous item.")
                    print(f"[{now()}]  >> Buying out remaining stacks...", end=" ")
                    
                    while tsm.get_scan_status() != "running":   # While the scan has not restarted its search
                        buyout_item()
                        delay(0.3*(1+random()))
                    print("finished buying.")
            
            else: delay(0.01)       # If no new item was found, wait 10ms before checking again
        else: pass                  # If the scan is not actively searching for items, do nothing
    
    


    final_money = get_users_money()
    final_gold  = final_money[0] + final_money[1]/100 + final_money[2]/10000
    delta_gold  = init_gold - final_gold
    delta_money = (int(delta_gold), int((delta_gold-int(delta_gold))*100), int((delta_gold-int(delta_gold)-int((delta_gold-int(delta_gold))*100)/100)*10000))
    
    log.write(f"Final time:   {now()}\n")
    log.file.close()
    # log.write(f"Final money:  {final_money[0]}g {final_money[1]}s {final_money[2]}c\n\n")
    log.write("Final money:  {}g {}s {}c\n\n".format(*final_money))
    log.file.close()
    
    log.write(f"Total time:   {seconds_since(program_start)} seconds\n")
    log.file.close()
    # log.write(f"Money spent:  {delta_money[0]}g {delta_money[1]}s {delta_money[2]}c\n")
    log.write("Money spent:  {}g {}s {}c\n".format(*delta_money))
    log.file.close()
    log.write(f"# of relogs:  {num_relogs}\n")
    log.file.close()
    log.write(f"Items bought: {i}\n\n")
    log.file.close()
    

    log.write(f"Rate of spending: {round(delta_gold/seconds_since(program_start),4)} gold/second\n")
    log.file.close()
    log.write(f"                  {round(delta_gold/seconds_since(program_start)*60,4)} gold/minute\n")
    log.file.close()
    log.write(f"                  {round(delta_gold/seconds_since(program_start)*3600,4)}  gold/hour\n")
    log.file.close()
    

    now_filename = now().replace(':', '.')
    
    cleanup_items_image()
    im = Image.open("items.png")
    im.save(f"logs/{now_filename} ITEMS.png")


    logfile = open("log.txt", "r")
    log_copy = logfile.read()
    logfile.close()

    logfile = open(f"logs/{now_filename} LOG.txt", "w")
    logfile.write(log_copy)
    logfile.close()


    os.remove("log.txt")
    os.remove("items.png")
    os.startfile(os.path.realpath("logs"))  # open the logs folder

