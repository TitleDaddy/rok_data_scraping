import win32gui
import win32api
import win32con
import win32ui
import time
import ctypes
from screen_coords import templates as temp
import cv2
import numpy as np
from pathlib import Path
import os
from PIL import Image
import pyautogui
import pytesseract
import nltk
import datetime as dt
import openpyxl
import csv
import clipboard

BASE_DIR = Path(__file__).resolve().parent

nltk.download("punkt")

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"  # Location of tesseract.eve file
)

def main():
    # while True:
    #     click(50,50)
    #     time.sleep(1)
    wait = 1
    number = 1
    csv_file_name = f"kingdom_analytics_{dt.date.today()}.csv"
    csv_file = os.path.join(BASE_DIR, csv_file_name)
    with open(
        csv_file, "w+", encoding="UTF8", newline=""
    ) as csv_master:
        writer = csv.writer(csv_master)
        writer.writerow([
            "ID",
            "NAME",
            "ALLIANCE",
            "POWER",
            "KILLS",
            "DEATHS",
            "DATE"
            ]
        )

    page = "ranking"
    take_screenshot(page)
    ranking = image_to_text("my_name", page, False)

    if ranking == "1000+":
        check_ranking = False
    else:
        check_ranking = True

    for n in range(900):

        if check_ranking:
            on_player = False
            if str(n) == ranking:
                on_player = True

        if n <= 3:
            print(f"\n\n\nPROFILE {number}")
            y = 300 + 120 * n

            # open profile
            click(960, y)
            time.sleep(wait)
            page = "profile"
            take_screenshot(page)

            # open more infos
            click(468, 796)           
            time.sleep(wait)

            #copy name
            click(455, 189)
            time.sleep(wait)
            name_paste = clipboard.paste()
            page = "more info"
            take_screenshot(page)           

            # close more info
            click(1677, 68)
            time.sleep(wait)

            # close profile
            click(1637, 128)
            time.sleep(wait)
            print ("FIRST BIT")
        if n > 3:
            print(f"\n\n\nPROFILE {number}")

            # open profile
            click(960, 720)
            time.sleep(wait)
            page = "profile"
            take_screenshot(page)

            # open more info
            click(468, 796)
            time.sleep(wait)

            #copy name
            click(455, 189)
            time.sleep(wait)
            name_paste = clipboard.paste()
            page = "more info"
            take_screenshot(page)  

            # close more info
            click(1677, 68)
            time.sleep(wait)

            # close profile
            click(1637, 128)
            time.sleep(wait)
            print ("SECOND BIT")

        page = "profile"
        player_info = image_to_text(name_paste, page, on_player)
        with open(
            csv_file, "a+", encoding="UTF8", newline=""
        ) as invoice_master:
            writer = csv.writer(invoice_master)
            writer.writerow(
                [
                    player_info["id"],
                    player_info["name"],
                    player_info["alliance"],
                    player_info["power"],
                    player_info["kills"],
                    player_info["deaths"],
                    dt.date.today()
                ]
            )
        number += 1
    csv_master.close()


def click(x, y):
    hwnd = win32gui.FindWindow(None, "Bluestacks")
    width = int(get_window_size(hwnd)["width"])
    height = int(get_window_size(hwnd)["height"])
    if width / height != 1920 / 1080:
        height = height - 40
        print(f"height: {height}")
        ratio = height / 1080
    else:
        ratio = 1
    print(f"x: {x}, y: {y}")
    x = x * ratio
    x = int(x)
    y = y * ratio
    y = int(y)
    print(f"x: {x}, y: {y}")
    lParam = win32api.MAKELONG(x, y)

    hwnd1 = win32gui.FindWindowEx(hwnd, None, None, None)
    win32gui.SendMessage(hwnd1, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
    win32gui.SendMessage(hwnd1, win32con.WM_LBUTTONUP, None, lParam)
    return


def take_screenshot(page):
    print("taking screenshot...")
    jpgfilenamename = os.path.join(BASE_DIR, "static", "screenshot.jpg")
    idfilename = os.path.join(BASE_DIR, "static", "id.jpg")
    namefilename = os.path.join(BASE_DIR, "static", "name.jpg")
    alliancefilename = os.path.join(BASE_DIR, "static", "alliance.jpg")
    powerfilename = os.path.join(BASE_DIR, "static", "power.jpg")
    killsfilename = os.path.join(BASE_DIR, "static", "kills.jpg")
    deathsfilename = os.path.join(BASE_DIR, "static", "deaths.jpg")
    rankingfilename = os.path.join(BASE_DIR, "static", "ranking.jpg")

    # width = 1920
    # height = 1080
    hwnd = win32gui.FindWindow(None, "Bluestacks")
    width = int(get_window_size(hwnd)["width"])
    height = int(get_window_size(hwnd)["height"])
    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    dataBitMap = win32ui.CreateBitmap()
    dataBitMap.CreateCompatibleBitmap(dcObj, width, height)
    cDC.SelectObject(dataBitMap)
    cDC.BitBlt((0, 0), (width, height), dcObj, (0, 0), win32con.SRCCOPY)
    dataBitMap.SaveBitmapFile(cDC, jpgfilenamename)
    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(dataBitMap.GetHandle())
    img = Image.open(jpgfilenamename)

    if width / height != 1920 / 1080:
        height_crop = height - 40
        width_crop = height_crop * 1920 / 1080

        screen_width, screen_height = pyautogui.size()
        print(f"SCREEN WIDTH: {screen_width}")
        if width == screen_width:
            print("cropping sidebars...")
            side_border = (screen_width - 40) / 2 - width_crop / 2
            img_cropped = img.crop(
                (side_border, 40, side_border + width_crop, height_crop + 40)
            )
        else:
            img_cropped = img.crop((1, 40, width_crop, height_crop + 40 - 1))
        img_resize = img_cropped.resize((1920,1080), Image.ANTIALIAS)
        img_resize.save(jpgfilenamename)

    if page == "profile":
        img_id = img_resize.crop(
            (
                temp["id"]["x0"],
                temp["id"]["y0"],
                temp["id"]["x1"],
                temp["id"]["y1"],
            )
        )
        img_name = img_resize.crop(
            (
                temp["name"]["x0"],
                temp["name"]["y0"],
                temp["name"]["x1"],
                temp["name"]["y1"],
            )
        )
        img_alliance = img_resize.crop(
            (
                temp["alliance"]["x0"],
                temp["alliance"]["y0"],
                temp["alliance"]["x1"],
                temp["alliance"]["y1"],
            )
        )
        img_power = img_resize.crop(
            (
                temp["power"]["x0"],
                temp["power"]["y0"],
                temp["power"]["x1"],
                temp["power"]["y1"],
            )
        )
        img_kills = img_resize.crop(
            (
                temp["kills"]["x0"],
                temp["kills"]["y0"],
                temp["kills"]["x1"],
                temp["kills"]["y1"],
            )
        )
        img_id.save(idfilename)
        img_name.save(namefilename)
        img_alliance.save(alliancefilename)
        img_power.save(powerfilename)
        img_kills.save(killsfilename)
    elif page == "more info":
        img_deaths = img_resize.crop(
            (
                temp["deaths"]["x0"],
                temp["deaths"]["y0"],
                temp["deaths"]["x1"],
                temp["deaths"]["y1"],
            )
        )
        img_deaths.save(deathsfilename)
    elif page == "ranking":
        img_ranking = img_resize.crop(
            (
                temp["ranking"]["x0"],
                temp["ranking"]["y0"],
                temp["ranking"]["x1"],
                temp["ranking"]["y1"],
            )
        )
        img_ranking.save(rankingfilename)
    print("finished screenshot")
    return


def get_window_size(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    x0 = rect[0]
    y0 = rect[1]
    x1 = rect[2]
    y1 = rect[3]
    w = x1 - x0
    h = y1 - y0
    # print (f"WINDOW DIMENSITON\nw: {w}, h: {h}")

    context = {
        "width": w,
        "height": h,
    }

    return context

def read_text(image):
    print("converting")
    custom_config = r"--oem 3 kor+chi_sim+eng+jpn+vie --psm 6"
    text = pytesseract.image_to_string(image, lang="eng+kor+vie+jap+sun_chi")
    return text

def image_to_text(name_paste, page, on_player):
    idfilename = os.path.join(BASE_DIR, "static", "id.jpg")
    namefilename = os.path.join(BASE_DIR, "static", "name.jpg")
    alliancefilename = os.path.join(BASE_DIR, "static", "alliance.jpg")
    powerfilename = os.path.join(BASE_DIR, "static", "power.jpg")
    killsfilename = os.path.join(BASE_DIR, "static", "kills.jpg")
    deathsfilename = os.path.join(BASE_DIR, "static", "deaths.jpg")
    rankingfilename = os.path.join(BASE_DIR, "static", "ranking.jpg")

    file_list = [
        idfilename,
        namefilename,
        alliancefilename,
        powerfilename,
        killsfilename,
        deathsfilename,
    ]
    round = 1
    context = {}
    if page == "profile":
        for file in file_list:
            image = cv2.imread(file)
            gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            if round == 1:
                ret, thresh_image = cv2.threshold(
                    gray_image, 127, 255, cv2.THRESH_BINARY_INV
                )
            elif round == 2:
                ret, thresh_image = cv2.threshold(
                    gray_image, 180, 255, cv2.THRESH_BINARY_INV
                )
            elif round ==3 or round == 6:
                ret, thresh_image = cv2.threshold(
                    gray_image, 180, 255, cv2.THRESH_BINARY_INV
                )
            else:
                ret, thresh_image = cv2.threshold(
                    gray_image, 220, 255, cv2.THRESH_BINARY_INV
                )
            thresh_image = cv2.GaussianBlur(thresh_image, (5, 5), 0)
            # cv2.imshow('image window', thresh_image)
            # cv2.waitKey(0)
            # cv2.destroyAllWindows()
            text = read_text(thresh_image)
            if round == 1:
                text = text.split()
                context["id"] = text
                print(f"ID: {text}")
            if round == 2:
                if on_player:
                    context["name"] = text
                else:
                    context["name"] = name_paste
                print(f"NAME: {name_paste}")
            if round == 3:
                context["alliance"] = text
                print(f"ALLIANCE: {text}")
            if round == 4:
                context["power"] = text
                print(f"POWER: {text}")
            if round == 5:
                context["kills"] = text
                print(f"KILLS: {text}")
            if round == 6:
                context["deaths"] = text
                print(f"DEATHS: {text}")
            round +=1

    if page == "ranking":
        image = cv2.imread(rankingfilename)
        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        ret, thresh_image = cv2.threshold(
            gray_image, 240, 255, cv2.THRESH_BINARY_INV
        )
        thresh_image = cv2.GaussianBlur(thresh_image, (5, 5), 0)
        text = read_text(thresh_image)
        context["ranking"] = text
        print(f"RANKING: {text}")

    return context

main()
