# Dependencies: pyautogui, requests, beautifulsoup4, pywin32

from datetime import datetime
from pathlib import Path
import time

import pyautogui
import requests
import win32com.client
from bs4 import BeautifulSoup
import win32con
import win32gui


URL = "https://www.goodreturns.in/gold-rates/chennai.html"
OUTPUT_DIR = Path(__file__).resolve().parent
timestamp = datetime.now()
date_stamp = timestamp.strftime("%Y-%m-%d")

page_screenshot_path = OUTPUT_DIR / "gold_Rate.png"
workbook_path = OUTPUT_DIR / f"daily_report_{date_stamp}.xlsx"
final_screenshot_path = OUTPUT_DIR / f"daily_report_{date_stamp}.png"


def extract_rate_information(html: str) -> str:
    """Extract fetched data from the page's gr_top_intro_content element."""
    soup = BeautifulSoup(html, "html.parser")
    content = soup.find(id="gr_top_intro_content")

    if content is None:
        return "The page element 'gr_top_intro_content' was not found."

    text = " ".join(content.get_text(" ", strip=True).split())
    return text or "The page element 'gr_top_intro_content' is empty."


# Open Chrome, navigate to the rates page, and capture it.
pyautogui.press("win")
pyautogui.write("chrome", interval=0.1)
pyautogui.press("enter")
time.sleep(3)

pyautogui.hotkey("ctrl", "l")
pyautogui.write(URL, interval=0.02)
pyautogui.press("enter")
time.sleep(6)
pyautogui.screenshot().save(page_screenshot_path)

# Fetch the rates page.
try:
    response = requests.get(
        URL,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=20,
    )
    response.raise_for_status()
    rate_information = extract_rate_information(response.text)
except requests.RequestException as error:
    rate_information = f"Unable to retrieve rate information: {error}"

# Open a new blank Excel workbook and fill in the report.
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

workbook = excel.Workbooks.Add()
sheet = workbook.Worksheets(1)
sheet.Name = "Daily Report"

sheet.Range("A1:D1").Value = (
    "Date & Time",
    "Fetched Data",
    "Information",
    "Screenshot",
)
sheet.Range("A2:C2").Value = (
    timestamp.strftime("%Y-%m-%d %H:%M:%S"),
    rate_information,
    "gold_Rate",
)

# Format the header and report row.
header = sheet.Range("A1:D1")
header.Font.Bold = True
header.Font.Color = 0xFFFFFF
header.Interior.Color = 0x784E1F

report_row = sheet.Range("A2:D2")
report_row.WrapText = True
report_row.VerticalAlignment = -4160  # xlTop

sheet.Columns("A").ColumnWidth = 22
sheet.Columns("B").ColumnWidth = 70
sheet.Columns("C").ColumnWidth = 24
sheet.Columns("D").ColumnWidth = 60
sheet.Rows(2).RowHeight = 300

# Embed the webpage screenshot in cell D2.
image = pyautogui.screenshot()  # keep Excel visible before final capture
from PIL import Image as PILImage

with PILImage.open(page_screenshot_path) as page_image:
    image_width, image_height = page_image.size

max_width = 450
max_height = 300
scale = min(max_width / image_width, max_height / image_height, 1)
picture_width = image_width * scale
picture_height = image_height * scale

cell = sheet.Range("D2")
sheet.Shapes.AddPicture(
    str(page_screenshot_path),
    0,  # msoFalse: do not link to the file
    -1,  # msoTrue: save the picture in the workbook
    cell.Left,
    cell.Top,
    picture_width,
    picture_height,
)

# Save the workbook, bring Excel to the foreground, and capture the sheet.
workbook.SaveAs(str(workbook_path), FileFormat=51)  # 51 = .xlsx
excel.Visible = True
excel.WindowState = -4137  # xlMaximized
workbook.Activate()
sheet.Activate()
sheet.Range("A1").Select()
excel.ActiveWindow.SplitRow = 1
excel.ActiveWindow.FreezePanes = True

excel_hwnd = excel.Hwnd
win32gui.ShowWindow(excel_hwnd, win32con.SW_MAXIMIZE)
win32gui.SetForegroundWindow(excel_hwnd)

time.sleep(3)
pyautogui.screenshot().save(final_screenshot_path)