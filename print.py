import win32gui
import win32ui
import docx
import io
from PIL import Image
import win32con
from pywinauto import Application
import time
import tkinter as tk
from tkinter import ttk

# Define the prefix of the window names
window_name_prefix = "Single Beam Max"

# Define the exact name of the window to include
window_name_exact = "Profile Window"


def show_loading_screen():
    loading_screen = tk.Tk()
    loading_screen.title("Loading...")
    loading_screen.geometry("300x100")
    screen_width = loading_screen.winfo_screenwidth()
    screen_height = loading_screen.winfo_screenheight()
    x = int((screen_width / 2) - (300 / 2))
    y = int((screen_height / 2) - (100 / 2))
    loading_screen.geometry("300x100+{}+{}".format(x, y))

    label = ttk.Label(loading_screen, text="Aguarde...")
    label.pack(pady=20)

    progress_bar = ttk.Progressbar(loading_screen, mode="indeterminate")
    progress_bar.pack(padx=30, pady=10)
    progress_bar.start()

    loading_screen.update()
    return loading_screen


loading_screen = show_loading_screen()

time.sleep(5) 
app = Application(backend='win32').connect(title_re='^' + window_name_prefix)

process_id = app.process

main_window = app.window(title_re='^' + window_name_prefix)
# Get the dropdown menu
dropdown = main_window.child_window(class_name="TComboBox")

# Get the full title of the main window
full_title = main_window.window_text()

# Remove the prefix from the full title
title_without_prefix = full_title.replace(window_name_prefix, "").strip()

# Get the number of items in the dropdown menu
item_count = dropdown.item_count()

# Take screenshots of the windows
screenshots = []
def enum_windows_callback(hwnd, dropdown_text, screenshots):
    text = win32gui.GetWindowText(hwnd)
    if text.startswith(window_name_prefix) or text == window_name_exact:
        # Get the dimensions of the window
        rect = win32gui.GetWindowRect(hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]

        # Resize the "Single Beam Max" window to match the width of the "Profile Window"
        if text.startswith(window_name_prefix):
            profile_window_width = None
            for _, img in screenshots:
                if img.width > width:
                    profile_window_width = img.width
                    break

            if profile_window_width is not None:
                win32gui.SetWindowPos(hwnd, 0, rect[0], rect[1], profile_window_width, height, win32con.SWP_NOZORDER)
                # Get the new dimensions of the window after resizing
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]

        # Get the device context of the window
        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        # Create a compatible memory DC and bitmap
        save_dc = mfc_dc.CreateCompatibleDC()
        save_bitmap = win32ui.CreateBitmap()
        save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(save_bitmap)
        # Copy the contents of the window into the bitmap
        time.sleep(0.2)
        win32gui.BitBlt(save_dc.GetSafeHdc(), 0, 0, width, height, hwnd_dc, 0, 0, win32con.SRCCOPY)
        # Convert the bitmap to a PIL.Image.Image object
        bmp_info = save_bitmap.GetInfo()
        bmp_str = save_bitmap.GetBitmapBits(True)
        screenshot = Image.frombuffer('RGB', (bmp_info['bmWidth'], bmp_info['bmHeight']), bmp_str, 'raw', 'BGRX', 0, 1)
        screenshots.append((dropdown_text, screenshot))

        # Clean up
        save_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        win32gui.DeleteObject(save_bitmap.GetHandle())
def enum_windows_callback_wrapper(hwnd, args):
    dropdown_text, screenshots = args
    enum_windows_callback(hwnd, dropdown_text, screenshots)


# Create a new Word document
document = docx.Document()



'''
# Find the button that opens the file browser in the "Single Beam Max" window
# Send the 'F2' key to the window
window_name_prefix.type_keys('{F2}')


# Connect to the file browser window (use the specific title of your file browser window)
file_browser = Application(backend='win32').connect(title='Abrir')

# Locate the list of folders in the file browser window
folders_list = file_browser.window(title='Abrir').child_window(control_type="List")

# Get the number of folders in the list
folder_count = folders_list.item_count()
'''



# Iterate through the items in the dropdown menu
def take_screenshot(window_title_prefix, crop_left, crop_top, crop_right, crop_bottom):
    def _window_enum_callback(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if title.startswith(window_title_prefix):
            window_handles.append(hwnd)

    window_handles = []
    win32gui.EnumWindows(_window_enum_callback, None)

    for hwnd in window_handles:
        rect = win32gui.GetWindowRect(hwnd)
        width = rect[2] - rect[0] - crop_left - crop_right
        height = rect[3] - rect[1] - crop_top - crop_bottom

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        save_bitmap = win32ui.CreateBitmap()
        save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(save_bitmap)

        time.sleep(0.2)
        win32gui.BitBlt(save_dc.GetSafeHdc(), 0, 0, width, height, hwnd_dc, crop_left, crop_top, win32con.SRCCOPY)

        bmp_info = save_bitmap.GetInfo()
        bmp_str = save_bitmap.GetBitmapBits(True)
        screenshot = Image.frombuffer('RGB', (bmp_info['bmWidth'], bmp_info['bmHeight']), bmp_str, 'raw', 'BGRX', 0, 1)

        save_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        win32gui.DeleteObject(save_bitmap.GetHandle())

        return screenshot

# Define cropping values
crop_left = 10
crop_top = 10
crop_right = 25
crop_bottom = 10

# Iterate through the items in the dropdown menu
for i in range(item_count):
    # Select the item in the dropdown menu
    dropdown.select(i)
    # Get the text of the selected dropdown item
    dropdown_text = dropdown.selected_text()

    # Take the Single Beam Max and Profile Window screenshots with the cropping values
    single_beam_screenshot = take_screenshot(window_name_prefix, crop_left, crop_top, crop_right, crop_bottom)
    profile_window_screenshot = take_screenshot(window_name_exact, crop_left, crop_top, crop_right, crop_bottom)

    # Combine the two screenshots into one image
    combined_width = max(single_beam_screenshot.width, profile_window_screenshot.width)
    combined_height = single_beam_screenshot.height + profile_window_screenshot.height

    combined_image = Image.new('RGB', (combined_width, combined_height))
    combined_image.paste(single_beam_screenshot, (0, 0))
    combined_image.paste(profile_window_screenshot, (0, single_beam_screenshot.height))

    # Save the combined image to the Word document
    image_bytes = io.BytesIO()
    combined_image.save(image_bytes, format='png')

    img = document.add_picture(image_bytes, width=docx.shared.Inches(6))

    img_paragraph = img._inline.getparent()
    img_paragraph.attrib.update({
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spaceBefore': '0',
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spaceAfter': '0',
    })

loading_screen.destroy()

# Save the Word document with the title without prefix
document.save(f'{title_without_prefix}.docx')
