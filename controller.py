import win32com.client

def is_powerpoint_open():
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        return ppt
    except Exception:
        return None

def start_slideshow():
    ppt = is_powerpoint_open()
    if ppt:
        try:
            presentation = ppt.ActivePresentation
            presentation.SlideShowSettings.Run()
            print("[INFO] Slideshow started.")
        except Exception as e:
            print("[ERROR] Could not start slideshow:", e)

def end_slideshow():
    ppt = is_powerpoint_open()
    if ppt:
        try:
            slideshow = ppt.SlideShowWindows(1)
            slideshow.View.Exit()
            print("[INFO] Slideshow ended.")
        except Exception as e:
            print("[ERROR] Could not end slideshow:", e)

def next_slide():
    ppt = is_powerpoint_open()
    if ppt:
        try:
            slideshow = ppt.SlideShowWindows(1)
            slideshow.View.Next()
            print("[INFO] Next slide.")
        except Exception as e:
            print("[ERROR] Could not go to next slide:", e)

def prev_slide():
    ppt = is_powerpoint_open()
    if ppt:
        try:
            slideshow = ppt.SlideShowWindows(1)
            slideshow.View.Previous()
            print("[INFO] Previous slide.")
        except Exception as e:
            print("[ERROR] Could not go to previous slide:", e)
