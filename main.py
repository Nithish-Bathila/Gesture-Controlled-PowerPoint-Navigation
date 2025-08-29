import os
import sys
import threading
import time
import cv2
import pystray
from pystray import MenuItem as item
from gesture_detector import GestureDetector
import controller
import tkinter as tk
from tkinter import messagebox
from PIL import Image
import customtkinter as ctk
from typing import Optional

running = True
show_preview_window = False
gesture_detector = GestureDetector()
cap: Optional[cv2.VideoCapture] = None
icon: Optional[pystray.Icon] = None  # Global tray icon

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def resource_path(relative_path: str) -> str:
    try:
        # noinspection PyProtectedMember
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def is_first_run() -> bool:
    appdata = os.getenv('APPDATA')
    if not appdata:
        print("[ERROR] APPDATA environment variable not found.")
        return False
    first_run_file = os.path.join(appdata, 'GestureControl_first_run.flag')
    if not os.path.exists(first_run_file):
        try:
            with open(first_run_file, 'w') as f:
                f.write("First run completed.")
            print(f"[INFO] Created first run flag at {first_run_file}")
        except Exception as e:
            print(f"[WARNING] Failed to write first run flag: {e}")
        return True
    else:
        print(f"[INFO] First run flag exists at {first_run_file}")
    return False


def gesture_loop():
    global running, cap, show_preview_window
    assert cap is not None

    while running:
        if not cap.isOpened():
            time.sleep(0.1)
            continue

        ret, frame = cap.read()
        if not ret:
            time.sleep(0.05)
            continue

        frame = cv2.flip(frame, 1)
        gesture = gesture_detector.detect_gesture(frame)

        if controller.is_powerpoint_open():
            if gesture == 'start_slideshow':
                controller.start_slideshow()
            elif gesture == 'end_slideshow':
                controller.end_slideshow()
            elif gesture == 'next_slide':
                controller.next_slide()
            elif gesture == 'prev_slide':
                controller.prev_slide()

        if show_preview_window:
            try:
                cv2.imshow("Gesture Control Preview", frame)
                if (cv2.getWindowProperty("Gesture Control Preview", cv2.WND_PROP_VISIBLE) < 1 or
                        (cv2.waitKey(1) & 0xFF == 27)):
                    hide_camera_preview()
            except cv2.error:
                pass
        else:
            try:
                cv2.destroyWindow("Gesture Control Preview")
            except cv2.error:
                pass

    if cap:
        cap.release()
    cv2.destroyAllWindows()


def show_camera_preview():
    global show_preview_window
    show_preview_window = True
    update_tray_menu()


def hide_camera_preview():
    global show_preview_window
    show_preview_window = False
    update_tray_menu()


def show_about():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("About Gesture Control",
                        "Gesture-Controlled PowerPoint Navigation\n"
                        "Version 2.9\n\n"
                        "Developed by Nithish\n"
                        "Uses OpenCV, MediaPipe, pywin32, pystray")
    root.destroy()


def quit_app(icon_: pystray.Icon, _: object):
    global running
    running = False
    icon_.stop()


def update_tray_menu():
    global icon
    if icon is None:
        return
    menu = (
        item('Hide Preview', lambda: hide_camera_preview()) if show_preview_window else
        item('Show Preview', lambda: show_camera_preview()),
        item('User Guide', lambda: threading.Thread(target=show_user_guide, daemon=True).start()),
        item('About', lambda: threading.Thread(target=show_about, daemon=True).start()),
        item('Exit', quit_app)
    )
    icon.menu = pystray.Menu(*menu)
    icon.update_menu()


def show_user_guide():
    class UserGuideWindow(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("Welcome to Gesture Control")

            window_width = 560
            window_height = 850
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x = int((screen_width / 2) - (window_width / 2))
            y = int((screen_height / 2) - (window_height / 2))
            self.geometry(f"{window_width}x{window_height}+{x}+{y}")
            self.resizable(False, False)

            self.current_page = 0

            try:
                img1 = Image.open(resource_path("welcome.png")).resize((560, 781), Image.Resampling.LANCZOS)
                img2 = Image.open(resource_path("gestures.png")).resize((560, 781), Image.Resampling.LANCZOS)
                self.images = [
                    ctk.CTkImage(light_image=img1, size=(560, 781)),
                    ctk.CTkImage(light_image=img2, size=(560, 781))
                ]
            except Exception as e:
                tk.messagebox.showerror("Error", f"Could not load guide images: {e}")
                self.destroy()
                return

            self.image_label = ctk.CTkLabel(self, image=self.images[0], text="")
            self.image_label.pack(pady=(20, 10))

            self.button_frame = ctk.CTkFrame(self)
            self.button_frame.pack(pady=(0, 10))

            self.prev_btn = ctk.CTkButton(self.button_frame, text="Previous", command=self.prev_page, width=100)
            self.prev_btn.grid(row=0, column=0, padx=10)

            self.next_btn = ctk.CTkButton(self.button_frame, text="Next", command=self.next_page, width=100)
            self.next_btn.grid(row=0, column=1, padx=10)

            self.exit_btn = ctk.CTkButton(self.button_frame, text="Exit", command=self.destroy, width=100)
            self.exit_btn.grid(row=0, column=2, padx=10)

            self.bind("<Right>", lambda event: self.next_page())
            self.bind("<Left>", lambda event: self.prev_page())
            self.focus_set()

            self.update_buttons()

        def next_page(self):
            if self.current_page < len(self.images) - 1:
                self.current_page += 1
                self.image_label.configure(image=self.images[self.current_page])
            self.update_buttons()

        def prev_page(self):
            if self.current_page > 0:
                self.current_page -= 1
                self.image_label.configure(image=self.images[self.current_page])
            self.update_buttons()

        def update_buttons(self):
            self.prev_btn.configure(state="normal" if self.current_page > 0 else "disabled")
            self.next_btn.configure(state="normal" if self.current_page < len(self.images) - 1 else "disabled")

    app = UserGuideWindow()
    app.mainloop()


def setup_tray():
    global icon
    try:
        icon_image = Image.open(resource_path('icon.png'))
    # Catching general Exception to ensure unexpected crashes are logged
    except Exception:
        print("Could not load tray icon.")
        return

    icon = pystray.Icon("GestureControl", icon_image, "Gesture Control")
    update_tray_menu()
    icon.run()


def main():
    try:
        if getattr(sys, 'frozen', False):
            os.chdir(os.path.dirname(sys.executable))
        else:
            os.chdir(os.path.dirname(os.path.abspath(__file__)))

        global cap
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("Error: Cannot access webcam.")
            return

        gesture_thread = threading.Thread(target=gesture_loop, daemon=True)
        gesture_thread.start()

        if is_first_run():
            try:
                show_user_guide()
            except Exception as e:
                print(f"[ERROR] Failed to show user guide: {e}")

        setup_tray()
    except Exception as e:
        print(f"[FATAL] Unexpected error occurred: {e}")


if __name__ == "__main__":
    main()
