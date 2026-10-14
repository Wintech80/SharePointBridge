"""System tray icon for the packaged exe."""
import os
import webbrowser

import pystray
from PIL import Image, ImageDraw

import winutil


def make_image():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((4, 4, 60, 60), radius=12, fill=(3, 120, 124, 255))  # SharePoint-ish teal
    d.polygon([(16, 24), (40, 24), (40, 18), (52, 28), (40, 38), (40, 32), (16, 32)], fill="white")  # right arrow
    d.polygon([(48, 40), (24, 40), (24, 34), (12, 44), (24, 54), (24, 48), (48, 48)], fill="white")  # left arrow
    return img


def run_tray(url, data_dir, log_path, password, on_quit):
    """Show the tray icon and block until Quit is chosen."""

    def open_log(icon, item):
        os.startfile(log_path)

    def open_folder(icon, item):
        os.startfile(data_dir)

    def copy_password(icon, item):
        winutil.copy_to_clipboard(password)
        icon.notify("Password copied to the clipboard.", "SharePoint Bridge")

    def toggle_autostart(icon, item):
        winutil.set_autostart(not winutil.autostart_enabled())

    def quit_app(icon, item):
        on_quit()
        icon.stop()

    items = [
        pystray.MenuItem(f"Running at {url}", None, enabled=False),
        pystray.MenuItem("Open in browser", lambda i, it: webbrowser.open(url)),
        pystray.MenuItem("View log", open_log),
        pystray.MenuItem("Open data folder", open_folder),
    ]
    if password:
        items.append(pystray.MenuItem("Copy password", copy_password))
    items += [
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Start with Windows", toggle_autostart, checked=lambda item: winutil.autostart_enabled()),
        pystray.MenuItem("Quit", quit_app),
    ]
    icon = pystray.Icon("SharePointBridge", make_image(), "SharePoint Bridge", pystray.Menu(*items))
    icon.run(setup=lambda i: (setattr(i, "visible", True), i.notify(f"Running at {url}", "SharePoint Bridge")))
