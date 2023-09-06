import pywinauto
import pywintypes
import pyproj
import win32gui
import time
import psutil
import sys
import re
import xml.etree.ElementTree as ET
import tkinter
import tkintermapview

element_re = re.compile(r"Element.*?(\d+)")

root = ET.parse(sys.argv[1])
utm = root.find("./Strecke/UTM")
utm_we = int(utm.attrib.get("UTM_WE", 0)) * 1000
utm_ns = int(utm.attrib.get("UTM_NS", 0)) * 1000
utm_zone = int(utm.attrib.get("UTM_Zone", 0))
elemente = dict(
    (el.attrib.get("Nr"), el) for el in root.findall("./Strecke/StrElement")
)
print(len(elemente))

proj = pyproj.Proj(proj="utm", zone=utm_zone, ellps="WGS84", preserve_units=False)

# create tkinter window
root_tk = tkinter.Tk()
root_tk.attributes("-topmost", True)
root_tk.geometry(f"{800}x{600}")
root_tk.title("Map data © OpenStreetMap contributors")
root_tk.update()

# create map widget
map_widget = tkintermapview.TkinterMapView(
    root_tk, width=800, height=600, corner_radius=0
)
map_widget.pack(fill="both", expand=True)
map_widget.canvas.unbind("<MouseWheel>")

use_orm = True
if use_orm:
    map_widget.set_overlay_tile_server(
        "http://a.tiles.openrailwaymap.org/standard/{z}/{x}/{y}.png"
    )  # railway infrastructure
    root_tk.title(
        "Map data © OpenStreetMap contributors, Style: CC-BY-SA 2.0 OpenRailwayMap"
    )

marker = map_widget.set_marker(52.516268, 13.377695)
active_hwnd = None
treeview_hwnd = None


def update_pos():
    global active_hwnd, treeview_hwnd
    root_tk.after(500, update_pos)

    hwnd = win32gui.GetForegroundWindow()
    print(hwnd)
    if (hwnd != active_hwnd) or (treeview_hwnd is None):
        active_hwnd = hwnd
        treeview_hwnd = None

        try:
            classname = win32gui.GetClassName(hwnd)
        except pywintypes.error:
            return

        if classname != "TFormZusi3DEditor":
            return

        app = pywinauto.Application().connect(handle=hwnd)
        win = app.window(handle=hwnd)

        child = win.window(title="Klicken + Anzeige")
        treeview_hwnd = child.window(class_name="TTreeView")

    tv = pywinauto.controls.common_controls.TreeViewWrapper(treeview_hwnd)
    for item in tv.print_items().splitlines():
        match = element_re.match(item)
        if not match:
            continue

        element = elemente[match.group(1)]
        g = element.find("g")
        b = element.find("b")

        x = utm_we + (float(g.get("X", 0)) + float(b.get("X", 0))) / 2.0
        y = utm_ns + (float(g.get("Y", 0)) + float(b.get("Y", 0))) / 2.0

        print(utm_zone, x, y)
        lon, lat = proj(x, y, inverse=True)
        print(lon, lat)

        marker.set_position(lat, lon)
        canvas_pos_x, canvas_pos_y = marker.get_canvas_pos(marker.position)
        if (
            (canvas_pos_x < 20)
            or (canvas_pos_y < 20)
            or (canvas_pos_x > map_widget.width - 20)
            or (canvas_pos_y > map_widget.height - 20)
        ):
            map_widget.set_position(lat, lon)
        break


root_tk.after(1000, update_pos)
tkinter.mainloop()
