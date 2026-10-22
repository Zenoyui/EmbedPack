#!/usr/bin/env python3
"""Screenshot via xdg-desktop-portal (Wayland). Usage: python3 portal_shot.py out.png"""
import sys, shutil, random
from gi.repository import Gio, GLib

out = sys.argv[1]
bus = Gio.bus_get_sync(Gio.BusType.SESSION, None)
token = "lab%d" % random.randint(0, 10**8)
sender = bus.get_unique_name()[1:].replace(".", "_")
path = "/org/freedesktop/portal/desktop/request/%s/%s" % (sender, token)
loop = GLib.MainLoop()
res = {}

def on_response(conn, snd, p, iface, sig, params, *a):
    code, data = params.unpack()
    res["code"], res["uri"] = code, data.get("uri")
    loop.quit()

bus.signal_subscribe("org.freedesktop.portal.Desktop", "org.freedesktop.portal.Request",
                     "Response", path, None, Gio.DBusSignalFlags.NONE, on_response)
bus.call_sync("org.freedesktop.portal.Desktop", "/org/freedesktop/portal/desktop",
              "org.freedesktop.portal.Screenshot", "Screenshot",
              GLib.Variant("(sa{sv})", ("", {"handle_token": GLib.Variant("s", token),
                                             "interactive": GLib.Variant("b", False)})),
              None, Gio.DBusCallFlags.NONE, -1, None)
GLib.timeout_add_seconds(30, loop.quit)
loop.run()
if res.get("code") != 0 or not res.get("uri"):
    sys.exit(1)
shutil.copy(GLib.filename_from_uri(res["uri"])[0], out)
