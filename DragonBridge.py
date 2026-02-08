# -*- coding: utf-8 -*-
"""
Dragon Bridge for LibreOffice
Bridges Dragon NaturallySpeaking with LibreOffice Writer via clipboard
monitoring and voice command processing.

Copyright (c) 2025 FJCCV - Fellowship of Jesus Christ in California Valley
License: MIT
"""

import uno
import unohelper
import threading
import time
import ctypes
import ctypes.wintypes
import traceback
import os
import json

from com.sun.star.task import XJobExecutor
from com.sun.star.lang import XServiceInfo

# Windows Clipboard API via ctypes

CF_UNICODETEXT = 13

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

OpenClipboard = user32.OpenClipboard
OpenClipboard.argtypes = [ctypes.wintypes.HWND]
OpenClipboard.restype = ctypes.wintypes.BOOL

CloseClipboard = user32.CloseClipboard
CloseClipboard.argtypes = []
CloseClipboard.restype = ctypes.wintypes.BOOL

GetClipboardData = user32.GetClipboardData
GetClipboardData.argtypes = [ctypes.wintypes.UINT]
GetClipboardData.restype = ctypes.wintypes.HANDLE

GetClipboardSequenceNumber = user32.GetClipboardSequenceNumber
GetClipboardSequenceNumber.argtypes = []
GetClipboardSequenceNumber.restype = ctypes.wintypes.DWORD

GlobalLock = kernel32.GlobalLock
GlobalLock.argtypes = [ctypes.wintypes.HANDLE]
GlobalLock.restype = ctypes.c_void_p

GlobalUnlock = kernel32.GlobalUnlock
GlobalUnlock.argtypes = [ctypes.wintypes.HANDLE]
GlobalUnlock.restype = ctypes.wintypes.BOOL


def get_clipboard_text():
    """Read current clipboard text using Windows API."""
    text = ""
    if OpenClipboard(0):
        try:
            handle = GetClipboardData(CF_UNICODETEXT)
            if handle:
                ptr = GlobalLock(handle)
                if ptr:
                    try:
                        text = ctypes.wstring_at(ptr)
                    finally:
                        GlobalUnlock(handle)
        finally:
            CloseClipboard()
    return text


def get_clipboard_seq():
    """Get clipboard sequence number to detect changes."""
    return GetClipboardSequenceNumber()


# Configuration

def get_config_path():
    """Get path for storing settings."""
    profile = os.environ.get("APPDATA", os.path.expanduser("~"))
    config_dir = os.path.join(profile, "DragonBridge")
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "settings.json")


def load_config():
    """Load settings from disk."""
    defaults = {
        "poll_interval_ms": 300,
        "auto_space": True,
        "show_notifications": True,
    }
    try:
        path = get_config_path()
        if os.path.exists(path):
            with open(path, "r") as f:
                saved = json.load(f)
            defaults.update(saved)
    except Exception:
        pass
    return defaults


def save_config(config):
    """Save settings to disk."""
    try:
        with open(get_config_path(), "w") as f:
            json.dump(config, f, indent=2)
    except Exception:
        pass


# Voice Command Processing

VOICE_COMMANDS = {
    "new line": "\n",
    "new paragraph": "\n\n",
    "tab key": "\t",
    "period": ".",
    "comma": ",",
    "exclamation point": "!",
    "exclamation mark": "!",
    "question mark": "?",
    "colon": ":",
    "semicolon": ";",
    "open paren": "(",
    "close paren": ")",
    "open quote": "\u201c",
    "close quote": "\u201d",
    "open single quote": "\u2018",
    "close single quote": "\u2019",
    "hyphen": "-",
    "dash": "\u2014",
    "ellipsis": "\u2026",
}

ACTION_COMMANDS = {
    "select all": ".uno:SelectAll",
    "undo that": ".uno:Undo",
    "redo that": ".uno:Redo",
    "bold that": ".uno:Bold",
    "italicize that": ".uno:Italic",
    "underline that": ".uno:Underline",
    "cut that": ".uno:Cut",
    "copy that": ".uno:Copy",
    "paste that": ".uno:Paste",
    "delete that": ".uno:SwBackspace",
    "save document": ".uno:Save",
    "scratch that": ".uno:Undo",
}


def process_voice_commands(text):
    """
    Check if text matches a known Dragon voice command.
    Returns (type, result) where type is 'action', 'text', or 'insert'.
    """
    lower = text.strip().lower()

    if lower in ACTION_COMMANDS:
        return ("action", ACTION_COMMANDS[lower])

    if lower in VOICE_COMMANDS:
        return ("text", VOICE_COMMANDS[lower])

    return ("insert", text)


# Clipboard Bridge

class ClipboardBridge:
    """
    Monitors the Windows clipboard for changes and inserts new text
    into the active LibreOffice Writer document at the cursor position.
    """

    def __init__(self, ctx):
        self.ctx = ctx
        self.running = False
        self.thread = None
        self.config = load_config()
        self.last_seq = get_clipboard_seq()
        self.last_text = get_clipboard_text()

    def start(self):
        if self.running:
            return
        self.running = True
        self.last_seq = get_clipboard_seq()
        self.last_text = get_clipboard_text()
        self.thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.thread.start()
        self._notify("Clipboard monitoring started - dictate into DragonPad and transfer.")

    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=2)
            self.thread = None
        self._notify("Clipboard monitoring stopped.")

    def is_running(self):
        return self.running

    def _monitor_loop(self):
        """Main polling loop - checks clipboard for changes."""
        poll_sec = self.config.get("poll_interval_ms", 300) / 1000.0

        while self.running:
            try:
                current_seq = get_clipboard_seq()
                if current_seq != self.last_seq:
                    self.last_seq = current_seq
                    current_text = get_clipboard_text()

                    if current_text and current_text != self.last_text:
                        self.last_text = current_text
                        self._process_clipboard_text(current_text)

            except Exception:
                traceback.print_exc()

            time.sleep(poll_sec)

    def _process_clipboard_text(self, text):
        """Process text from clipboard - check for commands or insert."""
        cmd_type, result = process_voice_commands(text)

        if cmd_type == "action":
            self._dispatch_uno_command(result)
        elif cmd_type == "text":
            self._insert_text(result)
        else:
            self._insert_text(text)

    def _get_active_document(self):
        """Get the active Writer document and its text cursor."""
        try:
            smgr = self.ctx.ServiceManager
            desktop = smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )
            doc = desktop.getCurrentComponent()
            if doc and doc.supportsService("com.sun.star.text.TextDocument"):
                controller = doc.getCurrentController()
                view_cursor = controller.getViewCursor()
                text = doc.getText()
                return doc, text, view_cursor
        except Exception:
            traceback.print_exc()
        return None, None, None

    def _insert_text(self, text):
        """Insert text at the current cursor position in the active document."""
        try:
            doc, doc_text, view_cursor = self._get_active_document()
            if doc_text and view_cursor:
                text_cursor = doc_text.createTextCursorByRange(
                    view_cursor.getStart()
                )
                if not view_cursor.isCollapsed():
                    text_cursor = doc_text.createTextCursorByRange(view_cursor)
                    doc_text.insertString(text_cursor, text, True)
                else:
                    doc_text.insertString(text_cursor, text, False)

                view_cursor.gotoEnd(False)
        except Exception:
            traceback.print_exc()

    def _dispatch_uno_command(self, cmd_url):
        """Execute a UNO dispatch command."""
        try:
            smgr = self.ctx.ServiceManager
            desktop = smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )
            frame = desktop.getCurrentFrame()
            if frame:
                dispatch = smgr.createInstanceWithContext(
                    "com.sun.star.frame.DispatchHelper", self.ctx
                )
                dispatch.executeDispatch(frame, cmd_url, "", 0, ())
        except Exception:
            traceback.print_exc()

    def _notify(self, message):
        """Show a notification in the status bar."""
        try:
            smgr = self.ctx.ServiceManager
            desktop = smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )
            doc = desktop.getCurrentComponent()
            if doc:
                controller = doc.getCurrentController()
                if controller:
                    sb = controller.getStatusBar()
                    if sb:
                        sb.setText(f"[Dragon Bridge] {message}")
        except Exception:
            pass


# Global bridge instance

_bridge_instance = None


def get_bridge(ctx):
    global _bridge_instance
    if _bridge_instance is None:
        _bridge_instance = ClipboardBridge(ctx)
    return _bridge_instance


# Toggle Bridge Command

class ToggleBridge(unohelper.Base, XJobExecutor, XServiceInfo):
    """Toggles the clipboard bridge on/off."""

    IMPLE_NAME = "org.fjccv.dragonbridge.ToggleBridge"
    SERVICE_NAMES = (IMPLE_NAME,)

    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        bridge = get_bridge(self.ctx)
        if bridge.is_running():
            bridge.stop()
        else:
            bridge.start()

    def getImplementationName(self):
        return self.IMPLE_NAME

    def supportsService(self, name):
        return name in self.SERVICE_NAMES

    def getSupportedServiceNames(self):
        return self.SERVICE_NAMES


# Settings Dialog

class SettingsDialog(unohelper.Base, XJobExecutor, XServiceInfo):
    """Settings configuration dialog."""

    IMPLE_NAME = "org.fjccv.dragonbridge.Settings"
    SERVICE_NAMES = (IMPLE_NAME,)

    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        try:
            self._open_settings()
        except Exception:
            traceback.print_exc()

    def _open_settings(self):
        config = load_config()
        smgr = self.ctx.ServiceManager

        dialog_model = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", self.ctx
        )
        dialog_model.Title = "Dragon Bridge Settings"
        dialog_model.Width = 280
        dialog_model.Height = 140

        # Poll interval
        lbl1 = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlFixedTextModel"
        )
        lbl1.Name = "lblPoll"
        lbl1.PositionX = 10
        lbl1.PositionY = 15
        lbl1.Width = 140
        lbl1.Height = 12
        lbl1.Label = "Clipboard poll interval (ms):"
        dialog_model.insertByName("lblPoll", lbl1)

        poll_field = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlNumericFieldModel"
        )
        poll_field.Name = "numPoll"
        poll_field.PositionX = 160
        poll_field.PositionY = 13
        poll_field.Width = 60
        poll_field.Height = 14
        poll_field.Value = config.get("poll_interval_ms", 300)
        poll_field.ValueMin = 100
        poll_field.ValueMax = 2000
        poll_field.DecimalAccuracy = 0
        dialog_model.insertByName("numPoll", poll_field)

        # Auto space checkbox
        chk_space = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlCheckBoxModel"
        )
        chk_space.Name = "chkAutoSpace"
        chk_space.PositionX = 10
        chk_space.PositionY = 40
        chk_space.Width = 260
        chk_space.Height = 12
        chk_space.Label = "Automatically add space before inserted text"
        chk_space.State = 1 if config.get("auto_space", True) else 0
        dialog_model.insertByName("chkAutoSpace", chk_space)

        # Show notifications checkbox
        chk_notify = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlCheckBoxModel"
        )
        chk_notify.Name = "chkNotify"
        chk_notify.PositionX = 10
        chk_notify.PositionY = 58
        chk_notify.Width = 260
        chk_notify.Height = 12
        chk_notify.Label = "Show status bar notifications"
        chk_notify.State = 1 if config.get("show_notifications", True) else 0
        dialog_model.insertByName("chkNotify", chk_notify)

        # Info text
        lbl_info = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlFixedTextModel"
        )
        lbl_info.Name = "lblSettingsInfo"
        lbl_info.PositionX = 10
        lbl_info.PositionY = 76
        lbl_info.Width = 260
        lbl_info.Height = 20
        lbl_info.MultiLine = True
        lbl_info.Label = (
            "Lower poll interval = faster response but more CPU usage. "
            "300ms is recommended for most systems."
        )
        dialog_model.insertByName("lblSettingsInfo", lbl_info)

        # OK button
        btn_ok = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlButtonModel"
        )
        btn_ok.Name = "btnOK"
        btn_ok.PositionX = 140
        btn_ok.PositionY = 108
        btn_ok.Width = 60
        btn_ok.Height = 22
        btn_ok.Label = "OK"
        btn_ok.PushButtonType = 1
        dialog_model.insertByName("btnOK", btn_ok)

        # Cancel button
        btn_cancel = dialog_model.createInstance(
            "com.sun.star.awt.UnoControlButtonModel"
        )
        btn_cancel.Name = "btnCancel"
        btn_cancel.PositionX = 210
        btn_cancel.PositionY = 108
        btn_cancel.Width = 60
        btn_cancel.Height = 22
        btn_cancel.Label = "Cancel"
        btn_cancel.PushButtonType = 2
        dialog_model.insertByName("btnCancel", btn_cancel)

        # Create and show
        dialog = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", self.ctx
        )
        dialog.setModel(dialog_model)
        toolkit = smgr.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", self.ctx
        )
        dialog.setVisible(False)
        dialog.createPeer(toolkit, None)

        result = dialog.execute()

        if result == 1:  # OK pressed
            num_poll = dialog.getControl("numPoll")
            chk_auto = dialog.getControl("chkAutoSpace")
            chk_ntfy = dialog.getControl("chkNotify")

            config["poll_interval_ms"] = int(num_poll.getValue())
            config["auto_space"] = chk_auto.getModel().State == 1
            config["show_notifications"] = chk_ntfy.getModel().State == 1
            save_config(config)

            bridge = get_bridge(self.ctx)
            bridge.config = config

        dialog.dispose()

    def getImplementationName(self):
        return self.IMPLE_NAME

    def supportsService(self, name):
        return name in self.SERVICE_NAMES

    def getSupportedServiceNames(self):
        return self.SERVICE_NAMES


# Component Registration

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    ToggleBridge,
    ToggleBridge.IMPLE_NAME,
    ToggleBridge.SERVICE_NAMES,
)

g_ImplementationHelper.addImplementation(
    SettingsDialog,
    SettingsDialog.IMPLE_NAME,
    SettingsDialog.SERVICE_NAMES,
  )
