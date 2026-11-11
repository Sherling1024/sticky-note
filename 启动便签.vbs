' Sticky Notes - silent launcher (no console flash)
Option Explicit
Dim fso, sh, dir, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = sh.ExpandEnvironmentStrings("%TEMP%")
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & dir & "\StickyNotes.ps1"""
sh.Run cmd, 0, False
