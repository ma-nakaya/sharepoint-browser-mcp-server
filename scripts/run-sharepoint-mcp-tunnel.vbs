Option Explicit

Dim shell
Dim profileDirectory
Dim command
Dim exitCode

Set shell = CreateObject("WScript.Shell")
profileDirectory = shell.ExpandEnvironmentStrings("%APPDATA%\tunnel-client")

command = """C:\Apps\TunnelClient\tunnel-client.exe""" _
    & " run --profile sharepoint-browser-chatgpt" _
    & " --profile-dir=""" & profileDirectory & """" _
    & " --log.file=""C:\Apps\TunnelClient\sharepoint-mcp-tunnel.log"""

exitCode = shell.Run(command, 0, True)
WScript.Quit exitCode
