VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{59398BDA-D1AC-11D3-BB80-0080C86D4E64}#7.0#0"; "CCCSMS.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1695
      TabIndex        =   9
      Text            =   "Text3"
      ToolTipText     =   "letters left till 160 char limit"
      Top             =   465
      Width           =   1530
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   4215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3300
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   1110
      Left            =   105
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Text2"
      ToolTipText     =   "status stuff"
      Top             =   1650
      Width           =   3195
   End
   Begin VB.TextBox SMS 
      Height          =   660
      Left            =   105
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Main.frx":0000
      ToolTipText     =   "text to send (and coded sms text after send button pressed)"
      Top             =   900
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send SMS"
      Height          =   615
      Left            =   2235
      TabIndex        =   4
      Top             =   915
      Width           =   1080
   End
   Begin VB.TextBox GSMNumber 
      Height          =   315
      Left            =   1665
      TabIndex        =   3
      Text            =   "8957168"
      ToolTipText     =   "gsm number"
      Top             =   75
      Width           =   1560
   End
   Begin VB.TextBox OwnNumber 
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Text            =   "000001"
      ToolTipText     =   "Sent from number"
      Top             =   90
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Text            =   "Combo1"
      ToolTipText     =   "Com Port selection"
      Top             =   465
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4395
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3630
      Width           =   3195
   End
   Begin CCCSMSSMSCControl.SMS SMS1 
      Height          =   480
      Left            =   1455
      TabIndex        =   8
      Top             =   2130
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label status 
      BorderStyle     =   1  'Fixed Single
      Height          =   2130
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "more status stuff"
      Top             =   2805
      Width           =   3180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'OK, I realize this code is EXTREMELY bad :)
'it was just meant as a test code for the actual
'component the SMS ocx.
'You have full permission to use the ocx in all your projects as this one
'is text only, i.e. you can send english text messages ONLY
'All foreign characters are completely ignored.
'Cheers, Icecube Ryder
'icecube.ryder@technologist.com





Sub dialsmsc()
' This just dials the SMSC and writes the status to a status window
If SentYet Then Exit Sub
status.Caption = vbCrLf + status.Caption + Format(Now, "  HH:MM:SS") + " Calling SmallMessageSystemCenter..  (SMSC)" + vbCrLf
status.Caption = status.Caption + Format(Now, "  HH:MM:SS") + " Using COMM port:" + MDMCommport + vbCrLf
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.RTSEnable = True
    MSComm1.SThreshold = 1
    MSComm1.RThreshold = 1
    MSComm1.PortOpen = True
    MSComm1.Output = "ATDT 9541010" + vbCrLf
 End Sub

Private Sub Hangup()
' Hangs up the Phoneline, and releases the com port
 status.Caption = status.Caption + Format(Now, "   HH:MM:SS") + " Hanging up SMSC" + vbCrLf
    Dim ret As Boolean
    If MSComm1.PortOpen Then
    
    On Error Resume Next
    'MSComm1.Output = "ATH"    ' Send hangup string
    End If

    ' If port is actually still open, then close it
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
   
    ' Notify user of error
    If Err Then MsgBox Error$, 48
    On Error GoTo 0
End Sub
Sub sendsms()

If SentYet Then
status.Caption = Format(Now, "   HH:MM:SS") + " Already gone..." & vbCrLf
'Hangup
Else
' If the port is still open, send the contents of the coded SMS.text window
If MSComm1.PortOpen Then
status.Caption = status.Caption + Format(Now, "   HH:MM:SS") + " Sending SMS Message..." & vbCrLf
MSComm1.Output = SMS + vbCr

End If
End If
End Sub
Private Sub Combo1_Change()
' New COM port selected.
Dim eaa As Integer
eaa = InStr(1, "COM", Combo1.Text, vbString)
MDMCommport = Mid$(Combo1.Text, eaa + 4, 1)
Debug.Print "Instring start "; eaa, "Comport "; MDMCommport

MDMDevice = Mid$(Combo1.Text, 11, Len(Combo1.Text))
        With MSComm1
            .CommPort = Val(Right$(MDMCommport, 1))
            .Handshaking = 1
            .RThreshold = 2
            .RTSEnable = True
            .Settings = "9600,n,8,1"
            .SThreshold = 1
            '.PortOpen = True
         End With
End Sub

Private Sub Command1_Click()
On Error GoTo errorhandler
status.Caption = ""

Dim eaa As Integer
eaa = InStr(1, Combo1.Text, "COM", vbTextCompare)
MDMCommport = Mid$(Combo1.Text, eaa + 4, 1)
Text2.Text = MDMCommport
If eaa = 0 Then MDMCommport = "1"
Debug.Print eaa, MDMCommport

MDMDevice = Mid$(Combo1.Text, eaa + 4, Len(Combo1.Text))
Text2.Text = "Commport: " + MDMCommport + "." + MDMDevice
        
        With MSComm1
            .CommPort = Val(MDMCommport)
            .Handshaking = 1
            .RThreshold = 12
            .RTSEnable = True
            .Settings = "9600,n,8,1"
            .SThreshold = 1
            '.PortOpen = True
            .InBufferCount = 0
         End With
         
' sms1.ownnumber is the senders number
SMS1.OwnNumber = OwnNumber.Text
' sms1.gsmnumber is the recipients number
SMS1.GSMNumber = GSMNumber.Text
' sms1.smsmessage is the plain text message
SMS1.SMSMessage = SMS.Text
' sms1.calcsms converts the plain text message to the o52 message
SMS1.CalcSMS
' sms1.smscmessage is the converted message, ready to be sent via mscomm
SMS.Text = SMS1.SMSCMessage

aa = 0
Timer1.Enabled = True
AckString = "CONNECT"
errorhandler:
If Err = 0 Then Exit Sub
'If Err = 5 Then Exit Sub

MsgBox (Str$(Err.Number) + "." + Err.Description)

End Sub


Private Sub Form_Load()
On Error GoTo errorhandler

 Dim Res As String
 Dim i As Long
 Dim res1 As String
 Dim res2 As String
 Dim res3 As String
 
 For i = 1 To 10
 Combo1.AddItem "COM" + Str$(i)
 Next i
 
 Combo1.ListIndex = 2
 
MDMCommport = Mid$(Combo1.Text, eaa + 3, 1)

Debug.Print eaa, MDMCommport

MDMDevice = Mid$(Combo1.Text, 11, Len(Combo1.Text))
        With MSComm1
            .CommPort = Val(Right$(MDMCommport, 1))
            .Handshaking = 2
            .RThreshold = 12
            .RTSEnable = True
            .Settings = "9600,n,8,1"
            .SThreshold = 1
            '.PortOpen = True
         End With
errorhandler:
 If Err.Number = 380 Then
 Combo1.AddItem " No Modems Detected"
 End If
 Err.Number = 0
 End Sub

Private Sub MSComm1_OnComm()
 Dim InBuff As String
 
 Dim EVMsg$
    Dim ERMsg$
    
    ' Branch according to the CommEvent property.
    Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
        If MSComm1.PortOpen Then
           InBuff = MSComm1.Input
           Call HandleInput(InBuff)
        End If
        Case comEvSend
        Case comEvCTS
            EVMsg$ = " Clear to send"
        Case comEvDSR
            EVMsg$ = " Change in DSR Detected"
        Case comEvCD
        sendsms
            EVMsg$ = " Carrier Status Toggled"
        Case comEvRing
            EVMsg$ = " The Phone is Ringing"
        Case comEvEOF
            EVMsg$ = " End of File Detected"

        ' Error messages.
        Case comBreak
            ERMsg$ = " Break Received"
            MSComm1.Break = False
            ERMsg$ = ""
        Case comCDTO
            ERMsg$ = " Carrier Detect Timeout"
        Case comCTSTO
            ERMsg$ = " CTS Timeout"
        Case comDCB
            ERMsg$ = " Error retrieving DCB"
        Case comDSRTO
            ERMsg$ = " DSR Timeout"
        Case comFrame
            ERMsg$ = " Framing Error"
        Case comOverrun
            ERMsg$ = " Overrun Error"
        Case comRxOver
            ERMsg$ = " Receive Buffer Overflow"
        Case comRxParity
            ERMsg$ = " Parity Error"
        Case comTxFull
            MSComm1.OutBufferCount = 0
            ERMsg$ = ""
        Case Else
            ERMsg$ = " Unknown error or event"
    End Select
    
    If Len(EVMsg$) Then
        ' Display event messages in the status bar.
        status.Caption = Format(Now, "   HH:MM:SS") + EVMsg$ & vbCrLf
                
        ' Enable timer so that the message in the status bar
        ' is cleared after 2 seconds

        
    ElseIf Len(ERMsg$) Then
        ' Display event messages in the status bar.
      '  sbrStatus.Panels("Status").Text = "Status: " & ERMsg$
        
        ' Display error messages in an alert message box.
        Beep
        ret = MsgBox(ERMsg$, 1, "Click Cancel to quit, OK to ignore.")
        
        ' If the user clicks Cancel (2)...
        If ret = 2 Then
            MSComm1.PortOpen = False    ' Close the port and quit.
        End If
 
    End If
End Sub

Private Sub SMS_Change()
' just so you won't type too much for the smsc to handle

Text3.Text = "Left:" + Str$(150 - Len(SMS.Text))

End Sub

Private Sub Timer1_Timer()
' times out after 40 seconds

aa = aa + 1
status.Caption = status.Caption + Str$(aa)
If aa = 1 Then dialsmsc

If aa = 40 Then
Timer1.Enabled = False
Hangup
Else
End If

End Sub
Sub HandleInput(InBuff As String)
status.Caption = Format(Now, "   HH:MM:SS") + "Message from Modem " + InBuff
If InStr(1, InBuff, "42", vbTextCompare) Then
sendsms
End If
If InStr(1, InBuff, "TON", vbTextCompare) Then
End If

End Sub
