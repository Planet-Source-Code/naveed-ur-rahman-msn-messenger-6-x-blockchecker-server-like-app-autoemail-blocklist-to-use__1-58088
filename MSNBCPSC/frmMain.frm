VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN Messenger Blockchecker"
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sck2 
      Index           =   0
      Left            =   5325
      Top             =   3465
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckEmail 
      Index           =   0
      Left            =   5865
      Top             =   3435
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetWant2Say 
      Left            =   2070
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSay 
      Caption         =   "Hi neeno, I &want to say..."
      Height          =   390
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5430
      Width           =   6375
   End
   Begin VB.Timer tmrBuddyIcon 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   3180
      Top             =   2640
   End
   Begin MSComctlLib.ImageList ImgLstMain 
      Left            =   825
      Top             =   2745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCB
            Key             =   "#signing1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1267
            Key             =   "#signing2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1803
            Key             =   "#member"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24DF
            Key             =   "#error"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A7B
            Key             =   "#completed"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3757
            Key             =   "#mail"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CF3
            Key             =   "#nomail"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":428F
            Key             =   "#done"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckMain 
      Index           =   0
      Left            =   2925
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   120
      TabIndex        =   7
      Top             =   5910
      Width           =   6375
   End
   Begin VB.Frame FrameLine 
      Height          =   75
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   6375
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2670
      Left            =   120
      TabIndex        =   0
      Top             =   1935
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4710
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImgLstMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FrameHelp 
      Height          =   1710
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   6375
      Begin VB.ComboBox cmbSMTP 
         Height          =   315
         ItemData        =   "frmMain.frx":462B
         Left            =   2850
         List            =   "frmMain.frx":4638
         TabIndex        =   12
         Text            =   "cmbSMTP"
         Top             =   1245
         Width           =   3255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Also specify SMTP server:"
         Height          =   195
         Left            =   915
         TabIndex        =   11
         Top             =   1305
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "This computer (server) IP (hostname):"
         Height          =   195
         Left            =   915
         TabIndex        =   10
         Top             =   960
         Width           =   2640
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmMain.frx":466F
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblHelp 
         Caption         =   $"frmMain.frx":533A
         Height          =   1035
         Index           =   0
         Left            =   915
         TabIndex        =   3
         Top             =   255
         Width           =   5295
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Please, don't hasitate in sending your comments, feedbacks and bug-reports:"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   5115
      Width           =   5430
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   -30
      Picture         =   "frmMain.frx":53D7
      Top             =   -30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright(c) 1999-2005, Naveed Software. All rights reserved."
      Height          =   195
      Left            =   195
      TabIndex        =   8
      Top             =   6075
      Width           =   4365
   End
   Begin VB.Label lblHyperLink 
      AutoSize        =   -1  'True
      Caption         =   "neenojee@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   2970
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4845
      Width           =   1665
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Programmed by Naveed ur Rehman"
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   5
      Top             =   4845
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   
    If App.PrevInstance = True Then End
    'One application = One time
    
    
On Error GoTo ErrorOccured

    'The data recording is in ListView control.
    'You can also use database programming.
    'well, I was not interested in doing fun
    'by recording other people IDs, their
    'contact lists and blocklists etc.
    'But hay !,
    'You can do many other good things too !!!
    'like using data fot some server side scripting
    'etc.
    'OK - read the codes...
        
    Dim w As Long, z As Integer
    
    w = LV1.Width / 4
    LV1.ColumnHeaders.Add , , "Sign-in Name", w
    LV1.ColumnHeaders.Add , , "Progress", w
    LV1.ColumnHeaders.Add , , "Nick", w
    LV1.ColumnHeaders.Add , , "Blocked by", LV1.Width * 5
    LV1.ColumnHeaders.Add , , "Date/Time", 0
    
    'hmmm, I think, describing each and every line
    'of codes is not good for your health !!!
    'It will make you LAZY :D
    
    For z = 0 To lblHyperLink.Count - 1
        lblHyperLink(z).MousePointer = 99
        lblHyperLink(z).MouseIcon = imgHand.Picture
    Next z
    
    'Please don't worry about the following label control.
    Label4.Caption = Label4.Caption & sckMain(0).LocalIP & " (" & sckMain(0).LocalHostName & ")"
    
    'Yeh ! selecting the SMTP server automatically.
    'I have given a few servers over there and
    'have tested, all are working !!!
    cmbSMTP.ListIndex = 0
    
    sckMain(0).Close
    sckMain(0).LocalPort = 20801
    sckMain(0).Listen
    
    'Hello :D call me babe call me ! I am waiting for
    'your hoooooooorrrrrraaaaaaa !!!
    
    Exit Sub

ErrorOccured:
    
    'ooOOppSss !
    MsgBox Err.Description & vbCrLf & vbCrLf & "Please click OK to terminating program.", vbCritical, Me.Caption
    End
    
End Sub

Private Sub sckMain_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
On Error Resume Next

    'My babez connected !!!
    'hehehe - counting babez number :D
    Dim NewIndex As Long
    
    NewIndex = sckMain.Count
    
    'Ok... lemme load my controls !!!
    'Just hold on...
    
    Load sckMain(NewIndex)  'This will accept the request
    Load sck2(NewIndex)     'Will connect to msn server
    Load sckEmail(NewIndex) 'For emailing
    
    Load tmrBuddyIcon(NewIndex) 'nah ! just for animation.
    
    'Adding user in the listview
    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Add(, "#" & NewIndex, "New Account", , "#member")
        LITEM.SubItems(1) = "Wait..."   'Oh Babe - Just hold on !
        LITEM.SubItems(2) = ""
        LITEM.SubItems(3) = ""
    
    With sckMain(NewIndex)
        .Close
        .Tag = ""
        .Accept requestID   'Accepted MSN messengers request
    End With
    
    'Connecting to msn server
    sck2(NewIndex).Connect "207.46.106.99", 1863
    
End Sub

Private Sub sckMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error Resume Next

    Dim DATA As String, _
        SOCK4String As String
    
    SOCK4String = Chr(4) & Chr(1) & Chr(7) & Chr(71) & Chr(207) & Chr(46) & Chr(104) & Chr(20) & Chr(0)
    
    sckMain(Index).GetData DATA
    'Tren Tren - MSN Messenger speaking...
    'Can you hear me holla !!!
    
    DATA = Replace(DATA, SOCK4String, "")

    If sck2(Index).State <> sckConnected Then
        'OMG ! If you get big big traffic
        'on server then try shaking your head
        'and think some good technique here.
        'hmmm... hint !
        'Initialized some collection or dictionary
        'as global and add sending data
        'and keep it stocking until sck2 connects
        'to its destination.
        'Again, I never have big traffic
        'thats why I dont need to shake my head.
        
        While sck2(Index).State <> sckConnected
        DoEvents
        Wend
    End If
    
    'OK, now sending to messenger server
    sck2(Index).SendData DATA

End Sub

Private Sub sckEmail_Close(Index As Integer)
    sckEmail(Index).Close
    
    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)
    
    If Val(sckEmail(Index).Tag) >= 6 Then
        LITEM.SmallIcon = "#done"
        LITEM.SubItems(1) = "Completed !!!"
    Else
        LITEM.SubItems(1) = "Fail to send email."
        LITEM.SmallIcon = "#nomail"
    End If
End Sub

Private Sub sckEmail_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error GoTo ErrorOccured

    Dim DATA As String, nData As Integer
    Dim st  As Integer
    Dim SND As String, TXT As String
    
    sckEmail(Index).GetData DATA
    st = Val(sckEmail(Index).Tag)
    sckEmail(Index).Tag = st + 1
    
    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)
    
    nData = Val(Left(DATA, 3))
    
    If nData = 250 Or nData = 220 Or nData = 354 Or nData = 221 Then
        
        Select Case st
            Case 0:
                SND = "HELO neenojee"
            Case 1:
                SND = "MAIL FROM:neenojee@hotmail.com"
            Case 2:
                SND = "RCPT TO:" & LITEM.text
            Case 3:
                SND = "DATA"
            Case 4:
            
            'OH-KAY !
            'Feel free to change  following :D
            'I love changing codes and
            'realeasing them by my name too
            'BUT
            'I havent done such a thing YET !
            'coz my mom said,
            'The bad thing is bad
            'and bad do bad things
            'If you do something bad
            'I will give you a smack :D
            'But
            'You can do BAD things !!!
            ' And even if you don't know
            'that what bad thing can you do
            'from following then
            'its simple !!!
            'Press <ctrl>+F
            'and replace "Naveed ur Rehman" by your name
            'and "neenojee" by your ID !!!
            'come on, I willn't mind !
            
            TXT = Me.Caption & vbCrLf _
                   & "Programmed by Naveed ur Rehman (neenojee@hotmail.com)" & vbCrLf _
                   & "" & vbCrLf _
                   & "User ID: " & LITEM.text & vbCrLf _
                   & "Nick: " & LITEM.SubItems(2) & vbCrLf _
                   & "Date/Time: " & LITEM.SubItems(4) & vbCrLf _
                   & "" & vbCrLf _
                   & "This ID is blocked by the following ID(s):" & vbCrLf _
                   & Replace(LITEM.SubItems(3), "    ", vbCrLf) & vbCrLf & vbCrLf & _
                   "Thank you for using our services."

                SND = "From:Naveed ur Rehman <neenojee@hotmail.com>" & vbCrLf & _
                      "To:" & LITEM.text & vbCrLf & _
                      "Subject:Your MSN Messenger Blocklist" & vbCrLf & _
                      "Reply-To:Naveed ur Rehman <neenojee@hotmail.com>" & vbCrLf & vbCrLf & _
                      TXT & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "."
            
            Case 5:
                SND = "QUIT"
                LITEM.SmallIcon = "#done"
                LITEM.SubItems(1) = "Completed !!!"
            
            Case 6:
                SND = ""
                sckEmail(Index).Close
                LITEM.SmallIcon = "#done"
                LITEM.SubItems(1) = "Completed !!!"
                
        End Select
    
        If SND <> "" Then SND = SND & vbCrLf
        If SND <> "" Then sckEmail(Index).SendData SND
        
    Else

ErrorOccured:
    sckEmail(Index).Close
    LITEM.SubItems(1) = "Fail to send email."
    'So man, select some other SMTP
    LITEM.SmallIcon = "#nomail"
    
    End If

End Sub

Private Sub sckEmail_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    sckEmail(Index).Close
    
    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)
    LITEM.SubItems(1) = "Fail to send email (error: " & Description & ")"
    'I said, select some other SMTP
    LITEM.SmallIcon = "#nomail"
    
End Sub


Private Sub sck2_Connect(Index As Integer)

On Error Resume Next

    Dim SOCK4String As String
    
    SOCK4String = Chr(4) & "Z" & String(6, 0)
    
    tmrBuddyIcon(Index).Enabled = True
    
    sckMain(Index).SendData SOCK4String
    
End Sub

Private Sub sck2_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error Resume Next

    Dim DATA As String
    
    sck2(Index).GetData DATA
    sckMain(Index).SendData DATA

    Dim SPLN As Variant
    Dim LN, LN2
    
    SPLN = Split(DATA, vbCrLf)

    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)
    
    Dim StatusCode As Integer
    Dim Identity_ As String
    
    For Each LN In SPLN
    
    'Broken
    If Left(LN, 4) <> UCase(Left(LN, 4)) Then
    LN = LN2 & LN
    End If
    
    'e.g. neenojee@hotmail.com
    If RetPar(LN, 1, " ") = "USR" And RetPar(LN, 3, " ") = "OK" Then
        LITEM.SubItems(1) = "User Name OK"
        LITEM.text = RetPar(LN, 4, " ")
    End If
    
    'e.g. My nick is Neeno
    If RetPar(LN, 1, " ") = "PRP" And RetPar(LN, 2, " ") = "MFN" Then
        LITEM.SubItems(1) = "User Nick OK"
        LITEM.SubItems(2) = Code2Normal(RetPar(LN, 3, " "))
    End If
    
    'e.g. LST N=shoaibalam24@hotmail.com F=maast%20rahoo%20masti%20main%20aaga%20laage%20bastti%20main. C=ca6bc026-2f69-412e-9210-01d80037b8aa 11 a7193d93-9a81-4b0a-ae68-d77c38e235a8
    If RetPar(LN, 1, " ") = "LST" Then
        
        Identity_ = RetVal(RetPar(LN, 2, " "))
        StatusCode = Val(RetPar(LN, 5, " "))
        
        LITEM.SubItems(1) = "Buddy: " & Identity_ & " " & GetMSNStatusCodeMeaning(StatusCode)
        'You can record buddies if you like.
        
        If StatusCode = 3 Then 'blocked CATCHA !!!
            LITEM.SubItems(3) = LITEM.SubItems(3) & "<" & Trim(Identity_) & ">    "
        End If
        
        sck2(Index).Tag = "OK" 'yo yo man !
        'MSN client is not reading from buffer
        
    End If
    
    If RetPar(LN, 1, " ") = "MSG2" Or _
       RetPar(LN, 1, " ") = "CHG" Then
    'If you know more then add in this else
    'it works perfect !

        sck2(Index).Close
        sckMain(Index).Close
        
        
        tmrBuddyIcon(Index).Enabled = False
        LITEM.SmallIcon = "#completed"
        LITEM.SubItems(1) = "Emailing..."
        LITEM.SmallIcon = "#mail"
        
        'Its my dating style !
        LITEM.SubItems(4) = DatePart("d", Now) & " " & MonthName(DatePart("m", Now), True) & ", " & Year(Now) & " (" & Time & ")"
        
        If sck2(Index).Tag = "" Then _
        LITEM.SubItems(3) = "(Sorry, MSN has started using its cache)"
        'bulshit !
        
        With sckEmail(Index)
            .Close
            .RemoteHost = cmbSMTP.text  'Best one is alredy selected
            .RemotePort = 25    'Well, you can specify other if you know
            .Connect
        End With
        
    End If

LN2 = LN    'Broken Setup
Next

End Sub

Private Sub tmrBuddyIcon_Timer(Index As Integer)
    
    If tmrBuddyIcon(Index).Tag = "" Then tmrBuddyIcon(Index).Tag = "#signing2"
    
    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)

    If tmrBuddyIcon(Index).Tag = "#signing1" Then
        tmrBuddyIcon(Index).Tag = "#signing2"
    Else
        tmrBuddyIcon(Index).Tag = "#signing1"
    End If
    
    LITEM.SmallIcon = tmrBuddyIcon(Index).Tag

    If sck2(Index).State = sckClosed Or sckMain(Index).State = sckClosed Then tmrBuddyIcon(Index).Enabled = False
    
End Sub

Private Sub sckMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   
    tmrBuddyIcon(Index).Enabled = False
    
    sckMain(Index).Close
    sck2(Index).Close
    
    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)
    
    LITEM.SmallIcon = "#error"
    LITEM.SubItems(1) = "Error: " & Description
    
End Sub

Private Sub sck2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    tmrBuddyIcon(Index).Enabled = False
    
    sckMain(Index).Close
    sck2(Index).Close

    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)

    LITEM.SmallIcon = "#error"
    LITEM.SubItems(1) = "Error: " & Description

End Sub

Private Sub sckMain_Close(Index As Integer)
    
    tmrBuddyIcon(Index).Enabled = False
    
    sckMain(Index).Close
    sck2(Index).Close

    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)

    LITEM.SmallIcon = "#error"
    LITEM.SubItems(1) = "Closed"

End Sub

Private Sub sck2_Close(Index As Integer)
    
    tmrBuddyIcon(Index).Enabled = False
    
    sckMain(Index).Close
    sck2(Index).Close

    Dim LITEM As ListItem
    Set LITEM = LV1.ListItems.Item("#" & Index)

    LITEM.SmallIcon = "#error"
    LITEM.SubItems(1) = "Closed"

End Sub

Private Sub lblHyperLink_Click(Index As Integer)
    
    If Index = 0 Then   'Email
        RunHyperlink "mailto:neenojee@hotmail.com?title=MSNBC"
    End If

End Sub

Private Sub LV1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorOccured
    
    If Button = vbRightButton And LV1.ListItems.Count > 0 And LV1.SelectedItem.SubItems(3) <> "" And LV1.SelectedItem.Index <> -1 Then PopupMenu mnuFile

ErrorOccured:
End Sub

Private Sub mnuCopy_Click()
On Error GoTo ErrorOccured
    
    Dim ID_ As String, _
        Nick_ As String, _
        BlockedBy_ As String, _
        DT_ As String
    ID_ = LV1.SelectedItem.text
    Nick_ = LV1.SelectedItem.SubItems(2)
    BlockedBy_ = LV1.SelectedItem.SubItems(3)
    DT_ = LV1.SelectedItem.SubItems(4)
    
    Clipboard.Clear
    Dim TXT As String
    
    TXT = Me.Caption & vbCrLf _
                   & "Programmed by Naveed ur Rehman (neenojee@hotmail.com)" & vbCrLf _
                   & "" & vbCrLf _
                   & "User ID: " & ID_ & vbCrLf _
                   & "Nick: " & Nick_ & vbCrLf _
                   & "Date/Time: " & DT_ & vbCrLf _
                   & "" & vbCrLf _
                   & "This ID is blocked by the following ID(s):" & vbCrLf _
                   & Replace(BlockedBy_, "    ", vbCrLf)
    
    Clipboard.SetText TXT
    
Exit Sub
ErrorOccured:
End Sub

Private Sub cmdSay_Click()
Dim IPT, i

cmdSay.Enabled = False
IPT = InputBox("Hello neeno," & vbCrLf & vbCrLf & "I want to say...,", "Say To Neeno", "(your-email)...(message)")
If IPT <> "" And UCase(IPT) <> UCase("(your-email)...(message)") Then
    IPT = Left(HexFormat(IPT), 255)
    cmdSay.Caption = "Sending... Please wait..."
    i = InetWant2Say.OpenURL(SendASP & "?a=2&t=" & IPT)
    If i <> "" Then
        MsgBox "Thank you very much." & vbCrLf & "Your message has been sent successfully.", vbInformation, "Message Sent"
    Else
        MsgBox "Your message couldn't be sent.", vbInformation, Me.Caption
        cmdSay.Caption = "Hi neeno, I &want to say..."
        cmdSay.Enabled = True
    End If
Else
        cmdSay.Caption = "Hi neeno, I &want to say..."
        cmdSay.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
