VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form main 
   Caption         =   "Mastodon 3.11 for Workgroups"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin MFW.Post Post1 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4215
      _ExtentX        =   5953
      _ExtentY        =   2566
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      Left            =   4440
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5190
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton refreshbt 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4215
   End
   Begin VB.Frame buttonframe 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton debug 
         Caption         =   "HTTP Debug"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton addbt 
         Caption         =   "New Toot"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mastodon 3.11 for Workgroups the Windows 9x Mastodon Client
' Copyright 2022 Maartje Eyskens
' inspired by awsom, the vb6 mastodon client: reds, 2018

Option Explicit
Dim statusPanel As Panel ' status panel
Dim posts() As Object ' list of rendered posts
Dim oldScrollValue As Integer ' make scrolling work
 
Private Sub SetStatus(status As String)
    statusPanel.Text = status
End Sub

Private Sub addbt_Click()
    newToot.Show
End Sub

Private Function cleanStatus(ByVal Content As String) As String
    Content = Replace(Content, "<p>", "")
    Content = Replace(Content, "</p>", "")
    cleanStatus = Content
End Function


Private Function loadTimeline() As JsonBag
    Dim hInternetSession    As Long
    Dim hInternetConnect    As Long
    Dim hHttpOpenRequest    As Long
    Dim sBuffer             As String
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim scUserAgent         As String
    Dim bDoLoop             As Boolean
    
    SetStatus "Opening connection..."

    hInternetSession = InternetOpen(scUserAgent, _
                            INTERNET_OPEN_TYPE_PRECONFIG, _
                            vbNullString, _
                            vbNullString, _
                            0)
    hInternetConnect = InternetConnect(hInternetSession, _
                            GetInstance(), _
                            INTERNET_DEFAULT_HTTP_PORT, _
                            vbNullString, _
                            vbNullString, _
                            INTERNET_SERVICE_HTTP, _
                            0, _
                            0)
    SetStatus "Sending request..."
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                                        "GET", _
                                        "/api/v1/timelines/home", _
                                        "HTTP/1.1", _
                                        vbNullString, _
                                        0, _
                                        INTERNET_FLAG_RELOAD, _
                                        0)
    Dim authHeader As String
    authHeader = "Authorization: Bearer " & GetToken() & vbCrLf
    
    HttpAddRequestHeaders hHttpOpenRequest, _
                            authHeader, _
                            Len(authHeader), _
                            HTTP_ADDREQ_FLAG_ADD
    HttpSendRequest hHttpOpenRequest, vbNullString, 0, 0, 0
    
    SetStatus "Reading data..."
    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
    SetStatus "Parsing data..."
    
    Dim JB
    Set JB = New JsonBag
    JB.JSON = sBuffer
    InternetCloseHandle (hInternetSession)
    
    SetStatus "Rendering..."
    
    Dim counter As Integer
    counter = 1
    While JB.Count >= counter And counter <= 20
        ' MsgBox JB.Item(counter).Item("content"), vbInformation, "Mastodon 3.11 for Workgroups"
        ReDim Preserve posts(counter)
        Set posts(counter) = Controls.Add("MFW.Post", "dynpost" & counter)
        posts(counter).Width = 4215
        posts(counter).Top = 1800 + 1455 * (counter - 1)
        posts(counter).Left = 120
        posts(counter).Height = 1455
        posts(counter).Visible = True
        posts(counter).Username = JB.Item(counter).Item("account").Item("acct")
        posts(counter).Content = cleanStatus(JB.Item(counter).Item("content"))
        
        counter = counter + 1
    Wend
    
    SetStatus "Cuddle a Blahaj"

End Function

Private Sub debug_Click()
    frmHttpQuery.Show
End Sub

Private Sub refreshbt_Click()
    Screen.MousePointer = vbHourglass
    VScroll1.Value = 0
    
    If GetInstance() <> "" Then
        loadTimeline
    Else
        loginform.Show
        main.Hide
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Post1.Visible = False ' post1 is used to demo in the form editor
    StatusBar.Panels.Clear ' clear default panels
    Set statusPanel = StatusBar.Panels.Add()

    With VScroll1 ' set up scrollbad TODO: make math better
        .Min = 0
        .Max = 20000
        .SmallChange = Screen.TwipsPerPixelY * 10
        .LargeChange = .SmallChange
    End With
End Sub


Private Sub VScroll1_Change()
    Dim eachctl As Control
    For Each eachctl In Me.Controls
        If TypeOf eachctl Is Post Then
            eachctl.Top = eachctl.Top + oldScrollValue - VScroll1.Value
        End If
    Next
    oldScrollValue = VScroll1.Value

    
End Sub
