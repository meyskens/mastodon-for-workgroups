VERSION 5.00
Begin VB.Form newToot 
   Caption         =   "New Toot"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Visibility 
      Height          =   315
      ItemData        =   "newToot.frx":0000
      Left            =   1680
      List            =   "newToot.frx":0002
      TabIndex        =   3
      Text            =   "Visibiliy"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton SendToot 
      Caption         =   "Toot!"
      DragIcon        =   "newToot.frx":0004
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MouseIcon       =   "newToot.frx":030E
      Picture         =   "newToot.frx":0618
      TabIndex        =   1
      ToolTipText     =   "Send a post to your Mastodon account"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Content 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label CharsLeft 
      Caption         =   "0 / 500"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "newToot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Visibility.AddItem "public"
   Visibility.AddItem "unlisted"
   Visibility.AddItem "private"
   Visibility.Text = Visibility.List(0)   ' Display first item.

End Sub

Private Sub SendToot_Click()
    Dim hInternetSession    As Long
    Dim hInternetConnect    As Long
    Dim hHttpOpenRequest    As Long
    Dim sBuffer             As String
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim scUserAgent         As String
    Dim bDoLoop             As Boolean
    
    SendToot.Enabled = False
    Screen.MousePointer = vbHourglass
    
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
                            
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, _
                                        "POST", _
                                        "/api/v1/statuses", _
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
    
    Dim contentType As String
    contentType = "Content-Type: application/x-www-form-urlencoded"
    HttpAddRequestHeaders hHttpOpenRequest, _
                            contentType, _
                            Len(contentType), _
                            HTTP_ADDREQ_FLAG_ADD
                            
    Dim formData As String
    formData = "status=" & URLEncode(Content.Text) & "&visibility=" & Visibility.Text
    
    HttpSendRequest hHttpOpenRequest, vbNullString, 0, ByVal formData, Len(formData)
    
    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend

    InternetCloseHandle (hInternetSession)
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Content_Change()
    CharsLeft.Caption = Len(Content.Text) & " / 500"
    
    If Len(Content.Text) > 500 Then
        CharsLeft.ForeColor = &HFF&
        SendToot.Enabled = False
    Else
        CharsLeft.ForeColor = vbWindowText
        If Len(Content.Text) > 0 Then
            SendToot.Enabled = True
        Else
            SendToot.Enabled = False
        End If
    End If
End Sub

