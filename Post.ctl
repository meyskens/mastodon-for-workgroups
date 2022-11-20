VERSION 5.00
Begin VB.UserControl Post 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   ScaleHeight     =   1365
   ScaleWidth      =   4215
   Begin VB.CommandButton Boost 
      Caption         =   "Boost"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Favorite 
      Caption         =   "Favorite"
      Height          =   255
      Left            =   1680
      MaskColor       =   &H0080FFFF&
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label ContentLabel 
      Caption         =   "Content"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label UsernameLabel 
      Caption         =   "Username"
      DataField       =   "pUsername"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   120
      Picture         =   "Post.ctx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "Post"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pUsername As String
Private pContent As String

Public Property Get Username() As String
   Username = pUsername
End Property

Public Property Let Username(ByVal NewValue As String)
   pUsername = NewValue
   UsernameLabel.Caption = pUsername
   PropertyChanged "Username"
End Property

Public Property Get content() As String
   content = pContent
End Property

Public Property Let content(ByVal NewValue As String)
   pContent = NewValue
   ContentLabel.Caption = pContent
   PropertyChanged "Content"
End Property

Private Sub Command1_Click()

End Sub

Private Sub UserControl_Initialize()
    ContentLabel.Caption = pContent
    UsernameLabel.Caption = pUsername
End Sub

