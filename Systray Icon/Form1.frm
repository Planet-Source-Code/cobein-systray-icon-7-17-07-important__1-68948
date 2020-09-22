VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Systray Icon Test"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Systary Icon"
      Height          =   1755
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   5535
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   60
         ScaleHeight     =   1455
         ScaleWidth      =   5415
         TabIndex        =   18
         Top             =   240
         Width           =   5415
         Begin VB.OptionButton optSysIcon 
            Caption         =   "Icon from resource (Compile the project to make it work)"
            Height          =   435
            Index           =   3
            Left            =   2100
            TabIndex        =   24
            Top             =   960
            Width           =   2655
         End
         Begin VB.OptionButton optSysIcon 
            Caption         =   "Icon from Dll"
            Height          =   315
            Index           =   2
            Left            =   2100
            TabIndex        =   23
            Top             =   600
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optSysIcon 
            Caption         =   "Icon from file"
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   22
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optSysIcon 
            Caption         =   "Icon from Picture"
            Height          =   315
            Index           =   0
            Left            =   300
            TabIndex        =   21
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Set Tooltip"
            Height          =   375
            Left            =   3300
            TabIndex        =   20
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   300
            TabIndex        =   19
            Text            =   "Tooltip"
            Top             =   180
            Width           =   2775
         End
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Another form"
      Height          =   435
      Left            =   5940
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Events"
      Height          =   4395
      Left            =   5940
      TabIndex        =   12
      Top             =   720
      Width           =   2715
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Balloon "
      Height          =   3075
      Left            =   240
      TabIndex        =   1
      Top             =   2100
      Width           =   5535
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   5235
         TabIndex        =   2
         Top             =   240
         Width           =   5235
         Begin VB.OptionButton optIcon 
            Caption         =   "Icon User"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   2340
            Width           =   1335
         End
         Begin VB.CommandButton cmdHideBall 
            Caption         =   "Hide"
            Height          =   375
            Left            =   900
            TabIndex        =   14
            Top             =   120
            Width           =   795
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2040
            TabIndex        =   11
            Text            =   "Title"
            Top             =   1140
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   2040
            TabIndex        =   10
            Text            =   "Text"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton cmdSetTitle 
            Caption         =   "Set Title"
            Height          =   375
            Left            =   4140
            TabIndex        =   9
            Top             =   1080
            Width           =   915
         End
         Begin VB.CommandButton cmdSetText 
            Caption         =   "Set Text"
            Height          =   375
            Left            =   4140
            TabIndex        =   8
            Top             =   1620
            Width           =   915
         End
         Begin VB.CommandButton cmdShowBall 
            Caption         =   "Show"
            Height          =   375
            Left            =   60
            TabIndex        =   7
            Top             =   120
            Width           =   795
         End
         Begin VB.OptionButton optIcon 
            Caption         =   "No Icon"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   900
            Width           =   1335
         End
         Begin VB.OptionButton optIcon 
            Caption         =   "Icon Info"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   1260
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optIcon 
            Caption         =   "Icon Warning"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   1620
            Width           =   1335
         End
         Begin VB.OptionButton optIcon 
            Caption         =   "Icon Error"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   3
            Top             =   1980
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox picIcon 
      Height          =   735
      Left            =   7740
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private WithEvents f_cSystray As cSystray
Attribute f_cSystray.VB_VarHelpID = -1

Private Sub cmdShowBall_Click()
    f_cSystray.BalloonShow True
End Sub

Private Sub cmdSetTitle_Click()
    f_cSystray.BalloonTitle = Text1
End Sub

Private Sub cmdSetText_Click()
    f_cSystray.BalloonText = Text2
End Sub

Private Sub cmdHideBall_Click()
    f_cSystray.BalloonShow False
End Sub

Private Sub cmdNew_Click()
    Dim frm As New Form1
    frm.Show
End Sub

Private Sub Command1_Click()
f_cSystray.SysTrayToolTip = Text3
End Sub

Private Sub f_cSystray_BalloonClick()
    Addthis "Balloon Click"
End Sub

Private Sub f_cSystray_BalloonClose()
    Addthis "Balloon Close"
End Sub

Private Sub f_cSystray_BalloonHide()
    Addthis "BalloonHide"
End Sub

Private Sub f_cSystray_BalloonShow()
    Addthis "BalloonShow"
End Sub

Private Sub f_cSystray_MouseDblClick(Button As Integer)
    Addthis "DblClick " & Button
End Sub

Private Sub f_cSystray_MouseDown(Button As Integer)
    Addthis "Mouse Down"
End Sub

Private Sub f_cSystray_MouseMove()
    Addthis "MouseMove"
End Sub

Private Sub f_cSystray_MouseUp(Button As Integer)
    Addthis "Mouse Up " & Button
    f_cSystray.BeforePopup
    PopupMenu mnuPopup
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    Set f_cSystray = New cSystray
    
    With f_cSystray
        If .IsBalloonCapable Then
            .BalloonIcon = TTIconInfo
            .BalloonText = "test"
            .BalloonTitle = "Title"
        End If
                .SysTrayIconFromCompRes "shell32.dll", 130
        .SysTrayToolTip = "Tooltip"
        .SysTrayShow True
        If .IsBalloonCapable Then .BalloonShow True, 3000
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set f_cSystray = Nothing
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub optIcon_Click(Index As Integer)
    f_cSystray.BalloonIcon = Index
End Sub

Private Sub Addthis(ByVal sData As String)
    List1.AddItem sData
    List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub optSysIcon_Click(Index As Integer)
    With f_cSystray
        Select Case Index
            Case 0: .SysTrayIconFromHandle picIcon.Picture
            Case 1: .SysTrayIconFromFile App.Path & "\xptray.ico"
            Case 2: .SysTrayIconFromCompRes "shell32.dll", 130
            Case 3: .SysTrayIconFromRes "ICON_0"
        End Select
    End With
End Sub
