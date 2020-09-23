VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emprison's ApiSpy"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5741
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Api Spy"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Code Generator"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Color Spy"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame1 
         Caption         =   "Code"
         Height          =   2775
         Left            =   -74940
         TabIndex        =   25
         Top             =   360
         Width           =   4155
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   2415
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   4260
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   3
            DisableNoScroll =   -1  'True
            TextRTF         =   $"Form1.frx":0054
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selector"
         Height          =   1275
         Left            =   -70740
         TabIndex        =   22
         Top             =   360
         Width           =   1635
         Begin VB.PictureBox Picture1 
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   1275
            TabIndex        =   23
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Drag This"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   1020
            Width           =   1275
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Options"
         Height          =   1515
         Left            =   -70740
         TabIndex        =   21
         Top             =   1620
         Width           =   1635
         Begin VB.CommandButton Command1 
            Caption         =   "Set Option"
            Height          =   255
            Left            =   60
            TabIndex        =   28
            Top             =   1200
            Width           =   1515
         End
         Begin VB.FileListBox File1 
            Height          =   1065
            Left            =   60
            TabIndex        =   27
            Top             =   180
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Selector"
         Height          =   1695
         Left            =   60
         TabIndex        =   18
         Top             =   360
         Width           =   1395
         Begin VB.PictureBox Picture2 
            Height          =   1155
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1095
            TabIndex        =   19
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Drag Me"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1155
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Preview"
         Height          =   1095
         Left            =   60
         TabIndex        =   16
         Top             =   2040
         Width           =   5835
         Begin VB.PictureBox Picture3 
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   5535
            TabIndex        =   17
            Top             =   240
            Width           =   5595
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Info"
         Height          =   1695
         Left            =   1500
         TabIndex        =   11
         Top             =   360
         Width           =   4395
         Begin VB.Label Label3 
            Caption         =   "Red:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4155
         End
         Begin VB.Label Label4 
            Caption         =   "Green:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label5 
            Caption         =   "Blue:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   2715
         End
         Begin VB.Label Label6 
            Caption         =   "Hex:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   4155
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Info"
         Height          =   2775
         Left            =   -74820
         TabIndex        =   1
         Top             =   360
         Width           =   5775
         Begin VB.Frame Frame8 
            Caption         =   "Selector"
            Height          =   1335
            Left            =   4140
            TabIndex        =   3
            Top             =   120
            Width           =   1515
            Begin VB.PictureBox Picture4 
               Height          =   795
               Left            =   120
               ScaleHeight     =   735
               ScaleWidth      =   1215
               TabIndex        =   4
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Drag This"
               Height          =   195
               Left            =   120
               TabIndex        =   5
               Top             =   1080
               Width           =   1275
            End
         End
         Begin VB.ListBox List1 
            Height          =   1035
            Left            =   120
            TabIndex        =   2
            Top             =   1620
            Width           =   5535
         End
         Begin VB.Label lblhwnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "hWnd:"
            Height          =   195
            Left            =   480
            TabIndex        =   10
            Top             =   540
            Width           =   3360
         End
         Begin VB.Label lblWintxt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Window text:"
            Height          =   195
            Left            =   480
            TabIndex        =   9
            Top             =   1020
            Width           =   3510
         End
         Begin VB.Label lblwininfo 
            AutoSize        =   -1  'True
            Caption         =   "Window Info"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblclass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            Height          =   195
            Left            =   480
            TabIndex        =   7
            Top             =   780
            Width           =   3300
         End
         Begin VB.Label Label8 
            Caption         =   "Parent's List"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim getinfonow As Boolean
Dim getcolornow As Boolean
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Function SpyColor()
Dim curpos As POINTAPI
Dim curcolor As RGB
    Call GetCursorPos(curpos)
    MainDC = GetDC(0)
    Picture3.BackColor = GetPixel(MainDC, curpos.X, curpos.Y)
    curcolor = GetRGB(Picture3.BackColor)
Label3.Caption = "Red: " & curcolor.Red
Label4.Caption = "Green: " & curcolor.Green
Label5.Caption = "Blue: " & curcolor.Blue
Label6.Caption = "Hex: " & Hex(Picture3.BackColor)
Call ReleaseDC(0, MainDC)
End Function

Private Sub lblParentText_Click()
End Sub

Private Sub Option1_Click()
strCodeOption = "close window"
End Sub

Private Sub Option2_Click()
strCodeOption = "enable window"
End Sub

Private Sub Option3_Click()
strCodeOption = "disable window"
End Sub

Private Sub Option4_Click()
strCodeOption = "show window"
End Sub

Private Sub Option5_Click()
strCodeOption = "hide window"
End Sub

Private Sub Option6_Click()
strCodeOption = "get text"
End Sub

Private Sub Option7_Click()
strCodeOption = "set text"
End Sub

Private Sub Command1_Click()
strCodeOption = File1.FileName
End Sub

Private Sub Form_Load()
File1.Path = "C:\Justin\WinApi Code Generator\code\"
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Long
Dim ab As POINTAPI
Call GetCursorPos(ab)
a = WindowFromPoint(ab.X, ab.Y)
RichTextBox1.Text = CreateCode(a)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
getcolornow = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If getcolornow = True Then
SpyColor
End If
End Sub


Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
getcolornow = False
End Sub


Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
getinfonow = True
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If getinfonow = True Then
Call GetWindowInformation(lblhwnd, lblclass, lblWintxt, List1)
End If
End Sub
Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
getinfonow = False
End Sub

Private Sub RichTextBox1_Change()
'SpyColor
End Sub

Sub SpyWindow()
Dim cat As POINTAPI, dog As Long, bird
Call GetCursorPos(cat)
dog = WindowFromPoint(cat.X, cat.Y)
lblhwnd.Caption = "hWnd: " & Str(dog)
Call GetWindowText(dog, bird, 254)
lblWintxt.Caption = bird
End Sub
