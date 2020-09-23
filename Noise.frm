VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Pixelize (Made by Johannes.B    Email: JB_Rulez_54@hotmail.com)"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   6000
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   9
      Top             =   1920
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Animation (cool)"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   3
      Left            =   960
      Max             =   25
      Min             =   2
      TabIndex        =   5
      Top             =   600
      Value           =   3
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save picture..."
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open picture..."
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLS"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pixelize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4260
      Left            =   120
      Picture         =   "Noise.frx":0000
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   0
      Top             =   960
      Width           =   5670
      Begin MSComDlg.CommonDialog CM 
         Left            =   4560
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label2 
      Caption         =   "3"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Pixel size:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BC
Dim A, B As Integer
Dim JB As Integer

Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()
Command1.Caption = "PLEASE WAIT..."
Set Picture1.Picture = image1.Picture
A = 0
B = 0

Do

'Get color
BC = Picture1.Point(A + 1, B + 1)
'Draw
image1.Line (A, B)-(A + HScroll1.Value, B + HScroll1.Value), BC, BF
'Incrace left
A = A + HScroll1.Value

If A > image1.ScaleWidth Then
'Incrase top
A = 0
B = B + HScroll1.Value
image1.Refresh
End If

Loop Until B > image1.ScaleHeight
image1.Refresh
Command1.Caption = "Pixelize"
End Sub

Private Sub Command2_Click()
image1.Cls
End Sub


Private Sub Command3_Click()
CM.CancelError = True
On Error GoTo err

CM.Filter = "All supported formats ()|*.BMP;*.JPG;*.GIF;*.WMF;*.EMF;*.DIB;*.ICO;*.CUR|Bitmap (*.BMP)|*.Bmp|Bitmap (*.DIB)|*.Dib|Gif Images (*.GIF)|*.Gif|Jpeg Images (*.JPG)|*.Jpg|Metafiles (*.WMF)|*.Wmf|Metafiles (*.EMF)|*.Emf|Icons (*.ICO)|*.Ico|Icons (*.CUR)|*.Cur"

CM.ShowOpen

image1.Picture = LoadPicture(CM.FileName)

Exit Sub
err:
Exit Sub
End Sub


Private Sub Command4_Click()
CM.CancelError = True
On Error GoTo err

CM.Filter = "Bitmap (*.BMP)|*.bmp"

CM.ShowSave

SavePicture image1.Image, CM.FileName

Exit Sub
err:
Exit Sub
End Sub


Private Sub Command5_Click()

JB = HScroll1.Value
HScroll1.Value = "3"
HScroll1.Max = 500
Do
HScroll1.Value = HScroll1.Value + 1
Command1.Value = True
Loop Until HScroll1.Value = 500
HScroll1.Value = JB
HScroll1.Max = 20
Command2.Value = True
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "PLEASE VOTE!"
End Sub


Private Sub HScroll1_Change()
Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Label2.Caption = HScroll1.Value
End Sub


Private Sub Picture2_Click()

End Sub


