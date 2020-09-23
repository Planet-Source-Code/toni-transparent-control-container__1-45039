VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin Project1.TranspContainer UserControl11 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   16748
      BackColor       =   16777215
      Picture         =   "Form1.frx":0000
      TransparentColor=   16777215
      Begin VB.VScrollBar VScroll1 
         Height          =   2415
         Left            =   7320
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Height          =   4965
         Left            =   0
         TabIndex        =   8
         Top             =   4320
         Width           =   1815
      End
      Begin VB.DirListBox Dir1 
         Height          =   3690
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   7680
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   9600
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Transparent ?"
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   7680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3720
         Width           =   3615
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0FF&
         Height          =   1575
         Left            =   7680
         ScaleHeight     =   1515
         ScaleWidth      =   3555
         TabIndex        =   1
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   1695
         Left            =   7680
         TabIndex        =   11
         Top             =   5400
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If UserControl11.Transparent = True Then
        UserControl11.Transparent = False
    Else
        UserControl11.Transparent = True
    End If
End Sub

Private Sub Form_Resize()
    UserControl11.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub UserControl11_Resize()
    Picture1.Move UserControl11.Width / 2 - Picture1.Width / 2 _
            , UserControl11.Height / 2 - Picture1.Height / 2
End Sub
