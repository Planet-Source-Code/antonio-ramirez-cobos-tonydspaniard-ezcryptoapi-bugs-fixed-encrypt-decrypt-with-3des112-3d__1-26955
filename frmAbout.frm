VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "EzCryptoApi Control by Antonio Ramirez Cobos [AKA: TonyDSpaniard :~)]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1410
      TabIndex        =   0
      Top             =   1590
      Width           =   5340
   End
   Begin VB.Image Image2 
      Height          =   1515
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Image2.Top = 0
    Image2.Left = 0
    Label1.Left = 0
    Label1.Width = Image2.Width - 50
    Label1.Top = Image2.Height + 50
    Me.Width = Image2.Width
    Me.Height = Label1.Top + Label1.Height + 50
End Sub


Private Sub Image2_Click()
    Unload Me
End Sub

