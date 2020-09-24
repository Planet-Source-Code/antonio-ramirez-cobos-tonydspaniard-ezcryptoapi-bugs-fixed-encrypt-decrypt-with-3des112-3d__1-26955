VERSION 5.00
Object = "*\AEzCryptoApi.vbp"
Begin VB.Form frmTest 
   Caption         =   "TestForm"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin CryptoApi.EzCryptoApi EzCryptoApi1 
      Left            =   5880
      Top             =   1320
      _ExtentX        =   1640
      _ExtentY        =   1905
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Decrypt to Destination File"
      Height          =   375
      Left            =   1695
      TabIndex        =   45
      Top             =   6135
      Width           =   3090
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Encrypt to Destination File"
      Height          =   375
      Left            =   1695
      TabIndex        =   44
      Top             =   5730
      Width           =   3090
   End
   Begin VB.TextBox Text5 
      Height          =   345
      Left            =   1080
      TabIndex        =   42
      Text            =   "C:\Hash.txt"
      Top             =   1320
      Width           =   3720
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Exit"
      Height          =   300
      Left            =   6945
      TabIndex        =   36
      Top             =   6630
      Width           =   1530
   End
   Begin VB.Frame Frame4 
      Caption         =   "Encryption/Decryption Speed"
      Height          =   2025
      Left            =   5100
      TabIndex        =   29
      Top             =   4455
      Width           =   3375
      Begin VB.OptionButton Option4 
         Caption         =   "100KB"
         Height          =   240
         Index           =   10
         Left            =   2055
         TabIndex        =   41
         Top             =   1665
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "80KB"
         Height          =   240
         Index           =   9
         Left            =   2055
         TabIndex        =   40
         Top             =   1380
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "60KB"
         Height          =   240
         Index           =   8
         Left            =   2055
         TabIndex        =   39
         Top             =   1095
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "50KB"
         Height          =   240
         Index           =   7
         Left            =   2055
         TabIndex        =   38
         Top             =   795
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "40KB"
         Height          =   240
         Index           =   6
         Left            =   2055
         TabIndex        =   37
         Top             =   510
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "20KB"
         Height          =   240
         Index           =   5
         Left            =   2055
         TabIndex        =   35
         Top             =   225
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "10KB"
         Height          =   240
         Index           =   4
         Left            =   225
         TabIndex        =   34
         Top             =   1440
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "5KB"
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   33
         Top             =   1185
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "3KB"
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   32
         Top             =   870
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "2KB"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   31
         Top             =   570
         Width           =   1000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "1KB"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   30
         Top             =   270
         Width           =   1000
      End
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   1095
      TabIndex        =   26
      Text            =   "Text4"
      Top             =   105
      Width           =   3750
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Decrypt Only Source File"
      Height          =   375
      Left            =   1695
      TabIndex        =   25
      Top             =   5325
      Width           =   3075
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Encrypt Only Source File"
      Height          =   375
      Left            =   1710
      TabIndex        =   24
      Top             =   4920
      Width           =   3075
   End
   Begin VB.Frame Frame3 
      Caption         =   "Encryption Algorithms"
      Height          =   1335
      Left            =   5100
      TabIndex        =   18
      Top             =   60
      Width           =   3420
      Begin VB.OptionButton Option3 
         Caption         =   "Triple DES 112"
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   23
         Top             =   990
         Width           =   1515
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Triple DES"
         Height          =   210
         Index           =   3
         Left            =   1860
         TabIndex        =   22
         Top             =   652
         Width           =   1515
      End
      Begin VB.OptionButton Option3 
         Caption         =   "RC4"
         Height          =   210
         Index           =   1
         Left            =   1860
         TabIndex        =   21
         Top             =   315
         Width           =   1245
      End
      Begin VB.OptionButton Option3 
         Caption         =   "DES"
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   20
         Top             =   652
         Width           =   1515
      End
      Begin VB.OptionButton Option3 
         Caption         =   "RC2"
         Height          =   210
         Index           =   0
         Left            =   330
         TabIndex        =   19
         Top             =   315
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Decrypt Text Box"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   6135
      Width           =   1560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Encrypt Text Box"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5730
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hash File"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5325
      Width           =   1560
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   1080
      TabIndex        =   14
      Text            =   "C:\Hash.txt"
      Top             =   720
      Width           =   3720
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hash Value Format"
      Height          =   1380
      Left            =   5100
      TabIndex        =   8
      Top             =   1575
      Width           =   3420
      Begin VB.OptionButton Option2 
         Caption         =   "Ascii String"
         Height          =   315
         Index           =   2
         Left            =   210
         TabIndex        =   11
         Top             =   960
         Width           =   1845
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Numeric String"
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   615
         Width           =   1845
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Hexadecimal String"
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   270
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hashing Algorithms"
      Height          =   1245
      Left            =   5100
      TabIndex        =   3
      Top             =   3135
      Width           =   3420
      Begin VB.OptionButton Option1 
         Caption         =   "MD2"
         Height          =   360
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MD4"
         Height          =   360
         Index           =   1
         Left            =   2025
         TabIndex        =   6
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MD5"
         Height          =   360
         Index           =   2
         Left            =   165
         TabIndex        =   5
         Top             =   750
         Width           =   870
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SHA"
         Height          =   360
         Index           =   3
         Left            =   2025
         TabIndex        =   4
         Top             =   750
         Width           =   870
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1290
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hash Text Box Data"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   900
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2100
      Width           =   4680
   End
   Begin VB.Label Label5 
      Caption         =   "Destination file to encrypt/decrypt/hash to [different than source]:"
      Height          =   270
      Left            =   120
      TabIndex        =   43
      Top             =   1080
      Width           =   4725
   End
   Begin VB.Label Label4 
      Caption         =   "Source file to encrypt/decrypt/hash:"
      Height          =   270
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   210
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Hash/Encryption/Decryption Value:"
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   3450
   End
   Begin VB.Label Label1 
      Caption         =   "Type Data To Hash or Encrypt:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3390
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'------------------------------------------------------'
' Project: EzCryptoAPI v1.0.7
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Module: frmTest
'
' Description: Test application for EzCryptoAPI ActiveX Control
'              I do not comment any of the code within this
'              application because it is the same explained
'              within the help file included in the ZIP file. The
'              only methods not-documented on the help file are
'              EncryptToDestFile and DecryptToDestFile but these
'              are well documented within the control's code.
'                               THIS IS IMPORTANT
'              Remember to register rsaenh.dll on your system's registry
'              using Regsvr32.dll.
'
'              From the Author:
'              'cause I consider myself in a continuous learning
'              path with no end on programming, please, if you
'              can improve this program
'              contact me at: *TONYDSPANIARD@HOTMAIL.COM*
'
'              I would be pleased to hear from your opinions,
'              suggestions, and/or recommendations. Also, if you
'              know something I don't know and wish to share it
'              with me, here you'll have your techy pal from Spain
'              that will do exactly the same towards you. If I can
'              help you in any way, just ask.
'
'              INTELLECTUAL COPYRIGHT STUFF [Is up to you anyway]
'              --------------------------------------------------
'              This code is copyright 2001 Antonio Ramirez Cobos
'              This code may be reused and modified for non-commercial
'              purposes only as long as credit is given to the author
'              in the programmes about box and it's documentation.
'              If you use this code, please email me at:
'              TonyDSpaniard@hotmail.com and let me know what you think
'              and what you are doing with it.
'
'              PS: Don't forget to vote for me buddy programmer!
'------------------------------------------------------'
'******************************************************'
Option Explicit
Dim m_hashFormat As Long
Dim m_sData As String
Private Sub Command1_Click()
On Error GoTo errHandler
With EzCryptoApi1
  .CreateHash
  If .IsHashReady Then
    .HashDigestData Text1.Text
    If .IsHashDataReady Then
        Text2.Text = .GetDigestedData(m_hashFormat)
    End If
    .DestroyHash
  End If
End With
Exit Sub
errHandler:
    MsgBox "An error has occurred: " & vbCrLf & _
        Err.Number & vbCrLf & Err.Description, , "Error"
End Sub

Private Sub Command2_Click()
On Error GoTo errHandler
Dim t1 As Variant
With EzCryptoApi1
        .CreateHash
        t1 = Timer
        Text1.Text = "Processing File..."
        DoEvents
        .HashDigestFile Trim(Text3.Text)
        While Not .IsHashDataReady
        Wend
        Text1.Text = "Done..." & vbCrLf & _
        "Elapsed Time: " & (Timer - t1) & " seconds"
        Text2.Text = .GetDigestedData(m_hashFormat)
    
    .DestroyHash
End With
Exit Sub
errHandler:
    Text1.Text = ""
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Command3_Click()
On Error GoTo errHandler
Dim sData As String, fLen As Long
Dim iCount As Integer, fDat() As Byte
sData = Text1.Text
With EzCryptoApi1
    sData = .EncryptData(sData)
End With
Text2.Text = ""

Text2.Text = Convert2Hex(sData)

Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Command4_Click()
On Error GoTo errHandler
Dim sData As String
Dim iCount As Integer
sData = convert2Ascii(Trim(Text2.Text))
With EzCryptoApi1
    sData = .DecryptData(sData)
End With
Text2.Text = ""
Text2.Text = sData
Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Command5_Click()
  On Error GoTo errHandler
  Dim t1 As Variant
  t1 = Timer
  If Dir(Text3.Text) = "" Then
    MsgBox "Source file does not exists!", vbExclamation, App.Title
    Exit Sub
  End If
  EzCryptoApi1.Password = Text4.Text
  EzCryptoApi1.EncryptFile Text3.Text, 6
  Text1.Text = "Encryption Elapsed time: " & CStr(Timer - t1)
  Exit Sub
errHandler:
    Text2.Text = "ERROR"
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Command6_Click()
On Error GoTo errHandler
Dim t1 As Variant
t1 = Timer
    If Dir(Text3.Text) = "" Then
        MsgBox "Source file does not exists!", vbExclamation, App.Title
        Exit Sub
    End If
  EzCryptoApi1.Password = Text4.Text
  EzCryptoApi1.DecryptFile Text3.Text, 6
  Text1.Text = "Decryption Elapsed time: " & CStr(Timer - t1)
    Exit Sub
errHandler:
    Text2.Text = "ERROR"
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Command7_Click()
    Unload Me
End Sub

Private Sub Command8_Click()
On Error GoTo errHandler
Dim t1 As Variant
t1 = Timer
  EzCryptoApi1.Password = Text4
  If Dir(Text3.Text) = "" Then
    MsgBox "Source file does not exists!", vbExclamation, App.Title
    Exit Sub
  End If
  If Dir(Text5.Text) <> "" Then
        If Not MsgBox("Destination file already exists. Overwrite?", vbYesNo, App.Title) = vbYes Then
            Exit Sub
        End If
  End If
  EzCryptoApi1.EncryptToDestFile Text3.Text, Text5.Text, 23
  Text1.Text = "Encryption Elapsed time: " & CStr(Timer - t1)
  Exit Sub
errHandler:
    Text2.Text = "ERROR"
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Command9_Click()
On Error GoTo errHandler
Dim t1 As Variant
t1 = Timer
  EzCryptoApi1.Password = Text4
  If Dir(Text3.Text) = "" Then
    MsgBox "Source file does not exists!", vbExclamation, App.Title
    Exit Sub
  End If
  If Dir(Text5.Text) <> "" Then
        If Not MsgBox("Destination file already exists. Overwrite?", vbYesNo, App.Title) = vbYes Then
            Exit Sub
        End If
  End If
  EzCryptoApi1.DecryptToDestFile Text3.Text, Text5.Text, 23
  Text1.Text = "Decryption Elapsed time: " & CStr(Timer - t1)
  Exit Sub
errHandler:
    Text2.Text = "ERROR"
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub EzCryptoApi1_DecryptionDataStart()
    Text2.Text = "Decrypting Data..."
End Sub

Private Sub EzCryptoApi1_DecryptionFileComplete()
    Me.Caption = "EzCryptoAPI ActiveX Control Demo"
    Text2.Text = "Decryption File complete!"
End Sub

Private Sub EzCryptoApi1_DecryptionDataComplete()
    Text2.Text = "Decryption Data complete"
End Sub

Private Sub EzCryptoApi1_DecryptionFileStart()
  Text2.Text = "Decrypting File..."
End Sub

Private Sub EzCryptoApi1_DecryptionFileStatus(ByVal lBytesProcessed As Long, ByVal lTotalBytes As Long)
    Me.Caption = "Decryption File processed: " & Fix((lBytesProcessed / lTotalBytes) * 100) & "%"
End Sub

Private Sub EzCryptoApi1_EncryptionDataComplete()
    Text2.Text = "Encryption data complete"
End Sub

Private Sub EzCryptoApi1_EncryptionFileComplete()
     Me.Caption = "EzCryptoAPI ActiveX Control Demo"
    Text2.Text = "Encryption File complete!"
End Sub

Private Sub EzCryptoApi1_EncryptionFileStatus(ByVal lBytesProcessed As Long, ByVal lTotalBytes As Long)
    Me.Caption = "Encryption File processed: " & Fix((lBytesProcessed / lTotalBytes) * 100) & "%"

End Sub

Private Sub EzCryptoApi1_HashDataComplete()
    Text2.Text = "Hash Data complete"
End Sub

Private Sub EzCryptoApi1_HashDataStart()
    Text2.Text = "Hashing Data..."
End Sub

Private Sub EzCryptoApi1_HashFileComplete()
    Me.Caption = "EzCryptoAPI ActiveX Control Demo"
    Text2.Text = "Hashing File complete!"
End Sub

Private Sub EzCryptoApi1_HashFileStart()
    Text2.Text = "Digesting File...[Burp!]"
End Sub

Private Sub EzCryptoApi1_HashFileStatus(ByVal lBytesProcessed As Long, ByVal lTotalBytes As Long)
    Me.Caption = "File Digestion processed: " & Fix((lBytesProcessed / lTotalBytes) * 100) & "%"

End Sub

Private Sub EzCryptoApi1_EncryptionDataStart()
  Text2.Text = "Encrypting Data..."
End Sub



Private Sub EzCryptoApi1_EncryptionFileStart()
  Text2.Text = "Encrypting File..."
End Sub

Private Sub Form_Load()
Me.Caption = "EzCryptoAPI ActiveX Control Demo"
Option1(EzCryptoApi1.HashAlgorithm).Value = True
Option2(m_hashFormat).Value = True
Option3(EzCryptoApi1.EncryptionAlgorithm).Value = True
Option4(EzCryptoApi1.Speed).Value = True
'MsgBox EzCryptoApi1.Provider
EzCryptoApi1.About
End Sub

Private Sub Form_Unload(Cancel As Integer)
With EzCryptoApi1
     If .IsHashReady Then .DestroyHash
End With
End Sub

Private Sub Option1_Click(Index As Integer)
 EzCryptoApi1.HashAlgorithm = Index
End Sub


Private Sub Option2_Click(Index As Integer)
 m_hashFormat = Index
End Sub

Private Sub Option3_Click(Index As Integer)
 EzCryptoApi1.EncryptionAlgorithm = Index
End Sub
Private Function Convert2Hex(ByVal sAsciiData As String) As String
Dim lDataLen As Long, icounter As Long
Dim sHexData As String, sReturnData As String
lDataLen = Len(sAsciiData)
For icounter = 1 To lDataLen
    sHexData = Hex(Asc(Mid$(sAsciiData, icounter, 1)))
    If Len(sHexData) < 2 Then sHexData = "0" & sHexData
    sReturnData = sReturnData & sHexData
    sHexData = ""
Next
Convert2Hex = sReturnData
End Function
Private Function convert2Ascii(ByVal sHexData As String) As String
Dim lDataLen As Long, icounter As Long
Dim sAsciiData As String, sReturnData As String
lDataLen = Len(sHexData)
For icounter = 1 To lDataLen Step 2
    sAsciiData = Chr$(CLng("&H" & (Mid$(sHexData, icounter, 2))))
    sReturnData = sReturnData & sAsciiData
    sAsciiData = ""
Next
convert2Ascii = sReturnData
End Function


Private Sub Option4_Click(Index As Integer)
  EzCryptoApi1.Speed = Index
End Sub
