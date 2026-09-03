VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H0038281B&
   BorderStyle     =   0  'None
   Caption         =   "czForm Demo"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   2  'CenterScreen
   Begin czFormDemo.czControl czTitleBar1 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   820
      ControlType     =   0
      Caption         =   "czForm Demo"
      BackColor       =   "#B0C8E2"
      CornerRadius    =   0
      FontSize        =   11
      FontBold        =   -1  'True
   End
   Begin czFormDemo.czControl czPanel1 
      Height          =   1350
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   2381
      ControlType     =   3
      BackColor       =   "#243447"
      CornerRadius    =   6
      ShowHeader      =   -1  'True
      HeaderText      =   "Data Traffic"
   End
   Begin czFormDemo.czControl czLabel1 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1170
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   529
      ControlType     =   2
      Caption         =   "Downloaded          3.4 MB"
      BackColor       =   "#243447"
      FontSize        =   9
   End
   Begin czFormDemo.czControl czLabel2 
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1530
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   529
      ControlType     =   2
      Caption         =   "Uploaded              6.2 MB"
      BackColor       =   "#243447"
      FontSize        =   9
   End
   Begin czFormDemo.czControl czLabel3 
      Height          =   420
      Left            =   180
      TabIndex        =   4
      Top             =   2400
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   741
      ControlType     =   2
      Caption         =   "Kill Switch"
      FontBold        =   -1  'True
   End
   Begin czFormDemo.czControl czToggle1 
      Height          =   420
      Left            =   5220
      TabIndex        =   5
      Top             =   2400
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ControlType     =   5
      Checked         =   -1  'True
   End
   Begin czFormDemo.czControl czLabel4 
      Height          =   300
      Left            =   180
      TabIndex        =   6
      Top             =   3060
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   529
      ControlType     =   2
      Caption         =   "Bandwidth Usage"
      ForeColor       =   "#8899AA"
      FontSize        =   8
   End
   Begin czFormDemo.czControl czProgress1 
      Height          =   330
      Left            =   180
      TabIndex        =   7
      Top             =   3510
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   582
      ControlType     =   6
      CornerRadius    =   11
      FontSize        =   8
      Progress        =   65
   End
   Begin czFormDemo.czControl czTextBox1 
      Height          =   510
      Left            =   180
      TabIndex        =   8
      Top             =   4050
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   900
      ControlType     =   4
      BackColor       =   "#243447"
      Text            =   "Search location..."
      PlaceholderText =   "Search location..."
   End
   Begin czFormDemo.czControl czButton1 
      Height          =   630
      Left            =   1500
      TabIndex        =   9
      Top             =   4800
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   1111
      Caption         =   "Connect"
      CornerRadius    =   22
      FontSize        =   12
      FontBold        =   -1  'True
   End
   Begin czFormDemo.czControl czLabel5 
      Height          =   330
      Left            =   180
      TabIndex        =   10
      Top             =   5700
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   582
      ControlType     =   2
      Caption         =   "Location: United States, Honolulu"
      ForeColor       =   "#8899AA"
      FontSize        =   9
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    If ScaleWidth > 0 Then
        czTitleBar1.Move 0, 0, ScaleWidth
    End If
    On Error GoTo 0
End Sub

Private Sub czButton1_Click()
    MsgBox "Connect button clicked!", vbInformation, "czForm Demo"
End Sub

Private Sub czToggle1_ValueChanged(ByVal NewValue As Boolean)
    If NewValue Then
        czLabel3.Caption = "Kill Switch (Enabled)"
        czLabel3.ForeColor = &HFFFFFF
    Else
        czLabel3.Caption = "Kill Switch (Disabled)"
        czLabel3.ForeColor = &H6B6BFF
    End If
End Sub

Private Sub czTextBox1_TextChanged(ByVal NewText As String)
    Debug.Print "Search: " & NewText
End Sub

Private Sub czTitleBar1_IconClick()
    MsgBox "czControl Demo" & vbCrLf & _
           "Modern UI Framework for VB6" & vbCrLf & _
           "Powered by GDI+", vbInformation, "About"
End Sub
