VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "LightGrid"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   2  'Bildschirmmitte
   Begin Projekt1.LightGrid LightGrid1 
      Height          =   7200
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   12700
      AutoRedraw      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":030A
      gCols           =   10
      gRows           =   50
      sAppearance     =   0
      sAppearance     =   0
      sBackStyleFixed =   1
      sBackStyleFixed =   1
      sBorderStyle    =   1
      BeginProperty sFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty sFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty sFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty sFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFillColor      =   12632319
      cBackColorFixed =   4194304
      cBackColorFixed =   4194304
      cBackColorSel   =   16711680
      cForeColor      =   255
      cForeColorFixed =   16777215
      cForeColorFixed =   16777215
      cForeColorSel   =   16776960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim X As Integer
Dim Y As Integer
Dim StopIt As Boolean

StopIt = True

For X = 1 To LightGrid1.gFixedRows * (LightGrid1.gCols - LightGrid1.gFixedCols)
    LightGrid1.TextCellFixedR X, X
Next X
For Y = 1 To LightGrid1.gFixedCols * (LightGrid1.gRows - LightGrid1.gFixedRows)
    LightGrid1.TextCellFixedC Y, Y
Next Y

Y = 0
X = 0

Do Until StopIt = False

LightGrid1.TextCell "cell " & X & "," & Y, X, Y

If X Mod 2 = 0 Then LightGrid1.CellAppearance X, Y, ug3D

If X Mod 2 = 0 And Y Mod 2 = 0 Then
    LightGrid1.CellBorderStyle X, Y, Kein
    LightGrid1.CellForeColor X, Y, vbWhite
End If


X = X + 1

If X >= LightGrid1.gCols - LightGrid1.gFixedCols Then Y = Y + 1: X = 0

If Y >= LightGrid1.gRows - LightGrid1.gFixedRows Then StopIt = False

Loop
End Sub

Private Sub LightGrid1_Click()

End Sub

