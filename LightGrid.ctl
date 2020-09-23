VERSION 5.00
Begin VB.UserControl LightGrid 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   ToolboxBitmap   =   "LightGrid.ctx":0000
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   6480
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   1680
      Max             =   100
      TabIndex        =   2
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox txtCell 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "Cell"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblCellFixedR 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "CellFixedR"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblCellFixedC 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "CellFixedC"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblCell 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Cell"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "LightGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EnumBorderStyle
    Kein = 0
    Fest_Einfach = 1
End Enum
Public Enum EnumBackStyle
    Transparent = 0
    Undurchsichtig = 1
End Enum
Public Enum EnumAppearance
    ug2D = 0
    ug3D = 1
End Enum
Public Enum EnumSelectionMode
    cell = 0
    horizontal = 1
    vertikal = 2
End Enum

'Standard-Eigenschaftswerte:
Const m_def_Cols = 2
Const m_def_Rows = 2
Const m_def_FixedCols = 1
Const m_def_FixedRows = 1
'Eigenschaftsvariablen:
Dim m_Cols As Integer
Dim m_Rows As Integer
Dim m_FixedCols As Integer
Dim m_FixedRows As Integer
Dim m_SelectionMode As EnumSelectionMode
'Ereignisdeklarationen:
Event Click() 'MappingInfo=lblCell(0),lblCell,0,Click
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event CellClick(Text As String, Index As Integer)  'MappingInfo=lblCell(0),lblCell,0,Click
Event DblClick() 'MappingInfo=lblCell(0),lblCell,0,DblClick
Attribute DblClick.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt und anschließend erneut drückt und wieder losläßt."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtCell,txtCell,-1,KeyDown
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtCell,txtCell,-1,KeyPress
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtCell,txtCell,-1,KeyUp
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblCell(0),lblCell,0,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblCell(0),lblCell,0,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblCell(0),lblCell,0,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Tritt auf, wenn ein beliebiger Teil eines Formulars oder Bildfeld-Steuerelements verschoben, vergrößert oder offengelegt wird."
Event HScroll() 'MappingInfo=HScroll1,HScroll1,-1,Scroll
Event VScroll() 'MappingInfo=VScroll1,VScroll1,-1,Scroll

'Variablen:
Private A As Long
Private B As Long

Private PrevCell As Integer
Private PrevStart As Integer
Private PrevEnde As Integer


Public Function TextCell(ByVal Text As String, ByVal tCol As Integer, ByVal tRow As Integer)
A = (m_Cols - m_FixedCols) * tRow

lblCell(A + tCol + 1).Caption = Text
End Function

Public Function TextCellFixedC(ByVal Text As String, ByVal Index As Integer)
    lblCellFixedC(Index).Caption = Text
End Function

Public Function TextCellFixedR(ByVal Text As String, ByVal Index As Integer)
    lblCellFixedR(Index).Caption = Text
End Function

Public Function CellBackColor(X As Integer, Y As Integer, ByVal Color As OLE_COLOR)
A = (m_Cols - m_FixedCols) * Y

lblCell(A + X + 1).BackColor = Color
End Function

Public Function CellForeColor(X As Integer, Y As Integer, ByVal Color As OLE_COLOR)
A = (m_Cols - m_FixedCols) * Y

lblCell(A + X + 1).ForeColor = Color
End Function

Public Function CellBorderStyle(X As Integer, Y As Integer, ByVal Style As EnumBorderStyle)
A = (m_Cols - m_FixedCols) * Y

lblCell(A + X + 1).BorderStyle = Style
End Function

Public Function CellBackStyle(X As Integer, Y As Integer, ByVal Style As EnumBackStyle)
A = (m_Cols - m_FixedCols) * Y

lblCell(A + X + 1).BackStyle = Style
End Function

Public Function CellAppearance(X As Integer, Y As Integer, ByVal Appearance As EnumAppearance)
A = (m_Cols - m_FixedCols) * Y

lblCell(A + X + 1).Appearance = Appearance
End Function

'Dies ist die wichtigste Sub    'This is the most important sub
Private Sub CreateGrid()
Dim X As Long
Dim Y As Long

'Entlade alle Controls  'Unload all controls
For A = 1 To lblCellFixedC.UBound
    Unload lblCellFixedC(A)
Next A
For A = 1 To lblCellFixedR.UBound
    Unload lblCellFixedR(A)
Next A
For A = 1 To lblCell.UBound
    Unload lblCell(A)
Next A


'Erstelle feste Spalten     'create fixed columns
X = 0: Y = 0
For A = 1 To m_FixedCols * (m_Rows - m_FixedRows)
    Load lblCellFixedC(A)
    
    lblCellFixedC(A).Left = X * lblCellFixedC(A).Width
    lblCellFixedC(A).Top = (Y + m_FixedRows) * lblCellFixedC(A).Height
    lblCellFixedC(A).Visible = True
    
    X = X + 1
    If X = m_FixedCols Then X = 0: Y = Y + 1
Next A

'Erstelle feste Reihen     'create fixed rows
X = 0: Y = 0
For A = 1 To m_FixedRows * (m_Cols - m_FixedCols)
    Load lblCellFixedR(A)

    lblCellFixedR(A).Left = (X + m_FixedCols) * lblCellFixedR(A).Width
    lblCellFixedR(A).Top = Y * lblCellFixedR(A).Height
    lblCellFixedR(A).Visible = True

    X = X + 1
    If X = m_Cols - m_FixedCols Then X = 0: Y = Y + 1
Next A


'erstelle Zellen    'create cells
X = 0: Y = 0
For A = 1 To (m_Cols - m_FixedCols) * (m_Rows - m_FixedRows)

    Load lblCell(A)
    lblCell(A).Left = (X + m_FixedCols) * lblCell(A).Width
    lblCell(A).Top = (Y + m_FixedRows) * lblCell(A).Height
    lblCell(A).Visible = True

    X = X + 1
    If X = m_Cols - m_FixedCols Then X = 0: Y = Y + 1
Next A

End Sub

Private Sub UserControl_Initialize()
    lblCell(0).Caption = ""
    txtCell.Text = ""
    lblCellFixedC(0).Caption = ""
    lblCellFixedR(0).Caption = ""
End Sub

Private Sub UserControl_Resize()

Dim FieldWidth As Long
Dim FieldHeight As Long

VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = 0
HScroll1.Max = 0

'Breite der Felder ausrechnen   'Calculate the width if the field
    
    'Restliche Controls hinzufügen    'Add remaining controls
    For A = 1 To m_Cols
        FieldWidth = FieldWidth + lblCell(0).Width
        'Fals Felder darüber: merken    'If squares over: notice
        If FieldWidth > UserControl.ScaleWidth Then HScroll1.Max = HScroll1.Max + 1
    Next A

'Höhe der Felder ausrechnen     'Calculate the height of the field
    
    'Restliche Controls hinzufügen    'Add remaining controls
    For A = 1 To m_Rows
        FieldHeight = FieldHeight + lblCell(0).Height
        'Fals Felder darüber: merken    'If squares over: notice
        If FieldHeight > UserControl.ScaleHeight Then VScroll1.Max = VScroll1.Max + 1
    Next A


'Scrollbar positionieren    'Position scrollbars
'HSCROLL:
HScroll1.Width = UserControl.ScaleWidth - 17
HScroll1.Height = 17
HScroll1.Left = 0
HScroll1.Top = UserControl.ScaleHeight - 17

'VSCROll:
VScroll1.Width = 17
VScroll1.Height = UserControl.ScaleHeight - 17
VScroll1.Left = UserControl.ScaleWidth - 17
VScroll1.Top = 0

'Überprüfen, ob ScrollBar benötigt wird     'Check if we need a scrollbar

'HSCROLL:
'Postition
If FieldWidth > UserControl.ScaleWidth Then
    HScroll1.Visible = True
Else
    HScroll1.Visible = False
End If

'VSCROLL:
'Postition
If FieldHeight > UserControl.ScaleHeight Then
    VScroll1.Visible = True
Else
    VScroll1.Visible = False
End If
End Sub

Private Sub HScroll1_Change()
Dim Width1 As Long
Dim Width2 As Long

B = 0

'Entferne das Textfeld  'delete the textbox
txtCell.Visible = False

'Fals keine Zellen, aus     'If there are no cells, exit
If lblCell.UBound = 0 Then Exit Sub


    'Verschiebungsgröße berechnen   'Calculate the width the labels must be shifted
    For A = 1 To HScroll1.Value
        Width1 = Width1 + lblCell(A).Width
    Next A
    


    'Jede Label verschieben     'shift each label
    For A = 1 To lblCell.UBound
        'Unsichtbar damit es schneller isr  'invisible to improve speed
        lblCell(A).Visible = False
        
        'Veschieben     'Shift
        lblCell(A).Left = (Width2 - Width1) + (lblCellFixedC(0).Width * m_FixedCols)
        
        
        
        VisibilityCheck
        
        
        
        Width2 = Width2 + lblCell(A).Width
        
        'Fals neue Zeile    'If next row
        If A - B >= m_Cols - m_FixedCols Then
            Width2 = 0
            B = A
        End If
    Next A
    
    
    
    
    Width2 = 0
    B = 0
    
    'Jede Feste Label verschieben   'shift each fixed label
    For A = 1 To lblCellFixedR.UBound
        'Unsichtbar damit es schneller isr  'invisible to improve speed
        lblCellFixedR(A).Visible = False
    
        'Veschieben     'Shift
        lblCellFixedR(A).Left = (Width2 - Width1) + (lblCellFixedC(0).Width * m_FixedCols)
                
        'Sichtbar oder unsichtbar   'Visible or invisible
        If lblCellFixedR(A).Left < lblCellFixedC(0).Width * m_FixedCols Then
            lblCellFixedR(A).Visible = False
        ElseIf lblCellFixedR(A).Left > UserControl.ScaleWidth Then
            lblCellFixedR(A).Visible = False
        Else
            lblCellFixedR(A).Visible = True
        End If
        
        Width2 = Width2 + lblCellFixedR(A).Width
        
        'Fals neue Zeile    'If next row
        If A - B >= m_Cols - m_FixedCols Then
            Width2 = 0
            B = A
        End If
    Next A
End Sub





Private Sub VScroll1_Change()
Dim Height1 As Long
Dim Height2 As Long

A = 0
B = 0

'Entferne das Textfeld  'delete the textbox
txtCell.Visible = False


'Fals keine Zellen, aus     'If there are no cells, exit
If lblCell.UBound = 0 Then Exit Sub


    'Verschiebungsgröße berechnen   'Calculate the height the labels must be shifted
    For A = 1 To VScroll1.Value
        Height1 = Height1 + lblCell(A).Height
    Next A
    
    
    
    
    'Jede Label verschieben     'shift each label
    For A = 1 To lblCell.UBound
        'Unsichtbar damit es schneller isr  'invisible to improve speed
        lblCell(A).Visible = False
    
        'Verschieben    'shift
        lblCell(A).Top = (Height2 - Height1) + (lblCellFixedC(0).Height * m_FixedRows)
        
        
        
        VisibilityCheck



        'Fals neue Zeile    'If next row
        If A - B >= m_Cols - m_FixedCols Then
            Height2 = Height2 + lblCell(A).Height
            B = A
        End If
    Next A
    
    
    
    
    
    
    B = 0
    Height2 = 0
    
    'Jede feste Label verschieben     'shift each fixed label
    For A = 1 To lblCellFixedC.UBound
        'Unsichtbar damit es schneller isr  'invisible to improve speed
        lblCellFixedC(A).Visible = False
    
        'Verschieben    'shift
        lblCellFixedC(A).Top = (Height2 - Height1) + (lblCellFixedC(0).Height * m_FixedRows)
        
        'Sichtbar oder unsichtbar   'Visible or invisible
        If lblCellFixedC(A).Top < lblCellFixedC(0).Height * m_FixedRows Then
            lblCellFixedC(A).Visible = False
        ElseIf lblCellFixedC(A).Top > UserControl.ScaleHeight - HScroll1.Height Then
            lblCellFixedC(A).Visible = False
        Else
            lblCellFixedC(A).Visible = True
        End If

        'Fals neue Zeile    'If next row
        If A - B >= m_FixedCols Then
            Height2 = Height2 + lblCellFixedC(A).Height
            B = A
        End If
    Next A

End Sub

Private Sub VisibilityCheck()
        'Sichtbar oder unsichtbar   'Visible or invisible
        If lblCell(A).Left < lblCellFixedC(0).Width * m_FixedCols Then
            lblCell(A).Visible = False
        ElseIf lblCell(A).Left > UserControl.ScaleWidth Then
            lblCell(A).Visible = False
        ElseIf lblCell(A).Top < lblCellFixedC(0).Height * m_FixedRows Then
            lblCell(A).Visible = False
        ElseIf lblCell(A).Top > UserControl.ScaleHeight - HScroll1.Height Then
            lblCell(A).Visible = False
        Else
            lblCell(A).Visible = True
        End If
End Sub


'Der Unterstrich hinter "Circle" ist erforderlich, da es
'sich um ein reserviertes Wort in VBA handelt.
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub Circle_(X As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
    UserControl.Circle (X, Y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Löscht Grafiken und Texte, die zur Laufzeit von einem Formular, Anzeige- oder Bildfeld-Steuerelement erzeugt wurden."
    UserControl.Cls
End Sub

Public Sub LineDraw(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
Attribute LineDraw.VB_Description = "Gibt Linien und Rechtecke auf einem Objekt aus."
    UserControl.Line (X1, Y1)-(X2, Y2), Color
End Sub

Public Sub RectDrawOpen(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    UserControl.Line (X1, Y1)-(X2, Y2), Color, B
End Sub

Public Sub RectDrawClose(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    UserControl.Line (X1, Y1)-(X2, Y2), Color, BF
End Sub

'Der Unterstrich hinter "Point" ist erforderlich, da es
'sich um ein reserviertes Wort in VBA handelt.
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "Gibt die RGB-Farbe des angegebenen Punkts in einem Formular oder Bildfeld als Integer vom Typ Long zurück."
    Point = UserControl.Point(X, Y)
End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Erzwingt ein vollständiges Neuzeichnen eines Objekts."
    UserControl.Refresh
End Sub

'Der Unterstrich hinter "PSet" ist erforderlich, da es
'sich um ein reserviertes Wort in VBA handelt.
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSet_(X As Single, Y As Single, Color As Long)
    UserControl.PSet (X, Y), Color
End Sub








'----------------------------------------------------------------------------
'--------------------------------- STANDART -----------------------------------
'----------------------------------------------------------------------------

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'-------------------------------------------------------------------------------

Public Property Get CurrentX() As Single
    CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'-------------------------------------------------------------------------------

Public Property Get CurrentY() As Single
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'-------------------------------------------------------------------------------

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'-------------------------------------------------------------------------------

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

'-------------------------------------------------------------------------------

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'-------------------------------------------------------------------------------

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'-------------------------------------------------------------------------------

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'-------------------------------------------------------------------------------

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'-------------------------------------------------------------------------------

Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property


'----------------------------------------------------------------------------
'--------------------------------- COLOR ------------------------------------
'----------------------------------------------------------------------------

Public Property Get cFillColor() As OLE_COLOR
    cFillColor = UserControl.FillColor
End Property

Public Property Let cFillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "cFillColor"
End Property

'-------------------------------------------------------------------------------

Public Property Get cBackColorBkg() As OLE_COLOR
    cBackColorBkg = UserControl.BackColor
End Property

Public Property Let cBackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    UserControl.BackColor() = New_BackColorBkg
    PropertyChanged "cBackColorBkg"
End Property

'-------------------------------------------------------------------------------

Public Property Get cBackColor() As OLE_COLOR
    cBackColor = lblCell(0).BackColor
End Property

Public Property Let cBackColor(ByVal New_BackColor As OLE_COLOR)
    For A = 0 To lblCell.UBound
        lblCell(A).BackColor() = New_BackColor
    Next A
    
    PropertyChanged "cBackColor"
End Property

'-------------------------------------------------------------------------------

Public Property Get cBackColorFixed() As OLE_COLOR
    cBackColorFixed = lblCellFixedC(0).BackColor
    'cBackColorFixed = lblCellFixedR(0).BackColor
End Property

Public Property Let cBackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).BackColor() = New_BackColorFixed
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(A).BackColor() = New_BackColorFixed
    Next A
    
    PropertyChanged "cBackColorFixed"
End Property

'-------------------------------------------------------------------------------

Public Property Get cBackColorSel() As OLE_COLOR
    cBackColorSel = txtCell.BackColor
End Property

Public Property Let cBackColorSel(ByVal New_BackColorSel As OLE_COLOR)
    txtCell.BackColor() = New_BackColorSel
    PropertyChanged "cBackColorSel"
End Property

'-------------------------------------------------------------------------------

Public Property Get cForeColor() As OLE_COLOR
    cForeColor = lblCell(0).ForeColor
End Property

Public Property Let cForeColor(ByVal New_ForeColor As OLE_COLOR)
    For A = 0 To lblCell.UBound
        lblCell(A).ForeColor() = New_ForeColor
    Next A
    
    PropertyChanged "cForeColor"
End Property

'-------------------------------------------------------------------------------

Public Property Get cForeColorFixed() As OLE_COLOR
    cForeColorFixed = lblCellFixedC(0).ForeColor
    'ForeColorFixed = lblCellFixedR(0).ForeColor
End Property

Public Property Let cForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).ForeColor() = New_ForeColorFixed
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(A).ForeColor() = New_ForeColorFixed
    Next A
    PropertyChanged "cForeColorFixed"
End Property

'-------------------------------------------------------------------------------

Public Property Get cForeColorSel() As OLE_COLOR
    cForeColorSel = txtCell.ForeColor
End Property

Public Property Let cForeColorSel(ByVal New_ForeColorSel As OLE_COLOR)
    txtCell.ForeColor() = New_ForeColorSel
    PropertyChanged "cForeColorSel"
End Property


'----------------------------------------------------------------------------
'---------------------------------- GRID ------------------------------------
'----------------------------------------------------------------------------

Public Property Get gCellHeight() As Integer
    gCellHeight = lblCell(0).Height
End Property

Public Property Let gCellHeight(ByVal New_CellHeight As Integer)
    For A = 0 To lblCell.UBound
        lblCell(A).Height = New_CellHeight
    Next A
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).Height = New_CellHeight
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(A).Height = New_CellHeight
    Next A
    
    txtCell.Height = New_CellHeight
    
    PropertyChanged "gCellHeight"
    
    CreateGrid
End Property

'-------------------------------------------------------------------------------

Public Property Get gCellWidth() As Integer
    gCellWidth = lblCell(0).Width
End Property

Public Property Let gCellWidth(ByVal New_CellWidth As Integer)
    For A = 0 To lblCell.UBound
        lblCell(A).Width = New_CellWidth
    Next A
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).Width = New_CellWidth
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(A).Width = New_CellWidth
    Next A

    txtCell.Width = New_CellWidth

    PropertyChanged "gCellWidth"
    
    CreateGrid
End Property

'-------------------------------------------------------------------------------


Public Property Get gCols() As Integer
    gCols = m_Cols
End Property

Public Property Let gCols(ByVal New_Cols As Integer)
    m_Cols = New_Cols
    PropertyChanged "gCols"
    
    CreateGrid
    UserControl_Resize
End Property

'-------------------------------------------------------------------------------

Public Property Get gRows() As Integer
    gRows = m_Rows
End Property

Public Property Let gRows(ByVal New_Rows As Integer)
    m_Rows = New_Rows
    PropertyChanged "gRows"
    
    CreateGrid
    UserControl_Resize
End Property

'-------------------------------------------------------------------------------

Public Property Get gFixedCols() As Integer
    gFixedCols = m_FixedCols
End Property

Public Property Let gFixedCols(ByVal New_FixedCols As Integer)
    If New_FixedCols > m_Cols Then Exit Property

    m_FixedCols = New_FixedCols
    PropertyChanged "gFixedCols"
    
    CreateGrid
End Property

'-------------------------------------------------------------------------------

Public Property Get gFixedRows() As Integer
    gFixedRows = m_FixedRows
End Property

Public Property Let gFixedRows(ByVal New_FixedRows As Integer)
    If New_FixedRows > m_Rows Then Exit Property
    
    m_FixedRows = New_FixedRows
    PropertyChanged "gFixedRows"
    
    CreateGrid
End Property

'-------------------------------------------------------------------------------

Public Property Get gSelectionMode() As EnumSelectionMode
    gSelectionMode = m_SelectionMode
End Property

Public Property Let gSelectionMode(ByVal New_SelectionMode As EnumSelectionMode)
    m_SelectionMode = New_SelectionMode
    PropertyChanged "gSelectionMode"
End Property


'----------------------------------------------------------------------------
'---------------------------------- STYLE -----------------------------------
'----------------------------------------------------------------------------

Public Property Get sAppearanceBkg() As EnumAppearance
    sAppearanceBkg = UserControl.Appearance
End Property

Public Property Let sAppearanceBkg(ByVal New_AppearanceBkg As EnumAppearance)
    UserControl.Appearance() = New_AppearanceBkg
    PropertyChanged "sAppearanceBkg"
End Property

Public Property Get sAppearance() As EnumAppearance
    sAppearance = lblCell(0).Appearance
    'txtCell.Appearance = sAppearance
End Property

Public Property Let sAppearance(ByVal New_Appearance As EnumAppearance)
    For A = 0 To lblCell.UBound
        lblCell(A).Appearance = New_Appearance
    Next A
    txtCell.Appearance = New_Appearance
    PropertyChanged "sAppearance"
End Property

Public Property Get sAppearanceFixed() As EnumAppearance
    sAppearanceFixed = lblCellFixedC(0).Appearance
    'sAppearanceFixed = txtCellFixedR.Appearance
End Property

Public Property Let sAppearanceFixed(ByVal New_AppearanceFixed As EnumAppearance)
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).Appearance() = New_AppearanceFixed
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(A).Appearance() = New_AppearanceFixed
    Next A

    PropertyChanged "sAppearanceFixed"
End Property

'-------------------------------------------------------------------------------

Public Property Get sBackStyle() As EnumBackStyle
    sBackStyle = lblCell(0).BackStyle
End Property

Public Property Let sBackStyle(ByVal New_BackStyle As EnumBackStyle)
    For A = 0 To lblCell.UBound
        lblCell(A).BackStyle = New_BackStyle
    Next A
    
    PropertyChanged "sBackStyle"
End Property

'-------------------------------------------------------------------------------

Public Property Get sBackStyleFixed() As EnumBackStyle
    sBackStyleFixed = lblCellFixedC(0).BackStyle
    'BackStyleFixed = lblCellFixedR(0).BackStyle
End Property

Public Property Let sBackStyleFixed(ByVal New_BackStyleFixed As EnumBackStyle)
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).BackStyle = New_BackStyleFixed
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(A).BackStyle = New_BackStyleFixed
    Next A

    PropertyChanged "sBackStyleFixed"
End Property

'-------------------------------------------------------------------------------

Public Property Get sBorderStyle() As EnumBorderStyle
    sBorderStyle = UserControl.BorderStyle
End Property

Public Property Let sBorderStyle(ByVal New_BorderStyle As EnumBorderStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "sBorderStyle"
End Property

'-------------------------------------------------------------------------------

Public Property Get sGridLines() As Boolean
    sGridLines = lblCell(0).BorderStyle * -1
End Property

Public Property Let sGridLines(ByVal New_GridLines As Boolean)
    For A = 0 To lblCell.UBound
        lblCell(A).BorderStyle() = New_GridLines * -1
    Next A
    
    PropertyChanged "sGridLines"
End Property

'-------------------------------------------------------------------------------

Public Property Get sGridLinesFixed() As Boolean
    sGridLinesFixed = lblCellFixedC(0).BorderStyle * -1
    'GridLinesFixed = lblCellFixedR(0).BorderStyle
End Property

Public Property Let sGridLinesFixed(ByVal New_GridLinesFixed As Boolean)
    For A = 0 To lblCellFixedC.UBound
        lblCellFixedC(A).BorderStyle() = New_GridLinesFixed * -1
    Next A
    For A = 0 To lblCellFixedR.UBound
        lblCellFixedR(0).BorderStyle() = New_GridLinesFixed * -1
    Next A
    
    PropertyChanged "sGridLinesFixed"
End Property

'-------------------------------------------------------------------------------

Public Property Get sFillStyle() As Integer
    sFillStyle = UserControl.FillStyle
End Property

Public Property Let sFillStyle(ByVal New_FillStyle As Integer)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "sFillStyle"
End Property

'-------------------------------------------------------------------------------

Public Property Get sFont() As Font
    Set sFont = lblCell(0).Font
End Property

Public Property Set sFont(ByVal New_Font As Font)

    For A = 0 To lblCell.UBound
        Set lblCell(A).Font = New_Font
    Next A
    For A = 0 To lblCellFixedC.UBound
        Set lblCellFixedC(A).Font = New_Font
    Next A
    For A = 0 To lblCellFixedR.UBound
        Set lblCellFixedR(A).Font = New_Font
    Next A
    
    Set txtCell.Font = New_Font
    
    PropertyChanged "sFont"
End Property


'----------------------------------------------------------------------------
'---------------------------------- DRAW ------------------------------------
'----------------------------------------------------------------------------

Public Property Get dDrawMode() As Integer
    dDrawMode = UserControl.DrawMode
End Property

Public Property Let dDrawMode(ByVal New_DrawMode As Integer)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "dDrawMode"
End Property

'-------------------------------------------------------------------------------

Public Property Get dDrawStyle() As Integer
    dDrawStyle = UserControl.DrawStyle
End Property

Public Property Let dDrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "dDrawStyle"
End Property

'-------------------------------------------------------------------------------

Public Property Get dDrawWidth() As Integer
    dDrawWidth = UserControl.DrawWidth
End Property

Public Property Let dDrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "dDrawWidth"
End Property






'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_Cols = m_def_Cols
    m_Rows = m_def_Rows
    m_FixedCols = m_def_FixedCols
    m_FixedRows = m_def_FixedRows
    m_HighLight = m_def_HighLight
    m_SelectionMode = m_def_SelectionMode
    Set UserControl.Font = Ambient.Font
    Set lblCell(0).Font = Ambient.Font
    Set lblCellFixedR.Font = Ambient.Font
    Set lblCellFixedC.Font = Ambient.Font
    Set txtCell.Font = Ambient.Font
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'STANDART
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    'GRID
    lblCell(0).Width = PropBag.ReadProperty("gCellWidth", 105)
    lblCellFixedR(0).Width = PropBag.ReadProperty("gCellWidth", 105)
    lblCellFixedC(0).Width = PropBag.ReadProperty("gCellWidth", 105)
    txtCell.Width = PropBag.ReadProperty("gCellWidth", 105)
    
    lblCell(0).Height = PropBag.ReadProperty("gCellHeight", 19)
    lblCellFixedR(0).Height = PropBag.ReadProperty("gCellHeight", 19)
    lblCellFixedC(0).Height = PropBag.ReadProperty("gCellHeight", 19)
    txtCell.Height = PropBag.ReadProperty("gCellHeight", 19)
    
    m_Cols = PropBag.ReadProperty("gCols", m_def_Cols)
    m_Rows = PropBag.ReadProperty("gRows", m_def_Rows)
    
    m_FixedCols = PropBag.ReadProperty("gFixedCols", m_def_FixedCols)
    m_FixedRows = PropBag.ReadProperty("gFixedRows", m_def_FixedRows)
    
    m_SelectionMode = PropBag.ReadProperty("gSelectionMode", 0)
        
    'STYLE
    UserControl.Appearance = PropBag.ReadProperty("sAppearanceBkg", 1)
    lblCell(0).Appearance = PropBag.ReadProperty("sAppearance", 1)
    txtCell.Appearance = PropBag.ReadProperty("sAppearance", 1)
    lblCellFixedC(0).Appearance = PropBag.ReadProperty("sAppearanceFixed", 1)
    lblCellFixedR(0).Appearance = PropBag.ReadProperty("sAppearanceFixed", 1)
    
    lblCell(0).BackStyle = PropBag.ReadProperty("sBackStyle", 0)
    lblCellFixedC(0).BackStyle = PropBag.ReadProperty("sBackStyleFixed", 0)
    lblCellFixedR(0).BackStyle = PropBag.ReadProperty("sBackStyleFixed", 0)
    
    UserControl.BorderStyle = PropBag.ReadProperty("sBorderStyle", 0)
    lblCell(0).BorderStyle = PropBag.ReadProperty("sGridLines", 1)
    txtCell.BorderStyle = PropBag.ReadProperty("sGridLines", 1)
    lblCellFixedC(0).BorderStyle = PropBag.ReadProperty("sGridLinesFixed", 1)
    lblCellFixedR(0).BorderStyle = PropBag.ReadProperty("sGridLinesFixed", 1)
    
    UserControl.FillStyle = PropBag.ReadProperty("sFillStyle", 1)
    
    lblCell(0).Font = PropBag.ReadProperty("sFont", Ambient.Font)
    lblCellFixedR(0).Font = PropBag.ReadProperty("sFont", Ambient.Font)
    lblCellFixedC(0).Font = PropBag.ReadProperty("sFont", Ambient.Font)
    txtCell.Font = PropBag.ReadProperty("sFont", Ambient.Font)
    
    'COLOR
    UserControl.BackColor = PropBag.ReadProperty("cBackColorBkg", &H8000000F)
    lblCell(0).BackColor = PropBag.ReadProperty("cBackColor", &H80000005)
    lblCellFixedC(0).BackColor = PropBag.ReadProperty("cBackColorFixed", &H80000005)
    lblCellFixedR(0).BackColor = PropBag.ReadProperty("cBackColorFixed", &H80000005)
    txtCell.BackColor = PropBag.ReadProperty("cBackColorSel", &H80000005)

    lblCell(0).ForeColor = PropBag.ReadProperty("cForeColor", &H80000008)
    lblCellFixedC(0).ForeColor = PropBag.ReadProperty("cForeColorFixed", &H80000008)
    lblCellFixedR(0).ForeColor = PropBag.ReadProperty("cForeColorFixed", &H80000008)
    txtCell.ForeColor = PropBag.ReadProperty("cForeColorSel", &H80000008)
    
    UserControl.FillColor = PropBag.ReadProperty("cFillColor", &H0&)
    
    'DRAW
    UserControl.DrawMode = PropBag.ReadProperty("dDrawMode", 13)
    UserControl.DrawStyle = PropBag.ReadProperty("dDrawStyle", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("dDrawWidth", 1)
    
    Call CreateGrid
    UserControl_Resize
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'STANDART
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
   
    'GRID
    Call PropBag.WriteProperty("gCellWidth", lblCell(0).Width, 105)
    Call PropBag.WriteProperty("gCellWidth", lblCellFixedR(0).Width, 105)
    Call PropBag.WriteProperty("gCellWidth", lblCellFixedC(0).Width, 105)
    Call PropBag.WriteProperty("gCellWidth", txtCell.Width, 105)
    
    Call PropBag.WriteProperty("gCellHeight", lblCell(0).Height, 19)
    Call PropBag.WriteProperty("gCellHeight", lblCellFixedR(0).Height, 19)
    Call PropBag.WriteProperty("gCellHeight", lblCellFixedC(0).Height, 19)
    Call PropBag.WriteProperty("gCellHeight", txtCell.Height, 19)
    
    Call PropBag.WriteProperty("gCols", m_Cols, m_def_Cols)
    Call PropBag.WriteProperty("gRows", m_Rows, m_def_Rows)
    
    Call PropBag.WriteProperty("gFixedCols", m_FixedCols, m_def_FixedCols)
    Call PropBag.WriteProperty("gFixedRows", m_FixedRows, m_def_FixedRows)
    
    Call PropBag.WriteProperty("gSelectionMode", m_SelectionMode, 0)
    
    'STYLE
    Call PropBag.WriteProperty("sAppearanceBkg", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("sAppearance", lblCell(0).Appearance, 1)
    Call PropBag.WriteProperty("sAppearance", txtCell.Appearance, 1)
    Call PropBag.WriteProperty("sAppearanceFixed", lblCellFixedC(0).Appearance, 1)
    Call PropBag.WriteProperty("sAppearanceFixed", lblCellFixedR(0).Appearance, 1)
    
    Call PropBag.WriteProperty("sBackStyle", lblCell(0).BackStyle, 0)
    Call PropBag.WriteProperty("sBackStyleFixed", lblCellFixedC(0).BackStyle, 0)
    Call PropBag.WriteProperty("sBackStyleFixed", lblCellFixedR(0).BackStyle, 0)
    
    Call PropBag.WriteProperty("sBorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("sGridLines", lblCell(0).BorderStyle, 1)
    Call PropBag.WriteProperty("sGridLines", txtCell.BorderStyle, 1)
    Call PropBag.WriteProperty("sGridLinesFixed", lblCellFixedC(0).BorderStyle, 1)
    Call PropBag.WriteProperty("sGridLinesFixed", lblCellFixedR(0).BorderStyle, 1)
    
    Call PropBag.WriteProperty("sFillStyle", UserControl.FillStyle, 1)
    
    Call PropBag.WriteProperty("sFont", lblCell(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("sFont", lblCellFixedR(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("sFont", lblCellFixedC(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("sFont", txtCell.Font, Ambient.Font)
    
    'COLOR
    Call PropBag.WriteProperty("cFillColor", UserControl.FillColor, &H0&)
    
    Call PropBag.WriteProperty("cBackColorBkg", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("cBackColor", lblCell(0).BackColor, &H80000005)
    Call PropBag.WriteProperty("cBackColorFixed", lblCellFixedC(0).BackColor, &H80000005)
    Call PropBag.WriteProperty("cBackColorFixed", lblCellFixedR(0).BackColor, &H80000005)
    Call PropBag.WriteProperty("cBackColorSel", txtCell.BackColor, &H80000005)
    
    Call PropBag.WriteProperty("cForeColor", lblCell(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("cForeColorFixed", lblCellFixedC(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("cForeColorFixed", lblCellFixedR(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("cForeColorSel", txtCell.ForeColor, &H80000008)
        
    'DRAW
    Call PropBag.WriteProperty("dDrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("dDrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("dDrawWidth", UserControl.DrawWidth, 1)

End Sub

Private Sub lblCellFixedC_Click(Index As Integer)
    txtCell.Visible = False
    RaiseEvent Click
End Sub

Private Sub lblCellFixedR_Click(Index As Integer)
    txtCell.Visible = False
    RaiseEvent Click
End Sub

Private Sub lblCellFixedR_DblClick(Index As Integer)
    txtCell.Visible = False
    RaiseEvent DblClick
End Sub

Private Sub lblCellFixedC_DblClick(Index As Integer)
    txtCell.Visible = False
    RaiseEvent DblClick
End Sub

Private Sub lblCell_DblClick(Index As Integer)
    txtCell.Visible = False
    RaiseEvent DblClick
End Sub

Private Sub txtCell_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_DblClick()
    txtCell.Visible = False
    RaiseEvent DblClick
End Sub

Private Sub lblCell_Click(Index As Integer)
Dim Start As Integer    'Start
Dim Ende As Integer     'End

Select Case m_SelectionMode
Case 0 'cell
    txtCell.Visible = True
    lblCell(PrevCell) = txtCell.Text
    
    txtCell.Move lblCell(Index).Left, lblCell(Index).Top, lblCell(Index).Width, lblCell(Index).Height
    txtCell.Text = lblCell(Index).Caption
    
    txtCell.SetFocus
    
    PrevCell = Index

Case 1 'horizontal
    txtCell.Visible = False

    'Alles löschen      'Delete all
    For A = PrevStart To PrevEnde
        lblCell(A).BackStyle = lblCell(0).BackStyle
        lblCell(A).BackColor = lblCell(0).BackColor
        lblCell(A).ForeColor = lblCell(0).ForeColor
    Next A
    
    'Erste Zelle suchen     'Search the first cell
    B = Index
    For A = Index To 0 Step -1
       
        If B Mod (m_Cols - m_FixedCols) = 0 Then
            B = B + 1
            Exit For
        End If
        B = B - 1
    Next A
    Start = B
    
    'Letzte Zelle suchen    'Search the last cell
    B = Index
    For A = Index To lblCell.UBound
        If B Mod (m_Cols - m_FixedCols) = 0 Then Exit For
        B = B + 1
    Next A
    Ende = B
    
    'Markieren  'mark
    For A = Start To Ende
        lblCell(A).BackStyle = 1
        lblCell(A).BackColor = txtCell.BackColor
        lblCell(A).ForeColor = txtCell.ForeColor
    Next A
    
    PrevStart = Start
    PrevEnde = Ende
Case 2 'vertikal
    
    
End Select

RaiseEvent Click
RaiseEvent CellClick(lblCell(Index).Caption, Index)
End Sub

Private Sub txtCell_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    txtCell.Visible = False
    RaiseEvent Click
End Sub


Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub


Private Sub txtCell_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCell_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtCell_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub lblCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblCell_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub HScroll1_Scroll()
    HScroll1_Change
    RaiseEvent HScroll
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
    RaiseEvent VScroll
End Sub
