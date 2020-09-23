VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      BackColor       =   &H00008000&
      Caption         =   "Frame1"
      Height          =   2955
      Left            =   510
      TabIndex        =   0
      Top             =   480
      Width           =   6465
      Begin VB.PictureBox Pic 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         DragIcon        =   "Form1.frx":0000
         DragMode        =   1  'Automatic
         Height          =   465
         Index           =   0
         Left            =   1680
         MousePointer    =   9  'Size W E
         ScaleHeight     =   465
         ScaleWidth      =   1785
         TabIndex        =   2
         Top             =   1500
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid msf 
         Height          =   2955
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   5212
         _Version        =   65541
         Rows            =   20
         Cols            =   20
         BackColorBkg    =   14737632
         GridColor       =   16761024
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   $"Form1.frx":0442
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Enter key activates/deactivates bars. (+) and (-) keys increases/decreases bars."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   3570
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Column As Integer
Public MyWidth As Integer
Public Margin As Integer
Public Amount As Integer

Private Sub Form_Load()
Dim x As Integer, Largo As Double
ReleaseMouse
Column = 1
MyWidth = 80
Margin = 200
Amount = 12
Amount = msf.ColWidth(Column) / Amount
Largo = msf.ColWidth(Column)
Pic(0).Height = msf.RowHeight(0) - MyWidth
msf.Col = 0
For x = 1 To msf.Rows - 1
    msf.TextMatrix(x, 0) = CStr(x)
    msf.TextMatrix(x, 2) = "Unlocked"
    Load Pic(Pic.Count)
    Set Pic(Pic.Count - 1).Container = fra
    Pic(Pic.Count - 1).Width = Largo / 2
    Pic(Pic.Count - 1).ZOrder 0
    Pic(Pic.Count - 1).Visible = True
    Pic(Pic.Count - 1).Tag = CStr(x)
    Pic(Pic.Count - 1).ToolTipText = "Valor actual de " + Pic(Pic.Count - 1).Tag + ": " + CStr(Pic(Pic.Count - 1).Width)
    MyRefresh Pic(Pic.Count - 1)
Next
msf.Col = 0
For x = 1 To msf.Rows - 1
    msf.Row = x
    msf.CellForeColor = RGB(0, 64, 64)
    msf.CellFontBold = True
Next
msf.Row = 1
Refill
Pic(0).Visible = False
End Sub

Private Sub Refill()
Dim x As Integer
For x = 1 To Pic.Count - 1
    Pic(x).Left = msf.Left + msf.ColWidth(0) + 25
    Pic(x).Top = msf.Top + x * msf.RowHeight(0) - ((msf.TopRow - 1) * msf.RowHeight(0)) + MyWidth / 2
    Pic(x).Visible = IIf(Pic(x).Top < msf.RowHeight(0), False, True)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
ReleaseMouse
End Sub

Private Sub msf_KeyPress(KeyAscii As Integer)
Dim Actual As Integer, Index As Integer, Largo As Double
Largo = msf.ColWidth(Column)
Index = msf.Row
Select Case KeyAscii
Case 13
    Actual = msf.Row
    If msf.Col = Column Then
        Pic(Actual).Enabled = IIf(Pic(Actual).Enabled, False, True)
        msf.TextMatrix(Actual, 2) = IIf(Pic(Actual).Enabled, "Unlocked", "***Locked***")
    End If
Case 43
    If Pic(Actual).Enabled And Pic(Index).Width + Amount <= msf.ColWidth(Column) Then
        Pic(Index).Width = Pic(Index).Width + Amount
        MyRefresh Pic(Index)
    End If
Case 45
    If Pic(Actual).Enabled And Pic(Index).Width - Amount >= 0 Then
        Pic(Index).Width = Pic(Index).Width - Amount
        MyRefresh Pic(Index)
    End If
End Select
End Sub

Private Sub MyRefresh(Pic As PictureBox)
Dim Largo As Double
Largo = msf.ColWidth(Column)
Pic.BackColor = RGB(255 * (Pic.Width) / Largo, 255 - 255 * (Pic.Width) / Largo, 0)
Pic.ToolTipText = "Actual value of column " + Pic.Tag + ": " + Format(Pic.Width / msf.ColWidth(Column) * 100, "##0") + "%"
Me.Caption = "Row : " + Pic.Tag
msf.Refresh
End Sub

Private Sub msf_RowColChange()
Refill
Me.Caption = "Row : " + CStr(msf.Row)
End Sub

Private Sub msf_Scroll()
msf.Row = msf.TopRow
End Sub

Private Sub Pic_Click(Index As Integer)
MsgBox "Index:" + Pic(Index).Tag + " Size:" + CStr(Pic(Index).Width), vbOKOnly, "Click!"
End Sub

Private Sub Pic_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
ReleaseMouse
End Sub

Private Sub msf_DragDrop(Source As Control, x As Single, y As Single)
ReleaseMouse
End Sub

Private Sub Pic_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
Dim Largo As Double
LimitMouse Index
Largo = msf.ColWidth(Column)
msf.Row = Val(Pic(Index).Tag)
If x + Margin < msf.ColWidth(Column) Then
    Pic(Index).Width = x + Margin
    MyRefresh Pic(Index)
End If
End Sub

Public Sub ReleaseMouse()
Dim erg As Long
Dim NewRect As RECT
With NewRect
    .Left = 0&
    .Top = 0&
    .Right = Screen.Width / Screen.TwipsPerPixelX
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
End With
erg& = ClipCursor(NewRect)
End Sub

Public Sub LimitMouse(Index As Integer)
Dim x As Long, y As Long, erg As Long
Dim NewRect As RECT, l As Long, lpRect As RECT
l = GetWindowRect(Pic(Index).hwnd, lpRect)
x& = Screen.TwipsPerPixelX
y& = Screen.TwipsPerPixelY
With NewRect
    .Left = lpRect.Left
    .Top = lpRect.Top
    .Right = lpRect.Right
    .Bottom = lpRect.Bottom
End With
erg& = ClipCursor(NewRect)
End Sub


