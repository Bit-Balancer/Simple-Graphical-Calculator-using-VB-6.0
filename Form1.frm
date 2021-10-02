VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator (Beta)Vr.1.0"
   ClientHeight    =   3792
   ClientLeft      =   36
   ClientTop       =   660
   ClientWidth     =   3624
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3792
   ScaleWidth      =   3624
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "a_LCDMini"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   240
      Width           =   3132
   End
   Begin VB.CommandButton Power 
      BackColor       =   &H80000016&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton Digit3 
      BackColor       =   &H80000016&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2280
      Width           =   492
   End
   Begin VB.CommandButton Digit2 
      BackColor       =   &H80000016&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2280
      Width           =   492
   End
   Begin VB.CommandButton Digit1 
      BackColor       =   &H80000016&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2280
      Width           =   492
   End
   Begin VB.CommandButton Digit4 
      BackColor       =   &H80000016&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton Digit5 
      BackColor       =   &H80000016&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton Digit6 
      BackColor       =   &H80000016&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton Digit7 
      BackColor       =   &H80000016&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton Digit8 
      BackColor       =   &H80000016&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton Digit9 
      BackColor       =   &H80000016&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton Digit0 
      BackColor       =   &H80000016&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   492
   End
   Begin VB.CommandButton Digit00 
      BackColor       =   &H80000016&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   492
   End
   Begin VB.CommandButton percentage 
      BackColor       =   &H80000016&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   492
   End
   Begin VB.CommandButton Ezequel 
      BackColor       =   &H80000016&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton Dot 
      BackColor       =   &H80000016&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   492
   End
   Begin VB.CommandButton Divide 
      BackColor       =   &H80000016&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton Multiply 
      BackColor       =   &H80000016&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   492
   End
   Begin VB.CommandButton Plus 
      BackColor       =   &H80000016&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   492
   End
   Begin VB.CommandButton Minus 
      BackColor       =   &H80000016&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   492
   End
   Begin VB.CommandButton C 
      BackColor       =   &H80000016&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   492
   End
   Begin VB.CommandButton Negative 
      BackColor       =   &H80000016&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   492
   End
   Begin VB.CommandButton CE 
      BackColor       =   &H80000016&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   492
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "a_LCDMini"
         Size            =   16.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3132
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu EditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&About"
      Begin VB.Menu AboutHelp 
         Caption         =   "Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu AboutAbout 
         Caption         =   "About Calculator"
      End
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Operand1 As Double, Operand2 As Double
Dim Operator As String
Dim Operator2 As Integer

Sub Define()

End Sub

Private Sub Digit0_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey0 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "0"
      Else
         Display.Caption = Display.Caption & "0"
      End If
  End If
End If
End Sub

Private Sub Digit1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey1 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "1"
      Else
         Display.Caption = Display.Caption & "1"
      End If
  End If
End If
End Sub

Private Sub Digit2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey2 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "2"
      Else
         Display.Caption = Display.Caption & "2"
      End If
  End If
End If
End Sub

Private Sub Digit3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey3 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "3"
      Else
         Display.Caption = Display.Caption & "3"
      End If
  End If
End If
End Sub

Private Sub Digit4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey4 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "4"
      Else
         Display.Caption = Display.Caption & "4"
      End If
  End If
End If
End Sub

Private Sub Digit5_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey5 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "5"
      Else
         Display.Caption = Display.Caption & "5"
      End If
  End If
End If
End Sub

Private Sub Digit6_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey6 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "6"
      Else
         Display.Caption = Display.Caption & "6"
      End If
  End If
End If
End Sub

Private Sub Digit7_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey7 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "7"
      Else
         Display.Caption = Display.Caption & "7"
      End If
  End If
End If
End Sub

Private Sub Digit8_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey8 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "8"
      Else
         Display.Caption = Display.Caption & "8"
      End If
  End If
End If
End Sub

Private Sub Digit9_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKey9 Then
   If Len(Display.Caption) > 16 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "9"
      Else
         Display.Caption = Display.Caption & "9"
      End If
  End If
End If
End Sub

Sub Relay()
On Error GoTo ErrorHandler
If Len(Display.Caption) > 14 Then
MsgBox "Overlode", vbExclamation
Else
If Operator2 > 0 Then
Text1.Text = ""
Operand2 = Display.Caption

If Operator = 1 Then
   If Operand2 = 0 Then
      Text1.Text = "You Can't Devide By Zero"
   Else
   Display.Caption = Operand1 / Operand2
   End If
End If

If Operator = 2 Then
   Display.Caption = Operand1 * Operand2
End If

If Operator = 3 Then
   Display.Caption = Operand1 + Operand2
End If

If Operator = 4 Then
   Display.Caption = Operand1 - Operand2
End If

If Operator = 5 Then
   Display.Caption = Operand1 ^ Operand2
Text1.Text = ""
End If
End If
End If
Exit Sub
ErrorHandler:
MsgBox "Error", vbExclamation
Text1.Text = ""
Display.Caption = "0"
End Sub

Private Sub AboutAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub AboutHelp_Click()
HelpD.Show
End Sub

Private Sub C_Click()
Display.Caption = ""
Display.Caption = "0"
Text1.Text = ""
Operand1 = "0"
Operand2 = "0"
Operator = "0"
Operator2 = 0
End Sub

Private Sub CE_Click()
 Display.Caption = "0"
End Sub

Private Sub Digit0_Click()
 If Len(Display.Caption) > 14 Then
  Beep
  Else
    If Display.Caption = "0" Then
        Display.Caption = "0"
    Else
        Display.Caption = Display.Caption & "0"
    End If
End If
End Sub

Private Sub Digit00_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "0"
      Else
         Display.Caption = Display.Caption & "0" + "0"
      End If
  End If
End Sub

Private Sub Digit1_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "1"
      Else
         Display.Caption = Display.Caption & "1"
      End If
  End If
End Sub

Private Sub Digit2_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "2"
      Else
         Display.Caption = Display.Caption & "2"
      End If
  End If
End Sub

Private Sub Digit3_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "3"
      Else
         Display.Caption = Display.Caption & "3"
      End If
  End If
End Sub

Private Sub Digit4_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "4"
      Else
         Display.Caption = Display.Caption & "4"
      End If
  End If
End Sub

Private Sub Digit5_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "5"
      Else
         Display.Caption = Display.Caption & "5"
      End If
  End If
End Sub

Private Sub Digit6_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "6"
      Else
         Display.Caption = Display.Caption & "6"
      End If
  End If
End Sub

Private Sub Digit7_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "7"
      Else
         Display.Caption = Display.Caption & "7"
      End If
  End If
End Sub

Private Sub Digit8_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "8"
      Else
         Display.Caption = Display.Caption & "8"
      End If
  End If
End Sub

Private Sub Digit9_Click()
  If Len(Display.Caption) > 14 Then
  Beep
  Else
      If Display.Caption = "0" Then
         Display.Caption = "9"
      Else
         Display.Caption = Display.Caption & "9"
      End If
  End If
End Sub

Private Sub Divide_Click()
If Operator2 > 0 Then
Relay
End If
Text1.Text = ""
If Display.Caption = "0" Then
   Display.Caption = "0"
Else
 Text1.Text = Text1.Text & Display.Caption & " " & "/" & " "
End If
 Operand1 = Display.Caption
 Display.Caption = "0"
 Operator = 1 ' 1 for Divide
 Operator2 = 1
End Sub

Private Sub Dot_Click()
If Len(Display.Caption) <> 14 Then
 If InStr(Display.Caption, ".") = False Then
    Display.Caption = Display.Caption & "."
 End If
End If
End Sub

Private Sub EditCopy_Click()
Clipboard.Clear
Clipboard.SetText Display.Caption
End Sub

Private Sub EditPaste_Click()
If IsNumeric(Clipboard.GetText) Then
Display.Caption = Clipboard.GetText
If Len(Display.Caption) > 14 Then
MsgBox "Overlode", vbExclamation
Text1.Text = ""
Display.Caption = "0"
End If
End If
End Sub

Private Sub Ezequel_Click()
On Error GoTo ErrorHandler
If Len(Display.Caption) > 14 Then
MsgBox "Overlode", vbExclamation
Else
Operator2 = 0
Operand2 = Display.Caption
Text1.Text = ""
Text1.Text = "="

If Operand2 = 0 And Operand1 = 0 Then
   Display.Caption = "0"
Else

If Operator = 1 Then
   If Operand2 = 0 Then
      Text1.Text = "You Can't Devide By Zero"
   Else
   Display.Caption = Operand1 / Operand2
   End If
End If

If Operator = 2 Then
   Display.Caption = Operand1 * Operand2
End If

If Operator = 3 Then
   Display.Caption = Operand1 + Operand2
End If

If Operator = 4 Then
   Display.Caption = Operand1 - Operand2
End If

If Operator = 5 Then
   Display.Caption = Operand1 ^ Operand2
End If
End If
End If
Exit Sub
ErrorHandler:
MsgBox "Error", vbExclamation
Text1.Text = ""
Display.Caption = "0"
End Sub

Private Sub Minus_Click()
If Operator2 > 0 Then
Relay
End If
Text1.Text = ""
If Display.Caption = "0" Then
   Display.Caption = "0"
Else
 Text1.Text = Text1.Text & Display.Caption & " " & "-" & " "
End If
 Operand1 = Display.Caption
 Display.Caption = "0"
 Operator = 4 ' 4 for Minus
 Operator2 = 4
End Sub

Private Sub Multiply_Click()
If Operator2 > 0 Then
Relay
End If
Text1.Text = ""
If Display.Caption = "0" Then
   Display.Caption = "0"
Else
 Text1.Text = Text1.Text & Display.Caption & " " & "*" & " "
End If
 Operand1 = Display.Caption
 Display.Caption = "0"
 Operator = 2 ' 2 for Multiply
 Operator2 = 2
End Sub

Private Sub Negative_Click()
If Len(Display.Caption) <> 14 Then
  If Display.Caption = "0" Then
    Display.Caption = "0"
    Else
     If InStr(Display.Caption, "-") = False Then
        Display.Caption = "-" & Display.Caption
     Else
        Display.Caption = Abs(Display.Caption)
     End If
  End If
End If
End Sub

Private Sub percentage_Click()
On Error GoTo ErrorHandler
If Len(Display.Caption) > 14 Then
MsgBox "Overlode", vbExclamation
Else
Dim a, b As Integer
Operand2 = Display.Caption
Text1.Text = ""
Text1.Text = "%"
If Operator = 2 Then
   Display.Caption = Operand1 * Operand2 / 100
End If

If Operator = 3 Then
   a = Operand1 * Operand2 / 100
   Display.Caption = Operand1 + a
End If

If Operator = 4 Then
   b = Operand1 * Operand2 / 100
   Display.Caption = Operand1 - b
End If
End If
Exit Sub
ErrorHandler:
MsgBox "Error", vbExclamation
Text1.Text = ""
Display.Caption = "0"
End Sub

Private Sub Plus_Click()
If Operator2 > 0 Then
Relay
End If
Text1.Text = ""
If Display.Caption = "0" Then
   Display.Caption = "0"
Else
 Text1.Text = Text1.Text & Display.Caption & " " & "+" & " "
End If
 Operand1 = Display.Caption
 Display.Caption = "0"
 Operator = 3 ' 3 for Plus
 Operator2 = 3
End Sub

Private Sub Power_Click()
If Operator2 > 0 Then
Relay
End If
If Display.Caption = "0" Then
   Display.Caption = "0"
Else
 Text1.Text = Text1.Text & Display.Caption & " " & "^" & " "
End If
 Operand1 = Display.Caption
 Display.Caption = "0"
 Operator = 5 ' 5 for Power
 Operator2 = 5
End Sub
