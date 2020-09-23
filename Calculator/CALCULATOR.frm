VERSION 5.00
Begin VB.Form FRMCALCULATOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   5160
   ClientLeft      =   4725
   ClientTop       =   2295
   ClientWidth     =   3630
   ControlBox      =   0   'False
   Icon            =   "CALCULATOR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTKEYS 
      Height          =   405
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Frame FRADISPLAY 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.Label LBLMEMORY 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label LBLDISPLAY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         TabIndex        =   3
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3375
      End
   End
   Begin VB.Frame FRAKEYS 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
      Begin VB.CommandButton CMDGAME 
         Caption         =   "&Start Game"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton CMDABOUT 
         Caption         =   "?"
         Height          =   375
         Left            =   2040
         TabIndex        =   41
         ToolTipText     =   "About Calculator"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton CMDPERCENT 
         Caption         =   "%"
         Height          =   375
         Left            =   2520
         TabIndex        =   40
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMDCOTAGENT 
         Caption         =   "Cot"
         Height          =   375
         Left            =   2520
         TabIndex        =   39
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton CMDROOT 
         Caption         =   "R"
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         Top             =   2160
         Width           =   400
      End
      Begin VB.CommandButton CMDANTILOG 
         Caption         =   "Anti"
         Height          =   375
         Left            =   3000
         TabIndex        =   37
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton CMDPOWER 
         Caption         =   "P"
         Height          =   375
         Left            =   2520
         TabIndex        =   36
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton CMDCOSECANT 
         Caption         =   "Csc"
         Height          =   375
         Left            =   1560
         TabIndex        =   35
         Top             =   720
         Width           =   405
      End
      Begin VB.CommandButton CMDLOG 
         Caption         =   "Log"
         Height          =   375
         Left            =   3000
         TabIndex        =   34
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMDSECANT 
         Caption         =   "Sec"
         Height          =   375
         Left            =   2040
         TabIndex        =   33
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton CMDEQUAL 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   32
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMD000 
         Caption         =   "000"
         Height          =   375
         Left            =   1080
         TabIndex        =   31
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "1"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   400
      End
      Begin VB.CommandButton CMD0 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMD4 
         Caption         =   "4"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "3"
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   2160
         Width           =   400
      End
      Begin VB.CommandButton CMDDEC 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   26
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMD2 
         Caption         =   "2"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   2160
         Width           =   400
      End
      Begin VB.CommandButton CMDSIGN 
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMDTANGENT 
         Caption         =   "Tan"
         Height          =   375
         Left            =   1080
         TabIndex        =   23
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton CMDCOSINE 
         Caption         =   "Cos"
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton CMDSINE 
         Caption         =   "Sin"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton CMDMM 
         Caption         =   "M--"
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton CMDMP 
         Caption         =   "M+"
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   2160
         Width           =   400
      End
      Begin VB.CommandButton CMDM 
         Caption         =   "M"
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMDMULTIPLY 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   2160
         Width           =   400
      End
      Begin VB.CommandButton CMDDIVIDE 
         Caption         =   "÷"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton CMDMINUS 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMDPLUS 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         TabIndex        =   14
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton CMDCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   405
      End
      Begin VB.CommandButton CMDC 
         Cancel          =   -1  'True
         Caption         =   "C"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton CMDINVERSE 
         Caption         =   "1/x"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMD00 
         Caption         =   "00"
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton CMD9 
         Caption         =   "9"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMD8 
         Caption         =   "8"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMD7 
         Caption         =   "7"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   400
      End
      Begin VB.CommandButton CMD6 
         Caption         =   "6"
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   1680
         Width           =   400
      End
      Begin VB.CommandButton CMD5 
         Caption         =   "5"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1680
         Width           =   400
      End
   End
End
Attribute VB_Name = "FRMCALCULATOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OPERAND1 As Double
Dim OPERAND2 As Double
Dim OPERATOR As String
Dim CLEAR As Boolean

Sub ZERO()
    If LBLDISPLAY.Caption = "0" Then
        LBLDISPLAY.Caption = ""
    End If
End Sub

Sub OPERAND()
    Static NUMBER As Boolean
    Select Case NUMBER
    Case True
        OPERAND2 = Val(LBLDISPLAY.Caption)
        NUMBER = False
    Case False
        OPERAND1 = Val(LBLDISPLAY.Caption)
        NUMBER = True
    End Select
    CLEAR = True
End Sub

Sub CLEARDISPLAY()
    If CLEAR = True Then
        LBLDISPLAY.Caption = ""
        CLEAR = False
    End If
End Sub

Sub EXPONENT()
    If InStr(LBLDISPLAY.Caption, "E") > 0 Then
        
    End If
End Sub

Private Sub CMD0_Click()
    Call ZERO
    Call CLEARDISPLAY
    If LBLDISPLAY.Caption <> "0" Then
    TXTKEYS.Text = TXTKEYS.Text + "0"
    End If
    TXTKEYS.SetFocus
End Sub

Private Sub CMD00_Click()
    Call ZERO
    Call CLEARDISPLAY
    If LBLDISPLAY.Caption <> "" And LBLDISPLAY.Caption <> "0" Then
        TXTKEYS.Text = TXTKEYS.Text + "00"
    End If
    TXTKEYS.SetFocus
End Sub

Private Sub CMD000_Click()
    Call ZERO
    Call CLEARDISPLAY
    If LBLDISPLAY.Caption <> "" And LBLDISPLAY.Caption <> "0" Then
        TXTKEYS.Text = TXTKEYS.Text + "000"
    End If
    TXTKEYS.SetFocus
End Sub

Private Sub CMD1_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "1"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD2_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "2"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD3_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "3"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD4_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "4"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD5_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "5"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD6_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "6"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD7_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "7"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD8_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "8"
    TXTKEYS.SetFocus
End Sub

Private Sub CMD9_Click()
    Call ZERO
    Call CLEARDISPLAY
    TXTKEYS.Text = TXTKEYS.Text + "9"
    TXTKEYS.SetFocus
End Sub

Private Sub CMDABOUT_Click()
    MsgBox "         This calculator is made by Shahid" & vbCrLf & "                          For Rowena." & vbCrLf & "          No help file is included because" & vbCrLf & "   every one knows how to use a calculator." & vbCrLf & "              Use it with care and vote me!             ", vbOKOnly + vbApplicationModal, "About Calculator"
    TXTKEYS.SetFocus
End Sub

Private Sub CMDANTILOG_Click()
    CLEAR = True
    LBLDISPLAY.Caption = Str$(Log(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDC_Click()
    TXTKEYS.Text = ""
    LBLDISPLAY.Caption = "0"
    TXTKEYS.SetFocus
End Sub

Private Sub CMDCE_Click()
    TXTKEYS.Text = ""
    LBLDISPLAY.Caption = "0"
    LBLMEMORY.Caption = "0"
    OPERAND1 = Empty
    OPERAND2 = Empty
    OPERATOR = Empty
    TXTKEYS.SetFocus
End Sub

Private Sub CMDCOSECANT_Click()
    CLEAR = True
    LBLDISPLAY.Caption = Str$(1 / Sin(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDCOSINE_Click()
    CLEAR = True
    CLEAR = True
    LBLDISPLAY.Caption = Str$(Cos(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDCOTAGENT_Click()
    CLEAR = True
    LBLDISPLAY.Caption = Str$(1 / Tan(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDDEC_Click()
    Call ZERO
    Call CLEARDISPLAY
    If InStr(LBLDISPLAY.Caption, ".") Then
        TXTKEYS.SetFocus
        Exit Sub
    ElseIf LBLDISPLAY.Caption = "" Then
        TXTKEYS.Text = TXTKEYS.Text + "0."
        TXTKEYS.SetFocus
    Else
        TXTKEYS.Text = TXTKEYS.Text + "."
        TXTKEYS.SetFocus
    End If
End Sub

Private Sub CMDDIVIDE_Click()
    Call OPERAND
    OPERATOR = "/"
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDEQUAL_Click()
    Call OPERAND
    Select Case OPERATOR
    Case "+"
        LBLDISPLAY.Caption = Str$(OPERAND1 + OPERAND2)
    Case "-"
        LBLDISPLAY.Caption = Str$(OPERAND1 - OPERAND2)
    Case "*"
        LBLDISPLAY.Caption = Str$(OPERAND1 * OPERAND2)
    Case "/"
        If OPERAND2 = 0 Then
            Beep
            LBLDISPLAY.Caption = "Division by zero!"
        Else
            LBLDISPLAY.Caption = Str$(OPERAND1 / OPERAND2)
        End If
    Case "P"
        LBLDISPLAY.Caption = Str$(OPERAND1 ^ OPERAND2)
    Case "R"
        LBLDISPLAY.Caption = Str$(OPERAND1 ^ (1 / OPERAND2))
    Case "%"
        LBLDISPLAY.Caption = Str$((OPERAND1 / OPERAND2) * 100) & "%"
    End Select
    Call EXPONENT
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDEXIT_Click()
    End
End Sub

Private Sub CMDGAME_Click()
    MsgBox "        The game is in progress." + vbCrLf + "          Sorry for inconvience!" + vbCrLf + "Please give me your ideas & help me." + vbCrLf + vbCrLf + "    shahid_shaukat@hotmail.com"
End Sub

Private Sub CMDINVERSE_Click()
    If LBLDISPLAY.Caption <> "0" Then
        LBLDISPLAY.Caption = Str$(1 / Val(LBLDISPLAY.Caption))
    End If
    TXTKEYS.SetFocus
End Sub

Private Sub CMDLOG_Click()
    CLEAR = True
    If Val(LBLDISPLAY.Caption) > 0 Then
        LBLDISPLAY.Caption = Str$(Log(Val(LBLDISPLAY.Caption)))
    Else
        LBLDISPLAY.Caption = "Logrithm of negative number!"
    End If
    TXTKEYS.SetFocus
End Sub

Private Sub CMDM_Click()
    Call OPERAND
    FRADISPLAY.Height = 1150
    FRAKEYS.Top = 1100
    If LBLMEMORY.Visible = False Then
        Me.Top = Me.Top - 300
    End If
    LBLMEMORY.Visible = True
    Me.Height = 4540
    LBLMEMORY.Caption = LBLDISPLAY.Caption
    TXTKEYS.SetFocus
End Sub

Private Sub CMDMINUS_Click()
    Call OPERAND
    OPERATOR = "-"
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDMM_Click()
    Call OPERAND
    LBLDISPLAY = Str$(Val(LBLDISPLAY.Caption) - Val(LBLMEMORY.Caption))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDMP_Click()
    Call OPERAND
    LBLDISPLAY = Str$(Val(LBLDISPLAY.Caption) + Val(LBLMEMORY.Caption))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDMULTIPLY_Click()
    Call OPERAND
    OPERATOR = "*"
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDPERCENT_Click()
    Call OPERAND
    OPERATOR = "%"
    TXTKEYS.SetFocus
End Sub

Private Sub CMDPLUS_Click()
    Call OPERAND
    OPERATOR = "+"
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDPOWER_Click()
    Call OPERAND
    OPERATOR = "P"
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDROOT_Click()
    Call OPERAND
    OPERATOR = "R"
    TXTKEYS.Text = ""
    TXTKEYS.SetFocus
End Sub

Private Sub CMDSECANT_Click()
    LBLDISPLAY.Caption = Str$(1 / Cos(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDSIGN_Click()
    LBLDISPLAY.Caption = Str$(-1 * Val(LBLDISPLAY.Caption))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDSINE_Click()
    CLEAR = True
    LBLDISPLAY.Caption = Str$(Sin(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub CMDTANGENT_Click()
    CLEAR = True
    LBLDISPLAY.Caption = Str$(Tan(Val(LBLDISPLAY.Caption)))
    TXTKEYS.SetFocus
End Sub

Private Sub Form_Click()
    TXTKEYS.SetFocus
End Sub

Private Sub Form_GotFocus()
    TXTKEYS.SetFocus
End Sub

Private Sub Form_Load()
    FRADISPLAY.Height = 780
    LBLMEMORY.Visible = False
    FRAKEYS.Top = 810
    Me.Height = 4340
End Sub

Private Sub FRADISPLAY_Click()
    TXTKEYS.SetFocus
End Sub

Private Sub FRAKEYS_Click()
    TXTKEYS.SetFocus
End Sub

Private Sub LBLDISPLAY_Click()
    TXTKEYS.SetFocus
End Sub

Private Sub LBLMEMORY_Click()
    TXTKEYS.SetFocus
End Sub

Private Sub TXTKEYS_Change()
    If Len(LBLDISPLAY.Caption) < 31 Then
        LBLDISPLAY.Caption = LBLDISPLAY.Caption + TXTKEYS.Text
        TXTKEYS.Text = ""
    Else
        Exit Sub
    End If
End Sub

Private Sub TXTKEYS_KeyPress(KeyAscii As Integer)
    Call ZERO
    Call CLEARDISPLAY
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyDelete And InStr(LBLDISPLAY.Caption, ".") = 0) Then
        If KeyAscii = vbKeyDelete Then
            Call CMDDEC_Click
            KeyAscii = 0
        ElseIf KeyAscii = vbKey0 Then
            Call CMD0_Click
            KeyAscii = 0
        End If
        Exit Sub
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub TXTKEYS_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyAdd
        CMDPLUS_Click
    Case vbKeySubtract
        CMDMINUS_Click
    Case vbKeyMultiply
        CMDMULTIPLY_Click
    Case vbKeyDivide
        CMDDIVIDE_Click
    End Select
End Sub
