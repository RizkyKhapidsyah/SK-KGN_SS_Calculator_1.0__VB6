VERSION 5.00
Begin VB.Form frm_calc 
   AutoRedraw      =   -1  'True
   Caption         =   "Calculator"
   ClientHeight    =   5175
   ClientLeft      =   3510
   ClientTop       =   2355
   ClientWidth     =   5430
   DrawMode        =   16  'Merge Pen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5430
   Begin VB.TextBox msv 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   555
      Left            =   2880
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox msv_res 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   555
      Left            =   1320
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btn_9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton btn_8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton btn_7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton btn_6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton btn_5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton btn_4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton btn_3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton btn_2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton btn_1 
      BackColor       =   &H00FF8080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton btn_eq 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   31
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton btn_sqrt 
         Caption         =   "sqrt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   30
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton btn_plus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   29
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton btn_minus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   28
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton btn_per 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   27
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton btn_pm 
         Caption         =   "+/-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   26
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton btn_c 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   25
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton btn_divide 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   24
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton btn_multi 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   23
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton btn_mc 
         Caption         =   "MC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   22
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton btn_mr 
         Caption         =   "MR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   21
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton btn_ms 
         Caption         =   "MS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   20
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton btn_mp 
         Caption         =   "M+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   19
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton btn_dec 
         Caption         =   "."
         Height          =   495
         Left            =   2055
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton btn_0 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   17
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox operator 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox disp_screen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox lv 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Operation In Process"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label LBL_LV 
         BackColor       =   &H00C0C000&
         Caption         =   "Last Value Entered"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu clear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Index           =   2
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu helpt 
         Caption         =   "&Help Topics"
      End
   End
End
Attribute VB_Name = "frm_calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset

Private Sub about_Click()
frmAbout.Show

End Sub

Private Sub btn_0_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 0
Else
        disp_screen.Text = disp_screen.Text + "0"
        
End If
End Sub

Private Sub btn_1_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 1
Else
        disp_screen.Text = disp_screen.Text + "1"
        
End If

End Sub

Private Sub btn_2_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 2
Else
        disp_screen.Text = disp_screen.Text + "2"
        
End If

End Sub

Private Sub btn_3_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 3
Else
        disp_screen.Text = disp_screen.Text + "3"
        
End If
End Sub

Private Sub btn_4_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 4
Else
        disp_screen.Text = disp_screen.Text + "4"
        
End If
End Sub

Private Sub btn_5_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 5
Else
        disp_screen.Text = disp_screen.Text + "5"
        
End If
End Sub

Private Sub btn_6_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 6
Else
        disp_screen.Text = disp_screen.Text + "6"
        
End If
End Sub

Private Sub btn_7_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 7
Else
        disp_screen.Text = disp_screen.Text + "7"
        
End If
End Sub

Private Sub btn_8_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 8
Else
        disp_screen.Text = disp_screen.Text + "8"
        
End If
End Sub

Private Sub btn_9_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = 9
Else
        disp_screen.Text = disp_screen.Text + "9"
        
End If
End Sub

Private Sub btn_c_Click()
lv.Text = " "
disp_screen.Text = " "
operator.Text = " "
msv.Text = " "
msv_res.Text = " "
disp_screen.SetFocus
End Sub

Private Sub btn_dec_Click()
If disp_screen.Text = " " Then
        disp_screen.Text = "."
Else
        disp_screen.Text = disp_screen.Text + "."
        
End If
End Sub

Private Sub btn_divide_Click()
operator.Text = "/"
lv.Text = disp_screen.Text
disp_screen.Text = " "

End Sub

Private Sub btn_eq_Click()
If operator.Text = "+" Then
disp_screen.Text = Val(lv.Text) + Val(disp_screen.Text)
lv.Text = " "
disp_screen.SetFocus

ElseIf operator.Text = "-" Then
disp_screen.Text = Val(lv.Text) - Val(disp_screen.Text)
lv.Text = " "
disp_screen.SetFocus

ElseIf operator.Text = "*" Then
disp_screen.Text = Val(lv.Text) * Val(disp_screen.Text)
lv.Text = " "
disp_screen.SetFocus

ElseIf operator.Text = "/" Then
disp_screen.Text = Val(lv.Text) / Val(disp_screen.Text)
lv.Text = " "
disp_screen.SetFocus

ElseIf operator.Text = "%" Then
disp_screen.Text = Val(lv.Text) * Val(disp_screen.Text) / 100
lv.Text = " "
disp_screen.SetFocus

ElseIf operator.Text = "sqrt" Then
disp_screen.Text = Sqr(lv.Text)
lv.Text = " "
disp_screen.SetFocus
End If




End Sub

Private Sub btn_mc_Click()
msv_res.Text = " "
msv.Text = " "
operator.Text = " "
End Sub

Private Sub btn_minus_Click()
operator.Text = "-"
lv.Text = disp_screen.Text
disp_screen.Text = " "



End Sub

Private Sub btn_mp_Click()
msv_res.Text = Val(disp_screen.Text) + Val(msv.Text)

End Sub

Private Sub btn_mr_Click()
If msv_res.Text = 0 Then
disp_screen.Text = msv.Text
Else
    disp_screen.Text = msv_res.Text
    
End If
End Sub

Private Sub btn_ms_Click()
operator.Text = "MS"
msv.Text = disp_screen.Text
disp_screen.Text = " "
End Sub

Private Sub btn_multi_Click()
operator.Text = "*"
lv.Text = disp_screen.Text
disp_screen.Text = " "

End Sub

Private Sub btn_per_Click()
operator.Text = "%"
lv.Text = disp_screen.Text
disp_screen.Text = " "
End Sub

Private Sub btn_plus_Click()
operator.Text = "+"
lv.Text = disp_screen.Text
disp_screen.Text = " "

End Sub

Private Sub btn_pm_Click()
'If disp_screen.Text <> " " Then
'substr(disp_screen.Text, 1, 1) = disp_screen.Text + "-"
' End If

 
 
End Sub

Private Sub btn_sqrt_Click()
operator.Text = "sqrt"
lv.Text = disp_screen.Text
disp_screen.Text = " "

End Sub

Private Sub clear_Click()
lv.Text = " "
disp_screen.Text = " "
operator.Text = " "
msv.Text = " "
msv_res.Text = " "
disp_screen.SetFocus
End Sub

Private Sub disp_screen_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
lv.Text = disp_screen.Text
operator.Text = "+"

ElseIf KeyAscii = 45 Then
lv.Text = disp_screen.Text
disp_screen.Text = " "
operator.Text = "-"

ElseIf KeyAscii = 42 Then
lv.Text = disp_screen.Text
disp_screen.Text = " "
operator.Text = "*"

ElseIf KeyAscii = 47 Then
lv.Text = disp_screen.Text
disp_screen.Text = " "
operator.Text = "/"

ElseIf KeyAscii = 61 Then
    If operator.Text = "+" Then
    disp_screen.Text = Val(lv.Text) + Val(disp_screen.Text)
    ElseIf operator.Text = "-" Then
    disp_screen.Text = Val(lv.Text) - Val(disp_screen.Text)
    ElseIf operator.Text = "*" Then
    disp_screen.Text = Val(lv.Text) * Val(disp_screen.Text)
    ElseIf operator.Text = "/" Then
    disp_screen.Text = Val(lv.Text) / Val(disp_screen.Text)
    End If
    
ElseIf KeyAscii = 13 Then
    If operator.Text = "+" Then
    disp_screen.Text = Val(lv.Text) + Val(disp_screen.Text)
    ElseIf operator.Text = "-" Then
    disp_screen.Text = Val(lv.Text) - Val(disp_screen.Text)
    ElseIf operator.Text = "*" Then
    disp_screen.Text = Val(lv.Text) * Val(disp_screen.Text)
    ElseIf operator.Text = "/" Then
    disp_screen.Text = Val(lv.Text) / Val(disp_screen.Text)
    End If
End If


End Sub

Private Sub exit_Click()
End
End Sub

