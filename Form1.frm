VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9540
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Sty 
      Height          =   405
      Left            =   2040
      TabIndex        =   12
      Text            =   "0"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Estilo de fuente."
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Tamaño de fuente."
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox C 
      Height          =   405
      Left            =   2040
      TabIndex        =   9
      Text            =   "0"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox size 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Text            =   "10"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox color 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Color de fuente"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Color de fondo."
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Subrayar."
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Oblicua."
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "NEGRITA."
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":1D52B
      Top             =   1440
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub Check1_Click()
If Check1.Value = 1 Then
        Text1.FontBold = True
    Else
        Text1.FontBold = False
        
    End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
        Text1.FontItalic = True
    Else
        Text1.FontItalic = False
        
    End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
        Text1.FontUnderline = True
    Else
        Text1.FontUnderline = False
    End If
End Sub

Private Sub Check4_Click()
x = color
 
If x = 1 Then
 x = &HFF80FF
 ElseIf x = 2 Then
 x = &HFF8080
 ElseIf x = 3 Then
 x = &HFFFF80
  ElseIf x = 4 Then
 x = &H80FF80
 ElseIf x = 5 Then
 x = &H80FFFF
   ElseIf x = 4 Then
 x = &H80C0FF
 ElseIf x = 5 Then
 x = &H8080FF
 ElseIf x = 0 Then
 x = &H4000&
 End If
 If Check4.Value = 1 Then
   Text1.BackColor = x
   Else
   Text1.BackColor = &H80000005
   End If
End Sub


Private Sub Check5_Click()
x = C
If x = 1 Then
 x = &HFF80FF
 ElseIf x = 2 Then
 x = &HFF8080
 ElseIf x = 3 Then
 x = &HFFFF80
  ElseIf x = 4 Then
 x = &H80FF80
 ElseIf x = 5 Then
 x = &H80FFFF
   ElseIf x = 4 Then
 x = &H80C0FF
 ElseIf x = 5 Then
 x = &H8080FF
 ElseIf x = 0 Then
 x = &H4000&
 End If

If Check5.Value = 1 Then
   Text1.ForeColor = x
   Else
   Text1.ForeColor = &H80000012
   End If
End Sub

Private Sub Check6_Click()
y = size
If Check6.Value = 1 Then
   Text1.FontSize = y
   Else
   Text1.FontSize = 10
   End If

End Sub

Private Sub Combo1_Change()
Dim font_family As FontFamily
Dim installed_fonts As New InstalledFontCollection

End Sub

Private Sub Check7_Click()
Text1.Font = Sty
If Check7.Value = 1 Then
   Text1.Font = Sty
   Else
   Text1.Font = ARIAL
   End If


End Sub

Private Sub Command1_Click()
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0





End Sub







