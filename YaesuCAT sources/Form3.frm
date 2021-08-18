VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808080&
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Commenter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox COMnumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Input COM port number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "YAESU FT-840 Control Program V 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public comnumero As Integer
Attribute comnumero.VB_VarMemberFlags = "200"
Sub Commenter_Click()
Dim comport As Integer

CommPort = COMnumber
If CommPort <= 0 Then
MsgBox "Please Enter correct COM port number"
Exit Sub
End If
If CommPort >= 5 Then
MsgBox "Please Enter correct COM port number"
Exit Sub
Else
comnumero = CommPort
Form1.Show
Form3.Hide
End If
End Sub


