VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Yaesu FT-840 Control Program V 1.0"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9525
   DrawMode        =   4  'Mask Not Pen
   FillColor       =   &H0080C0FF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   360
   End
   Begin VB.CommandButton About 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   69
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Textvfom 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   68
      Top             =   6120
      Width           =   615
   End
   Begin VB.PictureBox Barra 
      Height          =   255
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   7635
      TabIndex        =   49
      Top             =   6720
      Width           =   7695
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   720
      Top             =   120
   End
   Begin VB.TextBox Textmvfo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   6600
      MaxLength       =   3
      TabIndex        =   48
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton MVFO 
      Caption         =   "M>VFO"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   46
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "COM Port Setting"
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   2760
      TabIndex        =   36
      Top             =   2520
      Width           =   4215
      Begin VB.TextBox Text1 
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
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   38
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   615
         Left            =   360
         TabIndex        =   37
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   "COM Port Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Frame frMarco 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   4695
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   6255
      Begin VB.CommandButton DO5 
         BackColor       =   &H0080C0FF&
         Caption         =   "-5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton UP5 
         BackColor       =   &H0080C0FF&
         Caption         =   "+5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2745
         Width           =   735
      End
      Begin VB.CommandButton UP1 
         BackColor       =   &H0080C0FF&
         Caption         =   "+1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton DO1 
         BackColor       =   &H0080C0FF&
         Caption         =   "-1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton Clearfrec 
         BackColor       =   &H0000C000&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "9"
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
         Index           =   9
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "8"
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
         Index           =   8
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "7"
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
         Index           =   7
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "6"
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
         Index           =   6
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "5"
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
         Index           =   5
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "4"
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
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "3"
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
         Index           =   3
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "2"
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
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "1"
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
         Index           =   1
         Left            =   240
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
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
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1035
         Left            =   600
         MaxLength       =   7
         TabIndex        =   17
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdBotones 
         BackColor       =   &H0080C0FF&
         Caption         =   "0"
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
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3960
         Width           =   3375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "UP"
         Height          =   375
         Left            =   3840
         TabIndex        =   45
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "DOWN"
         Height          =   375
         Left            =   3840
         TabIndex        =   44
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Mhz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1035
         Left            =   4560
         TabIndex        =   32
         Top             =   360
         Width           =   1105
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Mhz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   1400
         Width           =   1105
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Khz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1770
         TabIndex        =   30
         Top             =   1395
         Width           =   1450
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Hz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3280
         TabIndex        =   29
         Top             =   1400
         Width           =   1015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         X1              =   1740
         X2              =   1740
         Y1              =   1350
         Y2              =   1600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         X1              =   3255
         X2              =   3255
         Y1              =   1320
         Y2              =   1600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         X1              =   4330
         X2              =   4330
         Y1              =   1320
         Y2              =   1610
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   " <--Scale "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   265
         Left            =   4355
         TabIndex        =   28
         Top             =   1380
         Width           =   1320
      End
      Begin VB.Line Line16 
         BorderWidth     =   7
         X1              =   0
         X2              =   6240
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   13
         Height          =   1500
         Left            =   490
         Top             =   255
         Width           =   5280
      End
   End
   Begin MCI.MMControl MCI 
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   "C:\Yaesu cat\Tone.wav"
   End
   Begin VB.CommandButton VFOM 
      Caption         =   "VFO>M"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton VFOAABB 
      Caption         =   "VFO A=B"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   6120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5040
      Top             =   480
   End
   Begin VB.TextBox Hora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "hh:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Left            =   7080
      TabIndex        =   10
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton PTT 
      BackColor       =   &H80000004&
      Caption         =   "PTT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton CW 
      Caption         =   "CW"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton USB 
      Caption         =   "USB"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton LSB 
      Caption         =   "LSB"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton FM 
      Caption         =   "FM"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton AM 
      Caption         =   "AM"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton VFOAB 
      Caption         =   "VFO A/B"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton LOOCK 
      BackColor       =   &H8000000A&
      Caption         =   "LOCK "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      BaudRate        =   4800
      StopBits        =   2
   End
   Begin VB.CommandButton SPLIT 
      Caption         =   "SPLIT "
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H00808080&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3120
      TabIndex        =   67
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label31 
      BackColor       =   &H00808080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   66
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label30 
      BackColor       =   &H00808080&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3960
      TabIndex        =   65
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label29 
      BackColor       =   &H00808080&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   64
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label28 
      BackColor       =   &H00808080&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7080
      TabIndex        =   63
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label27 
      BackColor       =   &H00808080&
      Caption         =   "150W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8400
      TabIndex        =   62
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label26 
      BackColor       =   &H00808080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   61
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label25 
      BackColor       =   &H00808080&
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   60
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label24 
      BackColor       =   &H00808080&
      Caption         =   "+40"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   59
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label23 
      BackColor       =   &H00808080&
      Caption         =   "+60dB"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8160
      TabIndex        =   58
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label22 
      BackColor       =   &H00808080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   57
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label21 
      BackColor       =   &H00808080&
      Caption         =   "+20"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   56
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2040
      TabIndex        =   55
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00808080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3300
      TabIndex        =   54
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1440
      TabIndex        =   53
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2640
      TabIndex        =   52
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   51
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Caption         =   "Signal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   50
      Top             =   7030
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Memory functions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   47
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "VFO  Functions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Transmit Mode"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   34
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808080&
      Caption         =   "Transmit control"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   33
      Top             =   360
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   585
      Left            =   7080
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   13
      Height          =   4695
      Left            =   360
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      Caption         =   "Operating Frecuency Display"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "YAESU FT-840 CONTROL PROGRAM "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub About_Click()
Form2.Show
End Sub

Private Sub Clearfrec_Click()
txtNumero = ""
MCI.From = 0
MCI.Command = "Play"

End Sub

Private Sub cmdBotones_Click(index As Integer)
Static a As Integer
MCI.From = 0
MCI.Command = "Play"
If txtNumero = "" Then a = 0
txtNumero = txtNumero + Format(index)
a = a + 1
If a = 7 Then
Sendfrec_Click
End If

End Sub




Private Sub AM_Click()
MCI.From = 0
MCI.Command = "Play"
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(4)
MSComm1.Output = Chr$(12)
End Sub

Private Sub Command2_Click()
Dim a As Integer
On Error GoTo Errorman
a = Text1
If a <= 0 Then
GoTo Errorman
End If
If a >= 5 Then
GoTo Errorman
End If
MSComm1.CommPort = Text1
MSComm1.Settings = "4800,N,8,2"
MSComm1.PortOpen = True
Frame1.Visible = False
Form1.Visible = False
Form1.Visible = True
txtNumero = 0 & 100000
Timer2.Enabled = True
Exit Sub
Errorman:
MsgBox "Invalid Port!!!"
End Sub

Private Sub CW_Click()
MCI.From = 0
MCI.Command = "Play"
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(2)
MSComm1.Output = Chr$(12)
End Sub

Private Sub DO1_Click()
On Error GoTo Manejoerror
If txtNumero <= 10000 Then
MsgBox "Out of Range!!!"
Exit Sub
End If
txtNumero.Text = txtNumero - 1
If (txtNumero < 1000000) Or (txtNumero < 99999) Then
txtNumero = 0 & txtNumero
End If
If (txtNumero < 100000) Or (txtNumero < 9999) Then
txtNumero = 0 & txtNumero
End If
Exit Sub
Manejoerror:
MsgBox "Please input a number", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
End Sub

Private Sub DO5_Click()
On Error GoTo Manejoerror
If txtNumero <= 10004 Then
MsgBox "Out of Range!!!"
Exit Sub
End If
txtNumero.Text = txtNumero - 5
If (txtNumero < 1000000) Or (txtNumero < 99999) Then
txtNumero = 0 & txtNumero
End If
If (txtNumero < 100000) Or (txtNumero < 9999) Then
txtNumero = 0 & txtNumero
End If
Exit Sub
Manejoerror:
MsgBox "Please input a number", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
End Sub

Private Sub FM_Click()
MCI.From = 0
MCI.Command = "Play"
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(6)
MSComm1.Output = Chr$(12)
End Sub

Private Sub Form_Load()
Timer2.Enabled = False

MCI.DeviceType = "WaveAudio"
MCI.FileName = App.Path + "\Tone.wav"
MCI.Wait = True
MCI.Notify = False
MCI.Command = "Open"
MCI.UpdateInterval = 100


End Sub



Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = False
For a = 1 To 1000000
Next a
Barra.Value = 0
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(15)
MCI.Command = "Close"
For a = 1 To 100000
Next a
End Sub

Private Sub MVFO_Click()
Dim MVFO As Integer

MCI.From = 0
MCI.Command = "Play"
MVFO = Val(Textmvfo)
If (MVFO < 1) Or (MVFO > 100) Then
MsgBox "Memory number error"
Exit Sub
Else

MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(Textmvfo)
MSComm1.Output = Chr$(6)

End If

End Sub

Private Sub LOOCK_Click()
Static LOOCK As Boolean
MCI.From = 0
MCI.Command = "Play"
If LOOCK = False Then
LOOCK = True
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
MSComm1.Output = Chr$(4)
Else
LOOCK = False
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(4)
End If
End Sub

Private Sub LSB_Click()
MCI.From = 0
MCI.Command = "Play"
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(12)
End Sub


Private Sub PTT_Click()
Static PTT As Boolean
Dim cadena2 As Integer
cadena2 = Len(txtNumero)
If cadena2 < 7 Then
MsgBox "Please input frecuency value ", vbCritical, "Yaesu FT-840 Error"
Exit Sub
End If
MCI.Notify = True
MCI.From = 0
MCI.Command = "Play"
If PTT = False Then
PTT = True
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
MSComm1.Output = Chr$(15)
Else
PTT = False
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(15)
End If
End Sub

Private Sub Sendfrec_Click()
Dim n1$, n2$, n3$, n4$
Dim cadena As Boolean
Dim z1, z2, z3, z4, pa1, pa2, pa3, pa4 As Integer
Static numero As String
On Error GoTo Manejoerror
cadena = txtNumero Like "#######"
If cadena = False Then
MsgBox "Please input frecuency value ", vbCritical, "Yaesu FT-840 Error"
txtNumero = numero
End If
If txtNumero = "" Then
MsgBox "Please input frecuency value ", vbCritical, "Yaesu FT-840 Error"
txtNumero = numero
Exit Sub
End If
If txtNumero >= 3000001 Then
MsgBox "The frecuency value must be < 30.000.00 Mhz", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
txtNumero = numero
Else
If txtNumero <= 9999 Then
MsgBox "The frecuency value must be > 100.00 Khz", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
txtNumero = numero
End If
End If

numero = txtNumero
n1$ = Left(txtNumero, 1)
n2$ = Mid$(txtNumero, 2, 2)
n3$ = Mid$(txtNumero, 4, 2)
n4$ = Right$(txtNumero, 2)
z1 = Val(n1$)
z2 = Val(n2$)
z3 = Val(n3$)
z4 = Val(n4$)
If z2 >= 10 Then
ze2 = Left(z2, 1)
z2 = z2 + (ze2 * 6)
End If
If z3 >= 10 Then
ze3 = Left(z3, 1)
z3 = z3 + (ze3 * 6)
End If
If z4 >= 10 Then
ze4 = Left(z4, 1)
z4 = z4 + (ze4 * 6)
End If
MSComm1.Output = Chr$(z4)
MSComm1.Output = Chr$(z3)
MSComm1.Output = Chr$(z2)
MSComm1.Output = Chr$(z1)
MSComm1.Output = Chr$(10)
MCI.From = 0
MCI.Command = "Play"
Exit Sub
Manejoerror:
MsgBox "Please input a number", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
End Sub

Private Sub SPLIT_Click()
Static SPLIT As Boolean
MCI.From = 0
MCI.Command = "Play"
If SPLIT = False Then
SPLIT = True
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
MSComm1.Output = Chr$(1)
Else
SPLIT = False
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
End If


End Sub

Private Sub Timer1_Timer()
Hora = Time$
End Sub

Private Sub Timer2_Timer()
On Error GoTo error

MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(247)
For a = 1 To 100000
Next a

   
DoEvents
For a = 1 To 5
buffer$ = buffer$ & MSComm1.Input
Next a
b$ = Left(buffer$, 1)
c = Asc(b$)
Barra.Value = c
Exit Sub
error:
c = 220
Resume
End Sub

Private Sub txtNumero_Change()
Dim cadena1 As Integer
MCI.From = 0
MCI.Command = "Play"
cadena1 = Len(txtNumero)
If cadena1 = 7 Then
Sendfrec_Click
End If
End Sub



Private Sub UP1_Click()
On Error GoTo Manejoerror
If txtNumero = 3000000 Then
MsgBox "Out of Range!!!"
Exit Sub
End If
txtNumero.Text = txtNumero + 1
If (txtNumero < 10000) Or (txtNumero <= 99999) Then
txtNumero = 0 & txtNumero
End If
If (txtNumero < 100000) Or (txtNumero <= 999999) Then
txtNumero = 0 & txtNumero
End If
Exit Sub
Manejoerror:
MsgBox "Please input a number", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
End Sub

Private Sub UP5_Click()
On Error GoTo Manejoerror
If txtNumero >= 2999996 Then
MsgBox "Out of Range!!!"
Exit Sub
End If
txtNumero.Text = txtNumero + 5
If (txtNumero < 10000) Or (txtNumero <= 99999) Then
txtNumero = 0 & txtNumero
End If
If (txtNumero < 100000) Or (txtNumero <= 999999) Then
txtNumero = 0 & txtNumero
End If
Exit Sub
Manejoerror:
MsgBox "Please input a number", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
End Sub

Private Sub USB_Click()
MCI.From = 0
MCI.Command = "Play"
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
MSComm1.Output = Chr$(12)
End Sub

Private Sub VFOAABB_Click()
MCI.From = 0
MCI.Command = "Play"
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
MSComm1.Output = Chr$(133)
End Sub

Private Sub VFOAB_Click()
Static VFO As Boolean
MCI.From = 0
MCI.Command = "Play"
If VFO = False Then
VFO = True
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(5)
Else
VFO = False
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(1)
MSComm1.Output = Chr$(5)
End If
End Sub

Private Sub VFOM_Click()
Dim VFOM As Integer
MCI.From = 0
MCI.Command = "Play"
VFOM = Val(Textvfom)
If (VFOM < 1) Or (VFOM > 100) Then
MsgBox "Memory number error"
Exit Sub
Else
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(VFOM)
MSComm1.Output = Chr$(3)
End If
End Sub
