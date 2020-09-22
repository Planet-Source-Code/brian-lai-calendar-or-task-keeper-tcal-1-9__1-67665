VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "TCal"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3135
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   7680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkComplete 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Check this if the task is done - then the alarm will not sound."
      Top             =   1080
      Width           =   255
   End
   Begin VB.Timer Tmr1 
      Interval        =   1000
      Left            =   240
      Top             =   1080
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   46
      Text            =   "frmMain.frx":038A
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   11
      Left            =   480
      TabIndex        =   45
      Text            =   "Details"
      Top             =   7905
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "frmMain.frx":038F
      Top             =   7080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   10
      Left            =   480
      TabIndex        =   41
      Text            =   "Details"
      Top             =   7305
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   38
      Text            =   "frmMain.frx":0394
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   37
      Text            =   "Details"
      Top             =   6705
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "frmMain.frx":0399
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   33
      Text            =   "Details"
      Top             =   6105
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "frmMain.frx":039E
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   29
      Text            =   "Details"
      Top             =   5505
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "frmMain.frx":03A3
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   25
      Text            =   "Details"
      Top             =   4905
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "frmMain.frx":03A8
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   21
      Text            =   "Details"
      Top             =   4305
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMain.frx":03AD
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmMain.frx":03B2
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   17
      Text            =   "Details"
      Top             =   3705
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmMain.frx":03B7
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Text            =   "Details"
      Top             =   3105
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmMain.frx":03BC
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   9
      Text            =   "Details"
      Top             =   2505
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAgenda 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmMain.frx":03C1
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Text            =   "Details"
      Top             =   1905
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtDetails 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Text            =   $"frmMain.frx":03C6
      Top             =   1305
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "Upcoming Events"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblSysDate 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "Please Wait..."
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Top             =   480
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   3000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image imgActiveArrow 
      Height          =   165
      Left            =   120
      Picture         =   "frmMain.frx":03D3
      Top             =   1440
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnPrefs 
      Height          =   225
      Left            =   90
      Picture         =   "frmMain.frx":041C
      Top             =   75
      Width           =   225
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   11
      Left            =   120
      Picture         =   "frmMain.frx":049B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   7740
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   48
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   47
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   11
      Left            =   2760
      Picture         =   "frmMain.frx":04EB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   7740
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   11
      Left            =   2760
      Picture         =   "frmMain.frx":053B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   7965
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   10
      Left            =   120
      Picture         =   "frmMain.frx":058B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   7140
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   44
      Top             =   7080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   43
      Top             =   7320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   10
      Left            =   2760
      Picture         =   "frmMain.frx":060B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   7140
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   10
      Left            =   2760
      Picture         =   "frmMain.frx":065B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   7365
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   9
      Left            =   120
      Picture         =   "frmMain.frx":06AB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   6540
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   40
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   39
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   9
      Left            =   2760
      Picture         =   "frmMain.frx":072B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   6540
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   9
      Left            =   2760
      Picture         =   "frmMain.frx":077B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   6765
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   8
      Left            =   120
      Picture         =   "frmMain.frx":07CB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   5940
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   35
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   8
      Left            =   2760
      Picture         =   "frmMain.frx":084B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   5940
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   8
      Left            =   2760
      Picture         =   "frmMain.frx":089B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   6165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   7
      Left            =   120
      Picture         =   "frmMain.frx":08EB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   5340
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   7
      Left            =   2760
      Picture         =   "frmMain.frx":096B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   5340
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   7
      Left            =   2760
      Picture         =   "frmMain.frx":09BB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   5565
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   6
      Left            =   120
      Picture         =   "frmMain.frx":0A0B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   4740
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   6
      Left            =   2760
      Picture         =   "frmMain.frx":0A8B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   4740
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   6
      Left            =   2760
      Picture         =   "frmMain.frx":0ADB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   4965
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   5
      Left            =   120
      Picture         =   "frmMain.frx":0B2B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   4140
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   5
      Left            =   2760
      Picture         =   "frmMain.frx":0BAB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   4140
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   5
      Left            =   2760
      Picture         =   "frmMain.frx":0BFB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   4365
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   4
      Left            =   120
      Picture         =   "frmMain.frx":0C4B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   3540
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   19
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   4
      Left            =   2760
      Picture         =   "frmMain.frx":0CCB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   3540
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   4
      Left            =   2760
      Picture         =   "frmMain.frx":0D1B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   3765
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   3
      Left            =   120
      Picture         =   "frmMain.frx":0D6B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   2940
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   3
      Left            =   2760
      Picture         =   "frmMain.frx":0DEB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   2940
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   3
      Left            =   2760
      Picture         =   "frmMain.frx":0E3B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   3165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   2
      Left            =   120
      Picture         =   "frmMain.frx":0E8B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   2340
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   2
      Left            =   2760
      Picture         =   "frmMain.frx":0F0B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   2340
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   2
      Left            =   2760
      Picture         =   "frmMain.frx":0F5B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   2565
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":0FAB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   1740
      Width           =   165
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   1
      Left            =   2760
      Picture         =   "frmMain.frx":102B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   1740
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   1
      Left            =   2760
      Picture         =   "frmMain.frx":107B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   1965
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgTriangle 
      Height          =   165
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":10CB
      Stretch         =   -1  'True
      ToolTipText     =   "Click to view this event in a new window"
      Top             =   1140
      Width           =   165
   End
   Begin VB.Image imgDetailsEdit 
      Height          =   150
      Index           =   0
      Left            =   2760
      Picture         =   "frmMain.frx":111B
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   1365
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgItemEdit 
      Height          =   150
      Index           =   0
      Left            =   2760
      Picture         =   "frmMain.frx":1165
      Stretch         =   -1  'True
      ToolTipText     =   "Click to edit this field"
      Top             =   1140
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image btnEnd 
      Height          =   210
      Left            =   2850
      Picture         =   "frmMain.frx":11AF
      Stretch         =   -1  'True
      ToolTipText     =   "Close TCal"
      Top             =   82
      Width           =   210
   End
   Begin VB.Label lblAGDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblAgenda 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblAppCaption 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "TCal #VER"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image imgResize 
      Height          =   180
      Left            =   2880
      Picture         =   "frmMain.frx":1338
      Top             =   4080
      Width           =   195
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Dim LastFieldIndex As Integer
Dim FetchDone As Boolean

Private Const MsWidth As Long = 4095
Private Const MlWidth As Long = 8415

Private Sub btnEnd_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub btnPrefs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        frmPrefs.Show 1
    ElseIf Button = 1 Then
        PopupMenu frmPrefs.titMainMenu, , 0, Shape1(0).Height, frmPrefs.titAddItem
    End If
End Sub

Private Sub chkComplete_Click(Index As Integer)
    On Error Resume Next
    Dim OptBuffer As Integer
    If FetchDone = True Then
        SetItemData Index, "ItemDone", chkComplete(Index).Value
        If chkComplete(Index).Value = 1 Then
            OptBuffer = Val(GetSet("AutoDel", "1"))
            If OptBuffer > 0 Then
                If OptBuffer = 1 Then
                    If MsgBox("This task is done. Do you want to delete this task now?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                End If
                DeleteItem Index
                FetchItem Index
            End If
        End If
    End If
End Sub

Private Sub Form_Click()
    On Error Resume Next
    Dim I As Integer
    HideAllTextBoxes
    txtAgenda_LostFocus LastFieldIndex - 1
    DoEvents
    txtDetails_LostFocus LastFieldIndex - 1
    DoEvents
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Integer
    Dim myWAV As String
    'DockingStart Me, True
    If GetSet("StartMinimize", "0") = "1" Then Me.WindowState = 1
    If GetSet("ShowInTaskbar", "1") = "1" Then
        SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
    Else
        SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_APPWINDOW)
    End If
    
    frmPrefs.titAoT.Checked = (GetSet("OnTop", "0") = "1")
    If GetSet("OnTop", "0") = "1" Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    If GetSet("Transparent", "1") = "1" Then MakeTransparent Me.hWnd, GetSet("Transparent Value", "200")
    If GetSet("Black", "0") = "1" Then
        Me.BackColor = RGB(0, 0, 0)
        lblAppCaption.ForeColor = RGB(255, 255, 255)
        lblSysDate.ForeColor = RGB(200, 200, 200)
        Label2.ForeColor = RGB(200, 200, 200)
        For I = 0 To 11 Step 1
            imgItemEdit(I).Picture = imgDetailsEdit(0).Picture
            imgDetailsEdit(I).Picture = imgDetailsEdit(0).Picture
            lblAgenda(I).ForeColor = RGB(255, 255, 255)
            imgTriangle(I).Picture = imgTriangle(11).Picture
            chkComplete(I).BackColor = RGB(0, 0, 0)
        Next
    Else
        For I = 0 To 11 Step 1
            imgItemEdit(I).Picture = imgItemEdit(0).Picture
            imgDetailsEdit(I).Picture = imgItemEdit(0).Picture
            imgTriangle(I).Picture = imgTriangle(0).Picture
        Next
    End If
    lblAppCaption.Caption = Replace(lblAppCaption.Caption, "#VER", MyVer)
    DropShadow Me.hWnd
    MoveForm
    FetchAllItems
    SortData
    InitSticky
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    NewX = X
    NewY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.Move Me.Left + X - NewX, Me.Top + Y - NewY
        If frmDetails.Visible = True Then
            If GetSet("MagneticDetails", "1") = "1" Then MoveDetails
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Shape1(0).Width = Me.ScaleWidth
    Shape1(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    imgResize.Move Shape1(1).Width - 255, Shape1(1).Height - 255
End Sub

Public Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveAllItems
    Dim Dx As Form
    For Each Dx In Forms
        Unload Dx
    Next
    End
    'DockingTerminate Me
End Sub

Private Sub imgDetailsEdit_Click(Index As Integer)
    On Error Resume Next
    txtDetails_KeyPress Index, vbKeyReturn
End Sub

Private Sub imgItemEdit_Click(Index As Integer)
    On Error Resume Next
    txtAgenda_KeyPress Index, vbKeyReturn
End Sub

Public Sub imgResize_Click()
    On Error Resume Next
    Dim I As Integer
    If Me.ScaleHeight <> MsWidth Then
        Me.Height = MsWidth
    Else
        Me.Height = MlWidth
    End If
    For I = 5 To 11 Step 1
        lblAgenda(I).Visible = (Me.ScaleHeight > MsWidth)
        lblAGDetails(I).Visible = lblAgenda(I).Visible
        imgTriangle(I).Visible = lblAgenda(I).Visible
    Next
    If GetSet("ResetMain", "1") = "1" Then MoveForm
    If frmDetails.Visible = True And GetSet("MagneticDetails", "1") = "1" Then MoveDetails
End Sub

Public Sub imgTriangle_Click(Index As Integer)
    On Error Resume Next
    OpenDetails Index
    If GetSet("Animate", "0") = "1" Then Call frmDetails.AnimateForm
    imgTriangle(Index).Picture = imgActiveArrow.Picture
    If lblAgenda(Index).Caption = "Enter Item Here" Then
        frmDetails.txtAgenda.Visible = True
        frmDetails.txtAgenda.SelStart = 0
        frmDetails.txtAgenda.SelLength = Len(frmDetails.txtAgenda.Text)
        frmDetails.txtAgenda.SetFocus
    End If
End Sub

Private Sub lblAGDetails_Click(Index As Integer)
    On Error Resume Next
    HideAllTextBoxes
    LastFieldIndex = Index + 1
    With txtDetails(Index)
        .Text = lblAGDetails(Index).Caption
        .Visible = True
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    imgDetailsEdit(Index).Visible = True
End Sub

Public Sub lblAgenda_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then 'Migrated
        HideAllTextBoxes
        LastFieldIndex = Index + 1
        With txtAgenda(Index)
            .Text = lblAgenda(Index).Caption
            .Visible = True
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        chkComplete(Index).Visible = False
        imgItemEdit(Index).Visible = True
    ElseIf Button = 2 Then
        frmPrefs.titTaskDetails.Tag = Index
        frmPrefs.titDeleteTask.Tag = Index
        PopupMenu frmPrefs.titTaskmenu, , lblAgenda(Index).Left, _
        lblAgenda(Index).Top + lblAgenda(Index).Height, frmPrefs.titTaskDetails
    End If
End Sub

Private Sub lblAppCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Form_MouseDown Button, Shift, X, Y
    ElseIf Button = 2 Then
        PopupMenu frmPrefs.titMainMenu, , , , frmPrefs.titAddItem
    End If
End Sub

Private Sub lblAppCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Public Sub MoveForm()
    On Error Resume Next
    If GetSet("PosDefault", "1") = "1" Then
        Me.Move Screen.Width - Me.Width - 255, Screen.Height - GetTaskbarHeight - Me.Height - 255
    Else
        Me.Move Screen.Width - Me.Width - 255, GetTaskbarHeight + 255
    End If
End Sub

Public Sub MoveDetails()
    On Error Resume Next
    frmDetails.Move Me.Left - frmDetails.Width - 60, Me.Top '+ (Me.Height - frmDetails.Height) / 2
End Sub

Public Sub HideAllTextBoxes()
    On Error Resume Next
    Dim I As Long
    For I = 0 To txtAgenda.Count
        txtAgenda(I).Visible = False
        txtDetails(I).Visible = False
        imgDetailsEdit(I).Visible = False
        imgItemEdit(I).Visible = False
    Next
End Sub

Public Function TimeNow(Optional TheFormat As Integer = 0, Optional DayOffset As Integer = 0) As String
    On Error Resume Next
    Dim ampm As String
    If TheFormat = 0 Then
        TimeNow = Format(Now + DayOffset, "Short Time")
    ElseIf TheFormat = 1 Then
        TimeNow = Replace(Format(Now + DayOffset, "Medium Time") + IIf(Hour(Now) > 12, "PM", "AM"), " ", "")
        If Left$(TimeNow, 1) = "0" Then TimeNow = Mid$(TimeNow, 2)
    ElseIf TheFormat = 2 Then
        TimeNow = Format(Now + DayOffset, "Medium Time") & IIf(Hour(Now) > 12, "PM", "AM")
        If Left$(TimeNow, 1) = "0" Then TimeNow = Mid$(TimeNow, 2)
    End If
    'Debug.Print TheFormat & " - " & LCase(TimeNow)
End Function

Public Function DateNow(Optional DayOffset As Integer = 0) As String
    On Error Resume Next
    DateNow = Format(Now + DayOffset, "Short Date")
End Function

Private Sub Tmr1_Timer()
    On Error Resume Next
    Dim I As Integer, j As Integer
    lblSysDate.Caption = "Today is " & Format(Now, "Short Date")
    For I = 0 To 11
        chkComplete(I).Visible = (lblAgenda(I).Visible = True And txtAgenda(I).Visible = False)
        For j = 0 To 2
            If LCase$(Trim(GetItemData(I, "ItemDate", Format(Now, "Short Date")))) = LCase$(DateNow(1)) Then
                ShowNotification I
                SetItemData I, "ItemNotifyShown0", "1"
            End If
            If LCase$(Trim(GetItemData(I, "ItemAlarmTime", lblAGDetails(I).Caption))) = LCase$(TimeNow(j)) Then
                'Debug.Print LCase$(TimeNow(j))
                If Not Me.Tag = TimeNow Then
                    If GetItemData(I, "ItemDoAlarm", "1") <> "1" Then Exit For
                    'Debug.Print GetItemData(I, "AlarmDay" & Weekday(Now, vbMonday), "1")
                    If chkComplete(I).Value = 1 Then Exit For
                    If GetItemData(I, "AlarmDay" & Weekday(Now)) <> "1" Then Exit For
                    OpenAlarm I
                    frmMain.Tag = TimeNow
                End If
            End If
            DoEvents
        Next
        DoEvents
    Next
End Sub

Private Sub txtAgenda_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If txtAgenda(Index).Text = "" Then txtAgenda(Index).Text = "Enter Item Here"
        lblAgenda(Index).Caption = txtAgenda(Index).Text
        SetItemData Index, "ItemName", txtAgenda(Index).Text
        HideAllTextBoxes
        SortData
        If frmDetails.Visible = True Then OpenDetails Index
    ElseIf KeyAscii = vbKeyTab Then
        lblAGDetails_Click Index
    End If
End Sub

Private Sub txtAgenda_LostFocus(Index As Integer)
    On Error Resume Next
    txtAgenda(Index).Visible = False
    If GetSet("SaveFields2", "1") = "0" Then
        txtAgenda_KeyPress Index, vbKeyReturn
    End If
    SortData
End Sub

Private Sub txtDetails_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyTab And Shift = 1 Then
        lblAgenda_MouseDown Index, 1, 0, 0, 0
    End If
End Sub

Private Sub txtDetails_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Dim I As Integer
    If KeyAscii = vbKeyReturn Then
        If txtDetails(Index).Text = "" Then txtDetails(Index).Text = "Enter Tagline Here"
        lblAGDetails(Index).Caption = txtDetails(Index).Text
        SetItemData Index, "ItemTime", txtDetails(Index).Text
        If GetSet("EventAlarm", "1") = "1" Then 'patch to use alarm time
            SetItemData Index, "ItemAlarmTime", txtDetails(Index).Text
            For I = 1 To 7 Step 1
                If GetItemData(Index, "AlarmDay" & I) = "" Then SetItemData Index, "AlarmDay" & I, "1" 'set all alarms to yes
            Next
        End If
        HideAllTextBoxes
        If frmDetails.Visible = True Then OpenDetails Index
    End If
End Sub

Private Sub txtDetails_LostFocus(Index As Integer)
    On Error Resume Next
    txtDetails(Index).Visible = False
End Sub

Public Function FetchAllItems()
    Dim I As Integer
    For I = 0 To txtAgenda.Count
        FetchItem I
    Next
End Function

Public Function FetchItem(Index As Integer)
    On Error Resume Next
    FetchDone = False
    txtAgenda(Index).Text = GetItemData(Index, "ItemName", "Enter Item Here")
    lblAgenda(Index).Caption = txtAgenda(Index).Text
    txtDetails(Index).Text = GetItemData(Index, "ItemTime", "Enter Tagline Here")
    lblAGDetails(Index).Caption = txtDetails(Index).Text
    chkComplete(Index).Value = Val(GetItemData(Index, "ItemDone", "0"))
    FetchDone = True
End Function

Public Function InitSticky()
    On Error Resume Next
    Dim I As Integer
    If GetSet("RestoreSticky", "1") <> "1" Then Exit Function
    For I = 0 To 11 Step 1
        If GetItemData(I, "ShowSticky", "0") = "1" Then MakeSticky I
        DoEvents
    Next
End Function

Public Function SaveAllItems()
    On Error Resume Next
    Dim I As Integer
    For I = 0 To txtAgenda.Count
        SetItemData I, "ItemName", txtAgenda(I).Text
        SetItemData I, "ItemTime", txtDetails(I).Text
    Next
End Function

Public Function OpenDetails(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    Dim buffer As String
    IsReady = False
    'GetSet("OpenDetails", "1") = "1"
    With frmDetails
        .ClearTriangle
        .lblItemName2.Caption = lblAgenda(Index).Caption
        .txtAgenda.Text = lblAgenda(Index).Caption
        .txtTime.Text = lblAGDetails(Index).Caption
        .txtDate.Text = GetItemData(Index, "ItemDate", Format(Now, "Short Date"))
        .chkAlarmMe.Value = Val(GetItemData(Index, "ItemDoAlarm", "1"))
        For I = 1 To 7 Step 1
            .chkWeekDay(I).Value = Val(GetItemData(Index, "AlarmDay" & I, "1"))
        Next
        buffer = "": buffer = GetItemData(Index, "ItemAlarmTime", lblAGDetails(Index).Caption)
        .txtAlarmTime.Text = buffer
        .chkSticky.Value = Val(GetItemData(Index, "ShowSticky", "0"))
        buffer = "": buffer = GetItemData(Index, "ItemDetails", "Enter Details Here")
        .txtDetails.Text = buffer
        buffer = "": buffer = GetItemData(Index, "ItemImage")
        If Len(buffer) > 0 Then
            .imgAppoint.Picture = LoadPicture(buffer)
        Else
            .imgAppoint.Picture = .imgNothing.Picture
        End If
        MoveDetails
        .Tag = Index: .Show
    End With
    IsReady = True
End Function

Public Function OpenAlarm(Index As Integer)
    On Error Resume Next
    With frmAlarm
        .lblItemData(0).Caption = GetItemData(Index, "ItemName")
        .lblItemData(1).Caption = GetItemData(Index, "ItemTime")
        .lblItemData(2).Caption = GetItemData(Index, "ItemDetails")
        .lblItemData(3).Caption = GetItemData(Index, "ItemAlarmTime")
        .Image1.Picture = LoadPicture(GetItemData(Index, "ItemImage"))
        .Tag = Index
        .Show 1
    End With
End Function

Public Function DeleteItem(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    txtAgenda(Index).Text = ""
    txtDetails(Index).Text = ""
    lblAgenda(Index).Caption = ""
    lblAGDetails(Index).Caption = ""
    SetItemData Index, "ItemTime", ""
    SetItemData Index, "ItemName", ""
    SetItemData Index, "ItemDetails", ""
    SetItemData Index, "ItemImage", ""
    SetItemData Index, "ItemDate", ""
    SetItemData Index, "ItemAlarmTime", ""
    SetItemData Index, "ItemDoAlarm", ""
    SetItemData Index, "ItemDone", ""
    SetItemData Index, "ShowSticky", ""
    SetItemData Index, "StickyX", ""
    SetItemData Index, "StickyY", ""
    SetItemData Index, "StickyW", ""
    SetItemData Index, "StickyH", ""
    SetItemData Index, "StickyClr", ""
    SetItemData Index, "ItemDate", ""
    SetItemData Index, "ItemNotifyShown0", ""
    For I = 1 To 7 Step 1
        SetItemData Index, "AlarmDay" & I, ""
    Next
End Function

Public Function SwapItem(FromWhich As Integer, ToWhich As Integer)
    On Error Resume Next
    Dim ItemData(20) As String 'Amount of data on one INI section
    Dim I As Integer
'read
    ItemData(0) = GetItemData(FromWhich, "ItemName")
    ItemData(1) = GetItemData(FromWhich, "ItemTime")
    ItemData(2) = GetItemData(FromWhich, "ItemAlarmTime")
    ItemData(3) = GetItemData(FromWhich, "ItemDoAlarm")
    For I = 1 To 7 Step 1
        ItemData(I + 3) = GetItemData(FromWhich, "AlarmDay" & I)
    Next
    ItemData(11) = GetItemData(FromWhich, "ItemDetails")
    ItemData(12) = GetItemData(FromWhich, "ItemDone")
    ItemData(13) = GetItemData(FromWhich, "ShowSticky")
    ItemData(14) = GetItemData(FromWhich, "StickyX")
    ItemData(15) = GetItemData(FromWhich, "StickyY")
    ItemData(16) = GetItemData(FromWhich, "ItemDate")
    ItemData(17) = GetItemData(FromWhich, "ItemNotifyShown0")
    ItemData(18) = GetItemData(FromWhich, "StickyW")
    ItemData(19) = GetItemData(FromWhich, "StickyH")
    ItemData(20) = GetItemData(FromWhich, "StickyClr")

'overwrite old
    SetItemData FromWhich, "ItemName", GetItemData(ToWhich, "ItemName")
    SetItemData FromWhich, "ItemTime", GetItemData(ToWhich, "ItemTime")
    SetItemData FromWhich, "ItemAlarmTime", GetItemData(ToWhich, "ItemAlarmTime")
    SetItemData FromWhich, "ItemDoAlarm", GetItemData(ToWhich, "ItemDoAlarm")
    For I = 1 To 7 Step 1
        SetItemData FromWhich, "AlarmDay" & I, GetItemData(ToWhich, "AlarmDay" & I)
    Next
    SetItemData FromWhich, "ItemDetails", GetItemData(ToWhich, "ItemDetails")
    SetItemData FromWhich, "ItemDone", GetItemData(ToWhich, "ItemDone")
    SetItemData FromWhich, "ShowSticky", GetItemData(ToWhich, "ShowSticky")
    SetItemData FromWhich, "StickyX", GetItemData(ToWhich, "StickyX")
    SetItemData FromWhich, "StickyY", GetItemData(ToWhich, "StickyY")
    SetItemData FromWhich, "ItemDate", GetItemData(ToWhich, "ItemDate")
    SetItemData FromWhich, "ItemNotifyShown0", GetItemData(ToWhich, "ItemNotifyShown0")
    SetItemData FromWhich, "StickyW", GetItemData(ToWhich, "StickyW")
    SetItemData FromWhich, "StickyH", GetItemData(ToWhich, "StickyH")
    SetItemData FromWhich, "StickyClr", GetItemData(ToWhich, "StickyClr")

'write new
    SetItemData ToWhich, "ItemName", ItemData(0)
    SetItemData ToWhich, "ItemTime", ItemData(1)
    SetItemData ToWhich, "ItemAlarmTime", ItemData(2)
    SetItemData ToWhich, "ItemDoAlarm", ItemData(3)
    For I = 1 To 7 Step 1
        SetItemData ToWhich, "AlarmDay" & I, ItemData(I + 3)
    Next
    SetItemData ToWhich, "ItemDetails", ItemData(11)
    SetItemData ToWhich, "ItemDone", ItemData(12)
    SetItemData ToWhich, "ShowSticky", ItemData(13)
    SetItemData ToWhich, "StickyX", ItemData(14)
    SetItemData ToWhich, "StickyY", ItemData(15)
    SetItemData ToWhich, "ItemDate", ItemData(16)
    SetItemData ToWhich, "ItemNotifyShown0", ItemData(17)
    SetItemData ToWhich, "StickyW", ItemData(18)
    SetItemData ToWhich, "StickyH", ItemData(19)
    SetItemData ToWhich, "StickyClr", ItemData(20)

End Function

Public Function NextEmptyItem(Optional ExpandWindow As Boolean) As Integer
    On Error Resume Next
    Dim I As Integer
    For I = 0 To 11 Step 1
        If frmMain.lblAgenda(I).Caption = "Enter Item Here" Then
        NextEmptyItem = I
        If ExpandWindow = True And I > 4 Then Call frmMain.imgResize_Click
        Exit Function
        End If
    Next
    NextEmptyItem = -1
End Function

Public Function LastItem() As Integer
    On Error Resume Next
    Dim I As Integer
    For I = 11 To 0 Step -1
        If frmMain.lblAgenda(I).Caption <> "Enter Item Here" Then
        LastItem = I
        Exit Function
        End If
    Next
    LastItem = -1
End Function

Public Function SortData()
    On Error Resume Next
    Dim I As Integer, j As Integer, K As Integer
    For I = 0 To 11 Step 1
        j = NextEmptyItem
        K = LastItem
        If j < K Then
            SwapItem j, K
        End If
        FetchItem j
        FetchItem K
    Next
End Function

Public Function MakeSticky(Index As Integer)
    On Error Resume Next
    DoEvents
    Dim K As New frmSticky
    K.Visible = False
    K.Tag = Index
    K.txtTag.Text = Index
    K.LoadSticky
End Function

Public Function ShowNotification(Index As Integer)
    On Error Resume Next
    If GetItemData(Index, "ItemNotifyShown0", "0") = "1" Then Exit Function
    Dim K As New frmNotify
    K.Tag = Index
    K.txtTag.Text = Index
    K.LoadNotify
End Function
