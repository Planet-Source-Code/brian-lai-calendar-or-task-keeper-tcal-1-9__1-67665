VERSION 5.00
Begin VB.Form frmNotify 
   Appearance      =   0  '¥­­±
   BackColor       =   &H80000005&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Notify"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   4480
      TabIndex        =   2
      Top             =   840
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   3360
      Top             =   360
   End
   Begin VB.Label lblReminder 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin TCal.Fader Fader1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1138
      _ExtentY        =   450
      FadeInSpeed     =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblItemName 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "TCal - Notification"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Click to edit this field"
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long

Private Sub Form_Activate()
    On Error Resume Next
    Fader1.FadeIn
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Shape1(0).Width = Me.ScaleWidth
    Shape1(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Integer
    If GetSet("Transparent", "1") = "1" Then MakeTransparent Me.hWnd, GetSet("Transparent Value", "200")
    If GetSet("Black", "0") = "1" Then
        Me.BackColor = RGB(0, 0, 0)
        lblItemName.ForeColor = RGB(255, 255, 255)
        lblReminder.ForeColor = RGB(255, 255, 255)
    End If
    DropShadow Me.hWnd
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 1
    Me.Move Screen.Width - Me.Width, Screen.Height - GetTaskbarHeight - Me.Height - NotifyTop
    NotifyTop = NotifyTop + Me.Height
    LoadNotify
End Sub

Public Sub LoadNotify()
    On Error Resume Next
    lblReminder.Caption = "The event """ & GetItemData(Val(txtTag.Text), "ItemName") & _
        """ is going to take place tomorrow."
    Me.Visible = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    NewX = X
    NewY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then Me.Move Me.Left + X - NewX, Me.Top + Y - NewY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Fader1.FadeOut
    NotifyTop = NotifyTop - Me.Height
End Sub

Private Sub lblItemName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblItemName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblReminder_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Unload Me
End Sub
