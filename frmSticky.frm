VERSION 5.00
Begin VB.Form frmSticky 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Sticky"
   ClientHeight    =   2775
   ClientLeft      =   -3000
   ClientTop       =   -3000
   ClientWidth     =   2895
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
   Icon            =   "frmSticky.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   -408
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   2760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image btnResize 
      Height          =   165
      Left            =   2730
      Picture         =   "frmSticky.frx":000C
      Top             =   2610
      Width           =   165
   End
   Begin VB.Label lblTagline 
      BackStyle       =   0  '³z©ú
      Caption         =   "Tagline"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblDetails 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  '³z©ú
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lblItemName 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "ItemName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2415
   End
   Begin VB.Image btnEnd 
      Height          =   210
      Left            =   2605
      Picture         =   "frmSticky.frx":0058
      Stretch         =   -1  'True
      ToolTipText     =   "Close TCal"
      Top             =   75
      Width           =   210
   End
   Begin VB.Image imgEvent 
      Height          =   2775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmSticky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long

Private Sub btnEnd_Click()
    On Error Resume Next
    SetItemData Val(Me.Tag), "ShowSticky", "0"
    Unload Me
End Sub

Public Function MyTag() As Integer
    On Error Resume Next
    MyTag = Val(txtTag.Text)
End Function

Private Sub btnResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    NewX = X
    NewY = Y
End Sub

Private Sub btnResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.Move Me.Left, Me.Top, Me.Width - NewX + X, Me.Height - NewY + Y
    End If
End Sub

Private Sub btnResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    SetItemData Val(Me.Tag), "StickyW", Me.Width
    SetItemData Val(Me.Tag), "StickyH", Me.Height
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Integer, j As Long
    Me.Visible = False
    'DockingStart Me
    If GetSet("OnTop", "0") = "1" Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    If GetSet("Transparent", "1") = "1" Then MakeTransparent Me.hWnd, GetSet("Transparent Value", "200")
    If GetSet("Black", "0") = "1" Then
        Me.BackColor = RGB(0, 0, 0)
        lblTime.ForeColor = RGB(255, 255, 255)
        lblItemName.ForeColor = RGB(255, 255, 255)
        lblTagline.ForeColor = RGB(255, 255, 255)
        lblDetails.ForeColor = RGB(255, 255, 255)
    End If
    btnResize.Visible = (GetSet("StickyResize", "1") = "1")
    If GetSet("StickyBack", "0") = "1" Then imgEvent.Picture = LoadPicture(GetItemData(MyTag, "ItemImage"))
    DropShadow Me.hWnd
    LoadSticky
End Sub

Public Function LoadSticky()
    On Error Resume Next
    lblItemName.Caption = GetItemData(MyTag, "ItemName")
    lblTime.Caption = GetItemData(MyTag, "ItemAlarmTime")
    lblTagline.Caption = GetItemData(MyTag, "ItemTime")
    lblDetails.Caption = GetItemData(MyTag, "ItemDetails")
    Me.Move Val(GetItemData(MyTag, "StickyX", Str(frmMain.Left - 1000))), Val(GetItemData(MyTag, "StickyY", Str(frmMain.Top))), _
                    Val(GetItemData(MyTag, "StickyW", "2895")), Val(GetItemData(MyTag, "StickyH", "2775"))
    j = Val(GetItemData(MyTag, "StickyClr", "16777215"))
    If IsNumeric(j) = True Then Me.BackColor = j
    Me.Visible = True
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    SetItemData Val(Me.Tag), "StickyX", Me.Left
    SetItemData Val(Me.Tag), "StickyY", Me.Top
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim j As Long
    Shape1(0).Width = Me.ScaleWidth
    Shape1(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    j = Me.ScaleWidth - 240
    lblTagline.Width = j
    lblDetails.Width = j
    lblTime.Width = j
    btnEnd.Move Me.ScaleWidth - 300, 82
    lblItemName.Width = Me.ScaleWidth - 500
    lblDetails.Height = Me.ScaleHeight - lblDetails.Top - 120
    Line1.X2 = j + 60
    btnResize.Move Me.ScaleWidth - btnResize.Width, Me.ScaleHeight - btnResize.Height
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    NewX = X
    NewY = Y
    If Button = 2 Then
        PopupMenu frmPrefs.titStickyMenu
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then Me.Move Me.Left + X - NewX, Me.Top + Y - NewY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'DockingTerminate Me
End Sub


Private Sub imgEvent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgEvent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgEvent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblItemName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblItemName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblItemName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblTagline_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblTagline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblTagline_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseUp Button, Shift, X, Y
End Sub
