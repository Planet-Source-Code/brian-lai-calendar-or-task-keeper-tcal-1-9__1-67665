VERSION 5.00
Begin VB.Form frmDetails 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstTab 
      Height          =   3255
      IntegralHeight  =   0   'False
      ItemData        =   "frmDetails.frx":0000
      Left            =   120
      List            =   "frmDetails.frx":000D
      TabIndex        =   20
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   3255
      Index           =   0
      Left            =   1200
      ScaleHeight     =   3255
      ScaleWidth      =   4455
      TabIndex        =   21
      Top             =   480
      Width           =   4455
      Begin VB.CheckBox chkSticky 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         Caption         =   "Show a sticky note for this event on the desktop"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox txtDetails 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  '««ª½±²¶b
         TabIndex        =   3
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtAgenda 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtTime 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Details for this event"
         Height          =   195
         Left            =   960
         TabIndex        =   30
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Image imgAppoint 
         Appearance      =   0  '¥­­±
         BorderStyle     =   1  '³æ½u©T©w
         Height          =   735
         Left            =   120
         OLEDropMode     =   1  '¤â°Ê
         Picture         =   "frmDetails.frx":0033
         Stretch         =   -1  'True
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Event name:"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Tagline:"
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   22
         Top             =   600
         Width           =   570
      End
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   3255
      Index           =   3
      Left            =   1200
      ScaleHeight     =   3255
      ScaleWidth      =   4455
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   3255
      Index           =   1
      Left            =   1200
      ScaleHeight     =   3255
      ScaleWidth      =   4455
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3076
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.CheckBox chkAlarmMe 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         Caption         =   "Alarm Me on"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtAlarmTime 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This alarm should only alarm on"
         Height          =   1575
         Left            =   0
         TabIndex        =   26
         ToolTipText     =   "Choose only the days you want the alarm to sound."
         Top             =   1680
         Width           =   4455
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sun"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   18
            Top             =   1200
            Width           =   735
         End
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sat"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   960
            TabIndex        =   17
            Top             =   1200
            Width           =   735
         End
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fri"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   960
            TabIndex        =   15
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Thu"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3480
            TabIndex        =   14
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Wed"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   13
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tue"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   12
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkWeekDay 
            Appearance      =   0  '¥­­±
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mon"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkChkAllWeeddays 
            Appearance      =   0  '¥­­±
            BackColor       =   &H80000005&
            Caption         =   "Weekdays"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkChkAllWeekends 
            Appearance      =   0  '¥­­±
            BackColor       =   &H80000005&
            Caption         =   "Weekends"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.CommandButton btnSelTime 
         Caption         =   "..."
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Event Date"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   3255
      Index           =   2
      Left            =   1200
      ScaleHeight     =   3255
      ScaleWidth      =   4455
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   2460
         IntegralHeight  =   0   'False
         ItemData        =   "frmDetails.frx":1AE5
         Left            =   0
         List            =   "frmDetails.frx":1B04
         Style           =   1  '¶µ¥Ø¥]§t®Ö¨ú¤è¶ô
         TabIndex        =   19
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '³z©ú
         Caption         =   "Choose times for reminders to show up. (not made yet)"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Image Image2 
      Height          =   165
      Left            =   5160
      Picture         =   "frmDetails.frx":1BC9
      Top             =   3840
      Width           =   525
   End
   Begin VB.Image btnApply 
      Height          =   210
      Left            =   5190
      Picture         =   "frmDetails.frx":20AF
      Stretch         =   -1  'True
      ToolTipText     =   "Click to apply changes"
      Top             =   75
      Width           =   210
   End
   Begin VB.Image imgNothing 
      Appearance      =   0  '¥­­±
      BorderStyle     =   1  '³æ½u©T©w
      Height          =   735
      Left            =   5760
      OLEDropMode     =   1  '¤â°Ê
      Picture         =   "frmDetails.frx":222B
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   90
      Picture         =   "frmDetails.frx":3CDD
      Top             =   90
      Width           =   195
   End
   Begin VB.Label btnDelItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Delete this item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1110
   End
   Begin VB.Label lblItemName2 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "ItemName (4095 x 5775)"
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
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Double Click to edit this field"
      Top             =   60
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Image btnEnd 
      Height          =   210
      Left            =   5475
      Picture         =   "frmDetails.frx":3EFF
      Stretch         =   -1  'True
      ToolTipText     =   "Close this window"
      Top             =   75
      Width           =   210
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Dim FormReady As Boolean, InfoEdited As Boolean

Private Sub btnApply_Click()
    On Error Resume Next
    Dim I As Integer
    SetItemData Val(Me.Tag), "ItemName", txtAgenda.Text
    SetItemData Val(Me.Tag), "ItemTime", txtTime.Text
    SetItemData Val(Me.Tag), "ItemDetails", txtDetails.Text
    SetItemData Val(Me.Tag), "ItemDoAlarm", Str(chkAlarmMe.Value)
    SetItemData Val(Me.Tag), "ItemDate", txtDate.Text
    SetItemData Val(Me.Tag), "ItemNotifyShown0", "0"
    For I = 1 To 7 Step 1
        SetItemData Val(Me.Tag), "AlarmDay" & I, chkWeekDay(I).Value
    Next
    frmMain.FetchItem Val(Me.Tag)
End Sub

Public Sub btnDelItem_Click()
    On Error Resume Next
    If MsgBox("Are you sure you want to delete this entry?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    frmMain.DeleteItem Val(Me.Tag)
    frmMain.FetchItem Val(Me.Tag)
    Unload Me
End Sub

Private Sub btnEnd_Click()
    On Error Resume Next
    If InfoEdited = True Then
        If MsgBox("Do you want to save the changes you have made?", vbYesNo + vbQuestion) = vbYes Then
            btnApply_Click
        End If
    End If
    FormReady = False
    InfoEdited = False
    Unload Me
End Sub

Private Sub btnSelTime_Click()
    On Error Resume Next
    Dim UserDate As Date
    UserDate = Date
    If frmCalendar.GetDate(UserDate) Then
        txtAlarmTime.Text = UserDate
    End If
End Sub

Private Sub chkAlarmMe_Click()
    On Error Resume Next
    If IsReady = True Then SetItemData Val(Me.Tag), "ItemDoAlarm", Str(chkAlarmMe.Value)
    FrameAlarm.Enabled = (chkAlarmMe.Value = 1)
End Sub


Private Sub chkChkAllWeeddays_Click()
    On Error Resume Next
    For I = 2 To 6 Step 1
        chkWeekDay(I).Value = chkChkAllWeeddays.Value
    Next
End Sub

Private Sub chkChkAllWeekends_Click()
    On Error Resume Next
    chkWeekDay(1).Value = chkChkAllWeekends.Value
    chkWeekDay(7).Value = chkChkAllWeekends.Value
End Sub

Private Sub chkSticky_Click()
On Error Resume Next
If IsReady = True Then
    If chkSticky.Value = 1 Then
        frmMain.MakeSticky Val(Me.Tag)
    Else
        Dim K As Form
        For Each K In Forms
            If K.Tag = Me.Tag And K.ScaleHeight <> Me.ScaleHeight And K.ScaleWidth <> Me.ScaleWidth Then
                Unload K
            End If
        Next
    End If
    SetItemData Val(Me.Tag), "ShowSticky", Str(chkSticky.Value)
End If
End Sub

Private Sub chkWeekDay_Click(Index As Integer)
On Error Resume Next
If IsReady = True Then
    SetItemData Val(Me.Tag), "AlarmDay" & Index, chkWeekDay(Index).Value
End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Integer
    FormReady = False
    If GetSet("Transparent", "1") = "1" Then MakeTransparent Me.hWnd, GetSet("Transparent Value", "200")
    If GetSet("OnTop", "0") = "1" Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    If GetSet("Black", "0") = "1" Then
        chkSticky.BackColor = RGB(0, 0, 0)
        chkSticky.ForeColor = RGB(255, 255, 255)
        Me.BackColor = RGB(0, 0, 0)
        lblItemName.ForeColor = RGB(255, 255, 255)
        lblItemName2.ForeColor = RGB(255, 255, 255)
        Label1.ForeColor = RGB(200, 200, 200)
        Label2(0).ForeColor = RGB(200, 200, 200)
        Label2(1).ForeColor = RGB(200, 200, 200)
        Label2(2).ForeColor = RGB(200, 200, 200)
        lblAGDetails.ForeColor = RGB(255, 255, 255)
        lblTime.ForeColor = RGB(255, 255, 255)
        chkAlarmMe.BackColor = RGB(0, 0, 0)
        chkAlarmMe.ForeColor = RGB(200, 200, 200)
        Frame1.BackColor = RGB(0, 0, 0)
        Frame1.ForeColor = RGB(200, 200, 200)
        LstTab.BackColor = RGB(0, 0, 0)
        LstTab.ForeColor = RGB(200, 200, 200)
        chkChkAllWeeddays.ForeColor = RGB(200, 200, 200)
        chkChkAllWeeddays.BackColor = RGB(0, 0, 0)
        chkChkAllWeekends.ForeColor = RGB(200, 200, 200)
        chkChkAllWeekends.BackColor = RGB(0, 0, 0)
        For I = 1 To 7 Step 1
            If I < 5 Then picTab(I - 1).BackColor = RGB(0, 0, 0)
            chkWeekDay(I).BackColor = RGB(0, 0, 0)
            chkWeekDay(I).ForeColor = RGB(200, 200, 200)
        Next
    End If
    'DockingStart Me
    DropShadow Me.hWnd
End Sub

Private Sub Form_Paint()
    On Error Resume Next
    FormReady = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Shape1(0).Width = Me.ScaleWidth
    Shape1(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
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
    frmMain.FetchItem Val(Me.Tag)
    ClearTriangle
    'DockingTerminate Me
End Sub

Private Sub Image2_Click()
    On Error Resume Next
    Shell "explorer http://thinc.no-ip.info", vbNormalFocus
End Sub

Private Sub imgAppoint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim nX As Long, nY As Long
    nX = imgAppoint.Left + picTab(0).Left
    nY = imgAppoint.Top + imgAppoint.Height + picTab(0).Top
    frmDetails.PopupMenu frmPrefs.titPicMenu, , nX, nY, frmPrefs.titChangeImg
End Sub

Private Sub imgAppoint_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ERR
    imgAppoint.Picture = LoadPicture(Data.Files(1))
    SetItemData Me.Tag, "ItemImage", Data.Files(1)
    Exit Sub
ERR:
    MsgBox "Error: " & ERR.Number & " - " & ERR.Description & vbCrLf & vbCrLf & _
    "Use the File Name textbox in the image chooser instead.", vbCritical
End Sub


Private Sub lblItemName2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblItemName2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Public Function ClearTriangle()
    On Error Resume Next
    If Val(Me.Tag) > 0 Then
        frmMain.imgTriangle(Val(Me.Tag)).Picture = frmMain.imgTriangle(0).Picture
    Else
        frmMain.imgTriangle(Val(Me.Tag)).Picture = frmMain.imgTriangle(1).Picture
    End If
End Function

Public Function AnimateForm()
    On Error Resume Next
    Dim I As Long
    Me.Enabled = False
    For I = frmMain.Left To frmMain.Left - Me.Width - 60 Step -180
        Me.Left = I
        Me.Width = frmMain.Left - Me.Left - 60
        DoEvents
    Next
    Me.Width = 5775
    Me.Left = frmMain.Left - Me.Width - 60
    Me.Enabled = True
End Function

Private Sub LstTab_Click()
    On Error Resume Next
    Dim I As Integer, K As Integer
    K = LstTab.ListIndex
    For I = 0 To picTab.Count - 1 Step 1
        If I <> K Then
            picTab(I).Visible = False
        Else
            picTab(I).Visible = True
            picTab(I).SetFocus
        End If
    Next
End Sub

Private Sub txtAgenda_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim I As String
    If KeyAscii = vbKeyReturn Then
        I = GetSet("SaveFields1", "1")
        If I = "0" Then
            btnApply_Click
            InfoEdited = False
        ElseIf I = "1" Then
            InfoEdited = True
        ElseIf I = "2" Then
            InfoEdited = False
        End If
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    txtAgenda_KeyPress KeyAscii
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    txtAgenda_KeyPress KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    txtAgenda_KeyPress KeyAscii
End Sub
