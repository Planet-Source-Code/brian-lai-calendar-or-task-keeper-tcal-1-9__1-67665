VERSION 5.00
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "TCal - Preferences"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   0
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Open Details Window when a field is edited"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Default Saving Action"
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   3615
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   3375
            TabIndex        =   40
            Top             =   240
            Width           =   3375
            Begin VB.OptionButton OptTaskEdited2 
               Caption         =   "Save"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   3375
            End
            Begin VB.OptionButton OptTaskEdited2 
               Caption         =   "Discard Changes"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   41
               Top             =   360
               Width           =   3375
            End
         End
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Move Main Window when resized"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Reposition Details pane when Main window moves"
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.ListBox LstTab 
      Height          =   2820
      IntegralHeight  =   0   'False
      ItemData        =   "frmPrefs.frx":000C
      Left            =   120
      List            =   "frmPrefs.frx":0025
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   2
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.Frame frmAutoDel 
         Caption         =   "Completed Tasks"
         Height          =   1335
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   3735
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   3495
            TabIndex        =   26
            Top             =   240
            Width           =   3495
            Begin VB.OptionButton OptTaskdone 
               Caption         =   "Do Nothing"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Width           =   3495
            End
            Begin VB.OptionButton OptTaskdone 
               Caption         =   "Prompt to delete"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   28
               Top             =   360
               Width           =   3495
            End
            Begin VB.OptionButton OptTaskdone 
               Caption         =   "Automatic Delete"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   27
               Top             =   720
               Width           =   3495
            End
         End
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   3
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Allow Resizing"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   3735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Use event image as sticky background"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   3735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Restore Sticky Notes when TCal starts"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   5
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Default position being bottom right"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Transparency"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   3735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Use Black as base colour"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   3735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Shadows"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Transparency (0-255)"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1560
         Width           =   2775
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   4
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   8
      Top             =   120
      Width           =   3975
      Begin VB.Timer tmrSnd 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1080
         Top             =   1920
      End
      Begin VB.CommandButton btnTestMedia 
         Caption         =   "Test"
         Height          =   375
         Left            =   3120
         TabIndex        =   43
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Change Alarm Time to Event Time when Event Time changes"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  '¤â°Ê
         TabIndex        =   14
         Text            =   "Argh"
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton btnShellMedia 
         Caption         =   "&Find Media..."
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Alarm Sound: (will replay every 100ms)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Find a media, then drag the media file into the text box above."
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   3690
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   1
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      Begin VB.Frame Frame1 
         Caption         =   "When data is edited"
         Height          =   1695
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   3615
         Begin VB.PictureBox PicContainer 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   3375
            TabIndex        =   32
            Top             =   240
            Width           =   3375
            Begin VB.OptionButton OptTaskEdited 
               Caption         =   "Discard Changes"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   35
               ToolTipText     =   "Data is not saved unless the user clicks ""Apply Changes"" manually"
               Top             =   720
               Width           =   3375
            End
            Begin VB.OptionButton OptTaskEdited 
               Caption         =   "Ask to Save"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   34
               ToolTipText     =   "Data is saved when the user tries to close the form"
               Top             =   360
               Width           =   3375
            End
            Begin VB.OptionButton OptTaskEdited 
               Caption         =   "Save Immediately"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   33
               ToolTipText     =   "Data is saved whenever the enter key is pressed"
               Top             =   0
               Width           =   3375
            End
            Begin VB.Label Label3 
               Alignment       =   2  '¸m¤¤¹ï»ô
               Caption         =   "Hover on the radio buttons for details"
               Height          =   255
               Left            =   0
               TabIndex        =   44
               Top             =   1080
               Width           =   3375
            End
         End
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Animate Details"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2775
      Index           =   6
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   3975
      TabIndex        =   10
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton btnWriteXPVS 
         Caption         =   "Use XP Visual Styles"
         Height          =   375
         Left            =   0
         TabIndex        =   45
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Start TCal Minimized (Main Window Only)"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton btnAbout 
         Caption         =   "About TCal"
         Height          =   375
         Left            =   2280
         TabIndex        =   30
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox ChkOpt 
         Caption         =   "Show Window on the task bar (Requires Restart)"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Menu titMainMenu 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu titAddItem 
         Caption         =   "Add Item"
      End
      Begin VB.Menu titSort 
         Caption         =   "Sort Items"
      End
      Begin VB.Menu jkldsjlldsfsd2 
         Caption         =   "-"
      End
      Begin VB.Menu titDeleteCompleted 
         Caption         =   "Deleted Completed Items"
      End
      Begin VB.Menu titDeleteAll 
         Caption         =   "Delete All Items"
      End
      Begin VB.Menu jkldsjlldsfsd 
         Caption         =   "-"
      End
      Begin VB.Menu titAoT 
         Caption         =   "Always on Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu titPrefs 
         Caption         =   "Preferences"
      End
      Begin VB.Menu titAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu jkldsjlldsfss 
         Caption         =   "-"
      End
      Begin VB.Menu titMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu titExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu titPicMenu 
      Caption         =   "PicMenu"
      Visible         =   0   'False
      Begin VB.Menu titChangeImg 
         Caption         =   "Change"
      End
      Begin VB.Menu titDelImg 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu titTaskmenu 
      Caption         =   "TaskMenu"
      Visible         =   0   'False
      Begin VB.Menu titTaskDetails 
         Caption         =   "Details..."
      End
      Begin VB.Menu titDeleteTask 
         Caption         =   "Delete this item"
      End
   End
   Begin VB.Menu titStickyMenu 
      Caption         =   "StickyMenu"
      Visible         =   0   'False
      Begin VB.Menu titViewEvent 
         Caption         =   "View this Event"
      End
      Begin VB.Menu titDelEvent 
         Caption         =   "Delete this Event"
      End
      Begin VB.Menu titChooseClr 
         Caption         =   "Choose Sticky Colour"
         Begin VB.Menu titClr 
            Caption         =   "Default"
            Index           =   0
         End
         Begin VB.Menu titClr 
            Caption         =   "White"
            Index           =   1
         End
         Begin VB.Menu titClr 
            Caption         =   "Yellow"
            Index           =   2
         End
         Begin VB.Menu titClr 
            Caption         =   "Red"
            Index           =   3
         End
         Begin VB.Menu titClr 
            Caption         =   "Blue"
            Index           =   4
         End
         Begin VB.Menu titClr 
            Caption         =   "Green"
            Index           =   5
         End
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAbout_Click()
    On Error Resume Next
    titAbout_Click
End Sub

Private Sub btnOK_Click()
    On Error Resume Next
    Dim I As Integer
    SaveSet "Shadow", ChkOpt(0).Value
    SaveSet "Black", ChkOpt(1).Value
    SaveSet "Transparent", ChkOpt(2).Value
    SaveSet "PosDefault", ChkOpt(3).Value
    SaveSet "EventAlarm", ChkOpt(4).Value
    SaveSet "Animate", ChkOpt(5).Value
    SaveSet "StickyBack", ChkOpt(6).Value
    SaveSet "ShowInTaskbar", ChkOpt(7).Value
    SaveSet "RestoreSticky", ChkOpt(8).Value
    SaveSet "MagneticDetails", ChkOpt(9).Value
    SaveSet "StickyResize", ChkOpt(10).Value
    SaveSet "StartMinimize", ChkOpt(11).Value
    SaveSet "ResetMain", ChkOpt(12).Value
    SaveSet "OpenDetails", ChkOpt(13).Value
    SaveSet "Transparent Value", Text1.Text
    SaveSet "Alarm Sound", Text2.Text
    For I = 0 To 2 Step 1
        If OptTaskdone(I).Value = True Then
            SaveSet "AutoDel", Str(I)
        End If
        If OptTaskEdited(I).Value = True Then
            SaveSet "SaveFields1", Str(I)
        End If
        If I <= 1 Then
            If OptTaskEdited2(I).Value = True Then
                SaveSet "SaveFields2", Str(I)
            End If
        End If
    Next
    tmrSnd.Enabled = False
    Unload Me
End Sub

Private Sub btnShellMedia_Click()
    On Error Resume Next
    Shell "explorer C:\WINDOWS\Media", vbNormalFocus
End Sub

Private Sub btnTestMedia_Click()
    On Error Resume Next
    tmrSnd.Enabled = Not tmrSnd.Enabled
End Sub

Private Sub btnWriteXPVS_Click()
    On Error Resume Next
    If MsgBox("This function will write the manifest file again to show the Windows XP Visual Styles if applicable.", _
    vbYesNo + vbQuestion) = vbNo Then Exit Sub
    XPVB True
    MsgBox "Manifest Written. Please restart " & App.ProductName & ".", vbInformation
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim OptBuffer As Integer
    If GetSet("OnTop", "0") = "1" Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
    ChkOpt(0).Value = GetSet("Shadow", "1")
    ChkOpt(1).Value = GetSet("Black", "0")
    ChkOpt(2).Value = GetSet("Transparent", "1")
    ChkOpt(3).Value = GetSet("PosDefault", "1")
    ChkOpt(4).Value = GetSet("EventAlarm", "1")
    ChkOpt(5).Value = GetSet("Animate", "0")
    ChkOpt(6).Value = GetSet("StickyBack", "0")
    ChkOpt(7).Value = GetSet("ShowInTaskbar", "1")
    ChkOpt(8).Value = GetSet("RestoreSticky", "1")
    ChkOpt(9).Value = GetSet("MagneticDetails", "1")
    ChkOpt(10).Value = GetSet("StickyResize", "1")
    ChkOpt(11).Value = GetSet("StartMinimize", "0")
    ChkOpt(12).Value = GetSet("ResetMain", "1")
    ChkOpt(13).Value = GetSet("OpenDetails", "1")
    OptBuffer = Val(GetSet("AutoDel", "1"))
    OptTaskdone(OptBuffer).Value = True
    OptBuffer = Val(GetSet("SaveFields1", "1"))
    OptTaskEdited(OptBuffer).Value = True
    OptBuffer = Val(GetSet("SaveFields2", "1"))
    OptTaskEdited2(OptBuffer).Value = True
    Text1.Text = GetSet("Transparent Value", "200")
    Text2.Text = GetSet("Alarm Sound", "C:\WINDOWS\Media\ding.wav")
    titAoT.Checked = (GetSet("OnTop", "0") = "1")
    'DockingStart Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'DockingTerminate Me
End Sub

Private Sub LstTab_Click()
    On Error Resume Next
    picTabSwitch(LstTab.ListIndex).ZOrder 0
End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Data.GetFormat(vbCFFiles) = False Then Exit Sub
    If Len(Data.Files(1)) > 0 Then
        Text2.Text = Data.Files(1)
    End If
End Sub

Private Sub titAbout_Click()
    On Error Resume Next
    MsgBox App.ProductName & " " & MyVer & ", all rights reserved by Thinc." & vbCrLf & _
    vbCrLf & "Made by Brian Lai" & vbCrLf & vbCrLf & "http://thinc.no-ip.info" & vbCrLf & vbCrLf & _
    "Help from: 1997 SoftCircuits Programming", vbInformation
End Sub

Private Sub titAddItem_Click()
    On Error Resume Next
    Dim I As Integer
    I = frmMain.NextEmptyItem(True)
    If I = -1 Then
        MsgBox "There aren't enough free data slots to add an item.", vbInformation
    Else
        Call frmMain.lblAgenda_MouseDown(I, 1, 0, 0, 0)
    End If
End Sub

Private Sub titAoT_Click()
    On Error Resume Next
    titAoT.Checked = Not titAoT.Checked
    SetWindowPos frmMain.hWnd, CInt(Not titAoT.Checked) - 1, 0, 0, 0, 0, 3
    SaveSet "OnTop", Abs(CInt(titAoT.Checked))
End Sub

Private Sub titChangeImg_Click()
    On Error Resume Next
    frmImage.Show 1
End Sub

Private Sub titClr_Click(Index As Integer)
    On Error Resume Next
    Dim K As String
    With Screen.ActiveForm
        Select Case Index
            Case 0
                SetItemData .txtTag.Text, "StickyClr", "Default"
            Case 1
                SetItemData .txtTag.Text, "StickyClr", "16777215"
            Case 2
                SetItemData .txtTag.Text, "StickyClr", "8454143"
            Case 3
                SetItemData .txtTag.Text, "StickyClr", "8421631"
            Case 4
                SetItemData .txtTag.Text, "StickyClr", "16754253"
            Case 5
                SetItemData .txtTag.Text, "StickyClr", "11599792"
            End Select
        K = GetItemData(Val(.txtTag.Text), "StickyClr", "Default")
        If IsNumeric(K) = True Then .BackColor = K
    End With
End Sub

Private Sub titDeleteAll_Click()
    On Error Resume Next
    Dim I As Integer
    If MsgBox("Are you sure you want to delete all items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    For I = 0 To 11 Step 1
        frmMain.DeleteItem I
    Next
    frmMain.FetchAllItems
End Sub

Private Sub titDeleteCompleted_Click()
    On Error Resume Next
    Dim I As Integer
    If MsgBox("Are you sure you want to delete all completed items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    For I = 0 To 11 Step 1
        If GetItemData(I, "ItemDone") = "1" Then frmMain.DeleteItem I
    Next
    frmMain.SortData
    frmMain.FetchAllItems
End Sub

Private Sub titDeleteTask_Click()
    On Error Resume Next
    If MsgBox("Are you sure you want to delete this entry?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    frmMain.DeleteItem Val(titDeleteTask.Tag)
    frmMain.FetchItem Val(titDeleteTask.Tag)
End Sub

Private Sub titDelEvent_Click()
    On Error Resume Next
    Call frmMain.imgTriangle_Click(Val(titTaskDetails.Tag))
    frmDetails.btnDelItem_Click
End Sub

Private Sub titDelImg_Click()
    On Error Resume Next
    frmDetails.imgAppoint.Picture = frmDetails.imgNothing.Picture
    SetItemData frmDetails.Tag, "ItemImage", ""
End Sub

Private Sub titExit_Click()
    On Error Resume Next
    End
End Sub

Private Sub titMinimize_Click()
    On Error Resume Next
'    Dim Dx As Form
'    For Each Dx In Forms
'        If Dx.Name = "frmMain" Then
'            Dx.WindowState = 1
'        Else
'            Dx.Visible = False
'        End If
'    Next
    frmMain.WindowState = 1
End Sub

Private Sub titPrefs_Click()
    On Error Resume Next
    frmPrefs.Show 1
End Sub

Private Sub titSort_Click()
    On Error Resume Next
    Call frmMain.SortData
End Sub

Private Sub titTaskDetails_Click()
    On Error Resume Next
    Call frmMain.imgTriangle_Click(Val(titTaskDetails.Tag))
End Sub

Private Sub titViewEvent_Click()
    On Error Resume Next
    Call frmMain.OpenDetails(Val(Screen.ActiveForm.txtTag.Text))
End Sub

Private Sub tmrSnd_Timer()
    On Error Resume Next
    sndPlaySound Text2.Text, &H1
End Sub
