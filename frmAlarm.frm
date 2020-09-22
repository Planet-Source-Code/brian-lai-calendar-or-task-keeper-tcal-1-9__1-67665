VERSION 5.00
Begin VB.Form frmAlarm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Name"
   ClientHeight    =   4575
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
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "Alarm Time:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblItemData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Time"
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
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label btnDelItem 
      BackStyle       =   0  '³z©ú
      Caption         =   "Delete Event"
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
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblItemData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblItemData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Time"
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
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblItemData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Name"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblItemName 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "Alarm!"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Click to edit this field"
      Top             =   60
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "Details:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "Time:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "Item Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long

Private Sub btnDelItem_Click()
On Error Resume Next
If MsgBox("Are you sure you want to delete this entry?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
frmMain.DeleteItem Val(Me.Tag)
frmMain.FetchItem Val(Me.Tag)
Unload frmDetails
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
Shape1(0).Width = Me.ScaleWidth
Shape1(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
'DockingTerminate Me
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblItemData_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblItemData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblItemName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblItemName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    sndPlaySound GetSet("Alarm Sound", "C:\WINDOWS\Media\ding.wav"), &H1
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim I As Integer
If GetSet("Transparent", "1") = "1" Then MakeTransparent Me.hWnd, GetSet("Transparent Value", "200")
If GetSet("OnTop", "0") = "1" Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
If GetSet("Black", "0") = "1" Then
    Me.BackColor = RGB(0, 0, 0)
    lblItemName.ForeColor = RGB(255, 255, 255)
    For I = 0 To 2 Step 1
        Label1(I).ForeColor = RGB(255, 255, 255)
        lblItemData(I).ForeColor = RGB(255, 255, 255)
    Next
End If
DropShadow Me.hWnd
'DockingStart Me
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 1
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

