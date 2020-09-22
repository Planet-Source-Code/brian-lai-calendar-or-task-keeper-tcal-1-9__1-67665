VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "TCal - Choose Image"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  '¤â°Ê
   ScaleHeight     =   1935
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.TextBox txtImgName 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '³z©ú
      Caption         =   "File name:"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "To choose an image, Select an image from Windows Explorer and drag the image into the square."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4020
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  '¥­­±
      BorderStyle     =   1  '³æ½u©T©w
      Height          =   735
      Left            =   120
      OLEDragMode     =   1  '¦Û°Ê
      OLEDropMode     =   1  '¤â°Ê
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
frmDetails.imgAppoint.Picture = LoadPicture(txtImgName.Text)
SetItemData frmDetails.Tag, "ItemImage", txtImgName.Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
If GetSet("OnTop", "0") = "1" Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
'DockingStart Me
DropShadow Me.hWnd
Me.Move frmDetails.Left + (frmDetails.Width - Me.Width) / 2, frmDetails.Top + (frmDetails.Height - Me.Height) / 2
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
'DockingTerminate Me
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR
Image1.Picture = LoadPicture(Data.Files(1))
Image1.Tag = Data.Files(1)
txtImgName.Text = Image1.Tag
Exit Sub
ERR:
MsgBox "Error: " & ERR.Number & " - " & ERR.Description & vbCrLf & vbCrLf & _
"Use the File Name textbox instead.", vbCritical
End Sub

