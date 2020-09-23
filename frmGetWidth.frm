VERSION 5.00
Begin VB.Form frmGetWidth 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Enter New Width"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   Icon            =   "frmGetWidth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmGetWidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' All that this form does is give a box asking for the new width of an arc
'

Private Storage As Integer

Public Function GetNewWidth(ByVal CurrentWidth As Integer) As Integer
Text1.Text = Trim(Str(CurrentWidth))
Me.Show 1
GetNewWidth = Storage
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

' Validate returned data
If Not IsNumeric(Text1.Text) Then MsgBox "Please enter a number between 1 and 255", vbInformation, "Error": Exit Sub
If Val(Text1.Text) < 1 Or Val(Text1.Text) > 255 Then MsgBox "Please enter a number between 1 and 255", vbInformation, "Error": Exit Sub
If InStr(Text1.Text, ".") <> 0 Then MsgBox "Please enter a whole number", vbInformation, "Error": Exit Sub
Storage = Val(Text1.Text)
Unload Me

End Sub
