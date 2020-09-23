VERSION 5.00
Begin VB.Form frmSelect 
   Caption         =   "Select Data to Gather"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Deselect All"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox lstData 
      Height          =   2985
      ItemData        =   "frmSelect.frx":0000
      Left            =   120
      List            =   "frmSelect.frx":0002
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAll_Click()
Dim i As Integer
For i = 0 To lstData.ListCount - 1
    lstData.Selected(i) = True
Next i

End Sub



Private Sub cmdNone_Click()
Dim i As Integer
For i = 0 To lstData.ListCount - 1
    lstData.Selected(i) = False
Next i
End Sub

Private Sub cmdOK_Click()
Me.Hide
If lstData.SelCount > 0 Then
    Form1.cmdStart.Enabled = True
Else
    Form1.cmdStart.Enabled = False
End If
End Sub
