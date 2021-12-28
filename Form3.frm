VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deðerler"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim ilk_, son_ As Integer

If hangisi = True Then
 ilk_ = 0
 son_ = 9999
 List1.ForeColor = vbBlack
Else
 ilk_ = 10000
 son_ = 10199
 List1.ForeColor = vbRed
End If

Me.Caption = ilk_ & " ile " & son_ & " arasý adýmlar için"

For i = ilk_ To son_
 List1.AddItem ("phi(" & i & ")=" & phi(i))
Next
End Sub

