VERSION 5.00
Begin VB.Form FrmAboutSoftWare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于本软件"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmAboutSoftWare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = "提取yande地址" & vbCrLf & "版本：V" & App.Major & "." & App.Minor & vbCrLf & "源代码提供："
End Sub

