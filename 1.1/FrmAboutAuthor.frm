VERSION 5.00
Begin VB.Form FrmAboutAuthor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于作者"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmAboutAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = "提取yande地址" & vbCrLf & "作者：小叶Little_Ye" & vbCrLf & "工作邮箱：littleye233@foxmail.com"
End Sub

