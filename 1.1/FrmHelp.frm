VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提取yande源代码 使用说明"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = "提取yande源代码" & vbCrLf & "在主界面指定位置输入欲搜索的关键词和页数，之后可进行其他操作。" & vbCrLf & "【获取】可以根据搜索参数获取对应页面的源代码（测试用）。" & vbCrLf & "【提取】根据页面源代码提取出png或jpg图片的地址，可以根据这些地址用批量下载工具下载图片。" & vbCrLf & "注：请保证输入的搜索参数合法，且在进行【提取】前请进行【获取】操作！"
End Sub
