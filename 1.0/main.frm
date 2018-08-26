VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提取yande地址（版本：V1.0    作者：小叶Little_Ye）"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   12585
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Init 
      Caption         =   "重置"
      Height          =   255
      Left            =   11160
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "进度"
      Height          =   1095
      Left            =   6360
      TabIndex        =   13
      Top             =   480
      Width           =   6135
      Begin ComctlLib.ProgressBar Progress 
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1296
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "目标URL"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   6135
      Begin VB.TextBox UrlView 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.TextBox Argument 
      Height          =   270
      Index           =   1
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Argument 
      Height          =   270
      Index           =   0
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton GetResult 
      Caption         =   "提取"
      Height          =   255
      Left            =   10440
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Result 
      Height          =   3855
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Caption         =   "网页源代码（由于文本框字数限制只能显示部分）"
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6135
      Begin VB.TextBox Source 
         Height          =   3855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton Work 
      Caption         =   "获取"
      Height          =   255
      Left            =   9720
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "提取结果"
      Height          =   4215
      Left            =   6360
      TabIndex        =   5
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "页数："
      Height          =   255
      Left            =   8280
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "搜索关键字："
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "网址域名：https://yande.re/post?"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UrlSource As String
'Dim page As Integer
'Dim UrlArrayLen As Integer
'Dim UrlArray() As String

Private Sub GetResult_Click()
    Dim matchesUrl As MatchCollection
    Dim matchesPic As MatchCollection
    Dim matchUrl As Match
    Dim matchPic As Match
    Dim patUrl As String
    Dim patPic As String
    Dim patPicBackup As String
    Dim urlSourceText As String
    
    patUrl = "https://yande.re/post/show/[\d]+"
    patPic = "https://files.yande.re/image/[\S]+.png"",""i"
    patPicBackup = "https://files.yande.re/image/[\S]+.jpg"",""i"
    
    Set matchesUrl = AnalyzeRegExp(patUrl, UrlSource)
    Progress.Value = 0
    Progress.Max = matchesUrl.Count
    For Each matchUrl In matchesUrl
        urlSourceText = GetUrl(matchUrl.Value)
        Set matchesPic = AnalyzeRegExp(patPic, urlSourceText)
        If matchesPic.Count = 0 Then
            Set matchesPic = AnalyzeRegExp(patPicBackup, urlSourceText)
        End If
        For Each matchPic In matchesPic
            Result.Text = Result.Text + Left(matchPic.Value, Len(matchPic.Value) - 4) & vbCrLf
        Next matchPic
        Progress.Value = Progress.Value + 1
    Next matchUrl
End Sub

Private Sub Init_Click()
    Argument(0).Text = ""
    Argument(1).Text = ""
    Progress.Value = 0
    Progress.Max = 100
    UrlView.Text = ""
    Source.Text = ""
    Result.Text = ""
End Sub

Private Sub Work_Click()
    Dim Url As String
    Url = "https://yande.re/post?"
    arg_cnt = 0
    If Argument(0) <> "" Then
        arg_cnt = arg_cnt + 1
    End If
    If Argument(1) <> "" Then
        arg_cnt = arg_cnt + 1
    End If
    If Argument(0) <> "" Then
        Url = Url + "tags=" + Argument(0).Text
    End If
    If arg_cnt = 2 Then
        Url = Url + "&"
    End If
    If Argument(1) <> "" Then
        Url = Url + "page=" + Argument(1).Text
    End If
    UrlView.Text = Url
    UrlSource = GetUrl(UrlView.Text)
    Source.Text = UrlSource
'    Dim UrlLen As Integer
'    UrlLen = StrLen(GetUrl(UrlView.Text))
'    UrlArrayLen = -CInt(-UrlLen / 65535#)
'    ReDim UrlArray(UrlArrayLen) As String
'    For idx = 0 To UrlArrayLen - 1
'        UrlArray(idx) = Mid(Url, 65535 * idx, Min(65535, UrlLen - 65535 * idx))
'    Next idx
'    Source.Text = UrlArray(page)
End Sub
