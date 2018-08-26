VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6465
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   12585
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "图片地址列表"
      Height          =   4215
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   6135
      Begin ComctlLib.ListView PictureUrlView 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6800
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin ComctlLib.ProgressBar ProgressTotal 
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6090
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "就绪"
            TextSave        =   "就绪"
            Key             =   "NowStatus"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "进度："
            TextSave        =   "进度："
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   ""
            Key             =   "NowProgress"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   ""
            Key             =   "TotalProgress"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "进度"
      Height          =   1095
      Left            =   6360
      TabIndex        =   8
      Top             =   480
      Width           =   6135
      Begin ComctlLib.ProgressBar Progress 
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "分进度"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "总进度"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "目标URL"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6135
      Begin VB.TextBox UrlView 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.TextBox Argument 
      Height          =   270
      Index           =   1
      Left            =   9360
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Argument 
      Height          =   270
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox Result 
      Height          =   3855
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      Caption         =   "提取结果"
      Height          =   4215
      Left            =   6360
      TabIndex        =   1
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "搜索关键字："
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "网址域名：https://yande.re/post?"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "页数生成器："
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu File 
      Caption         =   "文件"
      Begin VB.Menu SaveAuto 
         Caption         =   "保存至 output.txt"
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu Operation 
      Caption         =   "操作"
      Begin VB.Menu Get 
         Caption         =   "获取"
      End
      Begin VB.Menu Fliter 
         Caption         =   "提取"
      End
      Begin VB.Menu Initialization 
         Caption         =   "重置"
      End
   End
   Begin VB.Menu About 
      Caption         =   "关于"
      Begin VB.Menu Help 
         Caption         =   "帮助"
      End
      Begin VB.Menu AboutSoftware 
         Caption         =   "关于本软件"
      End
      Begin VB.Menu AboutAuthor 
         Caption         =   "关于作者"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UrlSource As String
Dim UrlSources() As String

Private Sub AboutAuthor_Click()
    FrmAboutAuthor.Show
End Sub

Private Sub AboutSoftware_Click()
    FrmAboutSoftWare.Show
End Sub

Private Sub Exit_Click()
    Dim exitReturn As Integer
    exitReturn = MsgBox("您确定退出吗？", 33, "提示")
    If exitReturn = 1 Then
        Unload Me
    End If
End Sub

Private Sub Fliter_Click()
    Status.Panels(1).Text = "提取目标地址"
    Dim matchesUrl As MatchCollection
    Dim matchesPic As MatchCollection
    Dim matchUrl As Match
    Dim matchPic As Match
    Dim patUrl As String
    Dim patPic As String
    Dim patPicBackup As String
    Dim urlSourceText As String
'    Dim UrlSource
    
    patUrl = "https://yande.re/post/show/[\d]+"
    patPic = "https://files.yande.re/image/[\S]+.png"",""i"
    patPicBackup = "https://files.yande.re/image/[\S]+.jpg"",""i"
'    ProgressTotal.Value = 0
'    ProgressTotal.Max = UBound(UrlSources) + 1
    
'    For Each UrlSource In UrlSources
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
'        ProgressTotal.Value = ProgressTotal.Value + 1
'    Next UrlSource
    Status.Panels(1).Text = "就绪"
End Sub

Private Sub Form_Load()
    ReDim UrlSources(0)
    Status.Panels(1).Text = "就绪"
    FrmMain.Caption = "提取yande地址（版本：V" & App.Major & "." & App.Minor & "    作者：小叶Little_Ye）"
End Sub

Private Sub Get_Click()
    Status.Panels(1).Text = "获取源代码"
    Dim page
    Dim pages
    Dim Url As String
    Dim complete_url As String
    complete_url = ""
    pages = PageGenerate(Argument(1).Text)
    Progress.Value = 0
    Progress.Max = UBound(pages) + 1
    For Each page In pages
        Url = "https://yande.re/post?tags=" & Argument(0).Text & "&page=" & page
        complete_url = complete_url + Url & vbCrLf
        UrlView.Text = complete_url
        UrlSource = UrlSource + GetUrl(Url)
'        UrlSources(UBound(UrlSources)) = GetUrl(Url)
'        ReDim UrlSources(UBound(UrlSources) + 1)
        Progress.Value = Progress.Value + 1
    Next page
    Status.Panels(1).Text = "就绪"
End Sub

Private Sub Help_Click()
    FrmHelp.Show
End Sub

Private Sub Initialization_Click()
    Argument(0).Text = ""
    Argument(1).Text = ""
    Progress.Value = 0
    Progress.Max = 100
    UrlView.Text = ""
    Result.Text = ""
    Status.Panels(1).Text = "就绪"
End Sub

Private Sub SaveAuto_Click()
    Dim SaveAutoReturn As Integer
    Dim SaveAutoReturn1 As Integer
    Status.Panels(1).Text = "保存"
    If Dir(App.path & "\output.txt") <> "" Then
        SaveAutoReturn = MsgBox("检测到工作目录下存在output.txt，确定覆盖吗？", 33, "提示")
        If SaveAutoReturn <> 1 Then
            Status.Panels(1).Text = "就绪"
            Exit Sub
        End If
    End If
    If Result.Text = "" Then
        SaveAutoReturn1 = MsgBox("检测到结果栏无文本，确定覆盖吗？", 33, "提示")
        If SaveAutoReturn1 <> 1 Then
            Status.Panels(1).Text = "就绪"
            Exit Sub
        End If
    End If
    SaveFileToLocal App.path & "\output.txt", Mid(Result.Text, 1, Len(Result.Text) - 2)
    MsgBox "保存成功！", 64, "提示"
    Status.Panels(1).Text = "就绪"
End Sub

