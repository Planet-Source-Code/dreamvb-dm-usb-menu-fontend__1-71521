VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FAC38F&
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   Icon            =   "fmrmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin Project1.dCDMenuList dCDMenuList1 
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   1350
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2566
      SelectBackColor =   8421504
      BackColor       =   14737632
   End
   Begin Project1.dTabStrip dTabStrip1 
      Height          =   345
      Left            =   -15
      TabIndex        =   5
      Top             =   1005
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   609
      BackColor       =   16434063
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabBackColor    =   16434063
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00D8A16D&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8325
      TabIndex        =   2
      Top             =   5085
      Width           =   8325
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H00D7A06C&
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6090
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Install"
         Top             =   90
         Width           =   1020
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00D7A06C&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7200
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Install"
         Top             =   105
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFE1CF&
         X1              =   0
         X2              =   315
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   8325
      TabIndex        =   0
      Top             =   0
      Width           =   8325
      Begin VB.Image ImgLogo 
         Height          =   720
         Left            =   120
         Top             =   90
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFE1CF&
         X1              =   0
         X2              =   315
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "#Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   885
         TabIndex        =   1
         Top             =   150
         Width           =   6810
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private mHeight As Long
Private mWidth As Long
Private mTabKey As String
Private MyIni As dINIFile
Private mAppPath As String

Private Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Private Sub RunApp(iHwnd As Long, OpenOp As String, FileName As String)
Dim Ret As Long
    Ret = ShellExecute(iHwnd, OpenOp, FileName, "", "", 1)
    If (Ret = 2) Then
        MsgBox "There was an error opening the file:" & vbCrLf & vbCrLf & FileName, vbExclamation, "File Not Found"
    End If
End Sub

Private Sub SetupUI()
    'Setup the logo
    ImgLogo.Picture = LoadPicture(mAppPath & MyIni.ReadValue("Main", "Logo", vbNullString))
    'Setup First title
    frmmain.Caption = MyIni.ReadValue("Main", "Title1", vbNullString)
    'Setup second title
    lblTitle.Caption = MyIni.ReadValue("Main", "Title2", vbNullString)
    'Title forecolor
    lblTitle.ForeColor = Val(MyIni.ReadValue("Color", "Title1", &HFFFFFF))
    'Header backcolor
    pTop.BackColor = Val(MyIni.ReadValue("Color", "Header", &HD8A16D))
    'Footer backcolor
    pBottom.BackColor = Val(MyIni.ReadValue("Color", "Footer", &HD8A16D))
    'Set backcolor
    frmmain.BackColor = Val(MyIni.ReadValue("Color", "BackColor", &HFAC38F))
    'Set tab backcolor
    dTabStrip1.BackColor = frmmain.BackColor
    'Set TabBackColor
    dTabStrip1.TabBackColor = Val(MyIni.ReadValue("Color", "TabBackColor", &HFAC38F))
    'Set Tab Selected Color
    dTabStrip1.TabSelectedColor = Val(MyIni.ReadValue("Color", "TabSelectColor", &HFFFFFF))
    'Tab ForeColor
    dTabStrip1.TextColor = Val(MyIni.ReadValue("Color", "TabForeColor", &H0&))
    'Set button colors
    cmdAbout.BackColor = Val(MyIni.ReadValue("Color", "Buttons", &HD7A06C))
    cmdExit.BackColor = cmdAbout.BackColor
    'Set ItemList BackColor
    dCDMenuList1.BackColor = Val(MyIni.ReadValue("Color", "ListBackColor", &HE0E0E0))
    'Set List Forecolor
    dCDMenuList1.ForeColor = Val(MyIni.ReadValue("Color", "ListForeColor", &H0&))
    'Set List selected color
    dCDMenuList1.SelectForeColor = Val(MyIni.ReadValue("Color", "ListSelectForeColor", &HFFFFFF))
    'Set List Item BackColor
    dCDMenuList1.SelectBackColor = Val(MyIni.ReadValue("Color", "ListSelectBackColor", &H808080))
End Sub

Private Sub FillItems(ByVal sGroup As String)
Dim ItemCnt As Integer
Dim Count As Integer
Dim iAppCaption As String
Dim iDesc As String
Dim iIconFile As String
    
    'Get the Number of items
    ItemCnt = Val(MyIni.ReadValue(sGroup, "Apps", 0))
    
    For Count = 1 To ItemCnt
        'Get Caption
        iAppCaption = Trim$(MyIni.ReadValue(sGroup, "App" & Count, vbNullString))
        'Get Description
        iDesc = Trim$(MyIni.ReadValue(sGroup, "Description" & Count, vbNullString))
        'Get Icon Filename
        iIconFile = mAppPath & Trim$(MyIni.ReadValue(sGroup, "Icon" & Count, vbNullString))
        'Add the item information
        dCDMenuList1.AddItem iAppCaption, iDesc, vbNullString, LoadPicture(iIconFile)
    Next Count
    
    'Select the first item in the list.
    dCDMenuList1.ListIndex = 1
End Sub

Private Sub AddTabs()
Dim gCount As Integer
Dim Count As Integer
    'Get the number of tabs to add
    gCount = Val(MyIni.ReadValue("Main", "Groups", 0))
    'Setup the tabs
    For Count = 1 To gCount
        dTabStrip1.AddTab MyIni.ReadValue("Group" & Count, "Name"), "Group" & Count
    Next Count
    'Select the first tab
    dTabStrip1.TabSelect = 1
End Sub

Private Function FixPath(ByVal lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Private Sub cmdAbout_Click()
Dim StrA As String
    'Get the About string
    StrA = MyIni.ReadValue("Main", "AboutStr", vbNullString)
    'Replace $ with a new line
    StrA = Replace(StrA, "$", vbCrLf, , , vbBinaryCompare)
    MsgBox StrA, vbInformation, "About"
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit now.", vbYesNo Or vbQuestion, "Exit") = vbYes Then
        Unload frmmain
    End If
End Sub

Private Sub dCDMenuList1_DblClick()
Dim lFile As String
    'Get Command
    lFile = mAppPath & MyIni.ReadValue(mTabKey, "Exec" & dCDMenuList1.ListIndex, vbNullString)
    'Execute the command
    Call RunApp(frmmain.hwnd, "open", lFile)
End Sub

Private Sub dTabStrip1_TabChange(Index As Integer, Key As String, Caption As String)
    'Store Tab Key
    mTabKey = Key
    'Clear the Items List
    Call dCDMenuList1.Clear
    'Fill items list
    Call FillItems(mTabKey)
End Sub

Private Sub Form_Load()
    mAppPath = FixPath(App.Path)
    'Store old form height and width
    mHeight = frmmain.Height
    mWidth = frmmain.Width
    'Setup ini Control
    Set MyIni = New dINIFile
    'Set the ini file to read
    MyIni.FileName = mAppPath & "menu.ini"
    'Check that the main ini file is found
    If Not FindFile(MyIni.FileName) Then
        MsgBox "Cannot find menu.ini", vbExclamation, "Fil Not Found"
        Unload frmmain
    End If
    
    'Set Up UI
    Call SetupUI
    'Add the Tabs
    Call AddTabs
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Resizeing Codes
    Line1.X2 = ScaleWidth
    Line2.X2 = ScaleWidth
    dTabStrip1.Width = (frmmain.ScaleWidth - dTabStrip1.Left) - 15
    dCDMenuList1.Height = (frmmain.ScaleHeight - pBottom.Height - dCDMenuList1.Top)
    dCDMenuList1.Width = (frmmain.ScaleWidth - dCDMenuList1.Left) - 1
    frmmain.Height = mHeight
    frmmain.Width = mWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MyIni = Nothing
    Set frmmain = Nothing
End Sub
