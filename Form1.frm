VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "NET NANNY CLONE"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   4515
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "HIDE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2280
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1920
      Top             =   3000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "REMOVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2190
      ItemData        =   "Form1.frx":0442
      Left            =   240
      List            =   "Form1.frx":0444
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "SITES"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "REMOVE SITE"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2565
      TabIndex        =   7
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   2520
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "ADD SITE KEYWORD"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   2535
      TabIndex        =   6
      Top             =   240
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1215
      Left            =   2520
      Top             =   360
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Dim mns As String 'store the site names
Dim ps As Long
Dim formheight As Integer
Dim formwidth As Integer

Private Sub Command1_Click()
'Add web site to list and to file
Dim assi As Integer
Dim i As Integer

On Error Resume Next
assi = Asc(Text1)

'checks whether site added is empty or not
If Text1 <> "" And assi <> 32 Then
For i = List1.ListCount To 0 Step -1
    If List1.List(i) = LCase(Text1) Then
        MsgBox "ALREADY ADDED", vbInformation, "ERROR"
        Exit Sub
    End If
Next i
    Call openfile
    'encryted string added to file
    mns = mns & "#" & ENCR(LCase(Text1)) & "*"
    Call writeINI(mns) 'write to file
    List1.AddItem LCase(Text1)
    Text1.Text = ""
    Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
If Not List1.ListIndex = -1 Then
    Kill App.Path & "\" & "config.ini"
    Call openfile
    mns = Replace(mns, "#" & ENCR(List1.Text) & "*", "")
    Call writeINI(mns)
    List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub Command3_Click()
    'make form invisible
    Form1.Hide
    Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As String
    formwidth = 4635
    formheight = 4215
    Form1.Height = formheight
    Form1.Width = formwidth
    
    'remove from task mgr applications tab
    App.TaskVisible = False
    Call RemoveCtrlAltDel
    
    'checks whether previous instance running or not
    If App.PrevInstance = True Then
        End
    End If
    
    Form1.Hide
    Call openfile
    List1.Clear
    ps = 1
    i = readINI(mns, "#", "*", ps)
    While ps <> 0
    it = DECR(i)
    List1.AddItem it
    i = readINI(mns, "#", "*", ps)
    Wend
    'create value in registry auto start application
    'whenever Windows OS Starts
    createregistrykey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & "NetN Clone", App.Path & "\" & App.EXEName
    createregistrykey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" & "NetN Clone", App.Path & "\" & App.EXEName
End Sub

Private Sub Form_Resize()
    Form1.Height = formheight
    Form1.Width = formwidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'writing string to file during unload
    If mns <> "" Then Call writeINI(mns)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then Call Command2_Click
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Author: Venkatesh K" & vbCrLf & "India", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHide_Click()
    Command3_Click
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call Command1_Click
End Sub

Private Sub Text2_Change()
    'if invisible text2's length = 50 then make it empty
    If Len(Text2) = 100 Then
        Text2.Text = ""
    End If
End Sub

Private Sub timer1_timer()
'checks whether browser title bar text
    Dim frhwnd As Long
    Dim temp As String
    Dim yr As Integer, mn As Integer, dt As Integer
        
    'get active window handle
    frhwnd = GetForegroundWindow()
    temp = Space(255)
    'gets activewindows text
    res = GetWindowText(frhwnd, temp, 255)
    temp = RTrim(temp)
    
    'checks for browser IE or Netscape
    If InStr(1, temp, "Microsoft Internet Explorer", vbTextCompare) Or InStr(1, temp, "Netscape", vbTextCompare) Then
    For i = List1.ListCount - 1 To 0 Step -1
        If InStr(1, temp, List1.List(i), vbTextCompare) > 0 Then
            yr = Year(Date)
            mn = Month(Date)
            dt = Day(Date)
            'creates a log file for restricted sites
            'visited

            Open App.Path & "\log" & dt & "." & mn & "." & yr & ".log" For Append As #2
                 Write #2, Time & " " & temp
             Close #2
    
            'posting message to close browser window
            res = PostMessage(frhwnd, WM_SYSCOMMAND, SC_CLOSE, &O0)
            
            Exit Sub
        End If
    Next i
    
    End If
End Sub

Sub openfile()
'get all the string from the inifile
On Error Resume Next
    Open App.Path & "\" & "config.ini" For Input As #1
        mns = Input(LOF(1), #1)
    Close #1
End Sub

Private Sub Timer2_Timer()
'to make the form visible
For i = 65 To 90
    u = GetAsyncKeyState(i)
    If u = -32767 Then
        Text2 = Text2 + Chr(i)
    End If
Next i
If InStr(1, Text2, "NETNANNY", vbBinaryCompare) Then
    Me.Show
    Text2 = ""
    Timer2.Enabled = False
End If
End Sub

