VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidfrost 2k1"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Server"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   7680
      Width           =   2175
   End
   Begin VB.ListBox LstIP 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6690
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   6
      Top             =   775
      Width           =   4935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Server"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   60
      Top             =   7320
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   5040
      TabIndex        =   0
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3735
      Left            =   5040
      TabIndex        =   1
      Top             =   3720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   360
      Left            =   6600
      TabIndex        =   2
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   635
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.ListBox LstHostIP 
      Appearance      =   0  'Flat
      Height          =   2760
      ItemData        =   "Form1.frx":001D
      Left            =   360
      List            =   "Form1.frx":0024
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView5 
      Height          =   1485
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2619
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hostname"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox TxtIndex 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "IP Address"
      Height          =   195
      Left            =   1440
      TabIndex        =   16
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ping to server"
      Height          =   195
      Left            =   9600
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LblCount 
      AutoSize        =   -1  'True
      Caption         =   "Player Count:"
      Height          =   195
      Left            =   7800
      TabIndex        =   13
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "N/A"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10710
      TabIndex        =   12
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6480
      TabIndex        =   9
      Top             =   7680
      Width           =   3915
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Server Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5055
      TabIndex        =   4
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "N/A"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
Option Explicit
Dim ctr As Integer
Dim ping1 As Currency
Dim hptimer As Boolean
Dim autorefreshon As Boolean
Dim strBaseGame As String
Dim strMod As String
Dim hnparam As String
Dim players As Integer
Dim maxplayers As Integer
Dim l1so As Boolean
Dim l2so As Boolean
Dim l2sortcolumn As Integer
Dim l1sortcolumn As Integer
Dim strGame As String
Dim strCmdLine As String
Dim strExePath As String
Dim autorefreshtime As Integer
Dim closebrowser As Boolean


Private Sub CmdConnect_Click()
MsgBox "That comes last :)"
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
If LstIP.ListIndex = -1 Then Exit Sub
LstHostIP.RemoveItem LstIP.ListIndex
LstIP.RemoveItem LstIP.ListIndex
ListView1.ListItems.Clear
ListView2.ListItems.Clear
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()
If Text1.Text = "" Then Exit Sub
LstIP.AddItem Text1.Text
LstHostIP.AddItem Text1.Text
LstIP.SetFocus
LstIP.ListIndex = LstIP.ListCount - 1
LstIP.ListIndex = LstIP.ListCount - 1
End Sub

Private Sub Form_Load()
    

   
        'strBaseGame = "Q3"

     
        ListView1.ColumnHeaders.Add , , "Player"
        ListView1.ColumnHeaders.Add , , "Score"
        ListView1.ColumnHeaders.Add , , "Ping"
        ListView1.ColumnHeaders.Add , , "Slot"
        
        ListView3.ColumnHeaders.Add , , "Ping"
        ListView3.ColumnHeaders.Add , , "Score"
        ListView3.ColumnHeaders.Add , , "Player"
        ListView2.ColumnHeaders.Add , , "Name"
        ListView2.ColumnHeaders.Add , , "Value"
        
        ListView1.SortKey = 1

ListView1.ColumnHeaders.Item(1).Width = 3089
ListView1.ColumnHeaders.Item(2).Width = 1000
ListView1.ColumnHeaders.Item(3).Width = 1000
ListView1.ColumnHeaders.Item(4).Width = 1000
'ListView2.ColumnHeaders.Item(2).Width = 800
ListView2.ColumnHeaders.Item(1).Width = 2280
ListView2.ColumnHeaders.Item(2).Width = 3830
'ListView5.ColumnHeaders.Item(1).Width = ListView5.Width
End Sub



Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

      
ColumnHeaderClick Me.ListView1, ColumnHeader
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ColumnHeaderClick Me.ListView2, ColumnHeader
End Sub


Public Sub ColumnHeaderClick(LV As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim LvIndex As Long
  On Error Resume Next
    LvIndex = LV.Index
    If Err Then
      LvIndex = 0
    End If
  On Error GoTo 0
  LV.ColumnHeaders.Item(LV.SortKey + 1).Icon = 0
  If LV.SortKey = ColumnHeader.Index - 1 Then
    If LV.SortOrder = lvwAscending Then
      LV.SortOrder = lvwDescending
    Else
      LV.SortOrder = lvwAscending
    End If
    SaveSetting App.CompanyName, App.Title, "SortOrder" & LV.name & LvIndex, LV.SortOrder
  Else
    LV.SortKey = ColumnHeader.Index - 1
    SaveSetting App.CompanyName, App.Title, "SortKey" & LV.name & LvIndex, LV.SortKey
  End If

  LV.SetFocus
  DoEvents
  If LV.ListItems.Count > 0 Then
    LV.ListItems(LV.SelectedItem.Index).EnsureVisible
  End If
End Sub

    
    



Private Sub LstHostIP_Click()
LstIP.ListIndex = LstHostIP.ListIndex
End Sub

Private Sub LstIP_Click()
If LstIP.Text = "" Then Exit Sub

TxtIndex.Text = LstIP.ListIndex
    
    LstHostIP.ListIndex = LstIP.ListIndex
    hnparam = LstHostIP.Text
    strGame = "Q3"
 
    Label2.Visible = False
    Label5.Visible = False
    Label2.Caption = ""
    Label5.Caption = ""
    'ListView2.ListItems.Clear


        Q3_sendData
  
   
  
   
    Label4.Visible = True
    Label4.Caption = "Waiting for response"
    'Label4.Left = (ListView1.Left + ListView1.Width) - (Label4.Width / 2)
 
End Sub

Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo error_handler
    Dim y As String
    Winsock1.GetData y
    Dim temp As String
    Dim ping2, freq As Currency
    Dim itmx As ListItem
      
    If hptimer Then
        QueryPerformanceCounter ping2
        QueryPerformanceFrequency freq
        ping2 = Int(((ping2 - ping1) / freq * 1000))
    Else
        ping2 = ctr
    End If
    Timer1.Enabled = False
    Dim rules As String
    Dim players As String
    Dim rule As String
    Dim rvalue As String
    Dim pname As String
    Dim pfrags As String
    Dim pping As String
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    If (InStr(1, y, "statusResponse", vbBinaryCompare)) > 0 Then
        'pack is probably valid
        rules = Mid$(y, InStr(1, y, vbLf, vbBinaryCompare) + 1)
        players = Mid$(rules, InStr(1, rules, vbLf, vbBinaryCompare) + 1)
        rules = Left$(rules, InStr(1, rules, vbLf, vbBinaryCompare))
       
        While rules <> vbLf
            'trim the slash
            rules = Mid$(rules, 2)
            rule = Mid$(rules, 1, InStr(1, rules, "\", vbBinaryCompare) - 1)
            rules = Mid$(rules, Len(rule) + 2)
            If InStr(1, rules, "\", vbBinaryCompare) <> 0 Then
                rvalue = Mid$(rules, 1, InStr(1, rules, "\", vbBinaryCompare) - 1)
            Else
                rvalue = Mid$(rules, 1, Len(rules) - 1)
            End If
            rules = Mid$(rules, Len(rvalue) + 1)
            lv2add rule, rvalue
        Wend
        While players <> ""
            pfrags = Left$(players, InStr(1, players, " ", vbBinaryCompare) - 1)
            players = Mid$(players, Len(pfrags) + 2)
            pping = Left$(players, InStr(1, players, " ", vbBinaryCompare) - 1)
            players = Mid$(players, Len(pping) + 3)
            pname = Left$(players, InStr(1, players, Chr$(34), vbBinaryCompare) - 1)
            players = Mid$(players, Len(pname) + 3)
            pname = Q3_StripColours(pname)
            'lv1add pping, pfrags, pname
            lv1add pname, pfrags, pping
        Wend
        If ping2 > 200 Then
            Label2.ForeColor = vbRed
        Else
            Label2.ForeColor = vbBlack
        End If
        Label2.Caption = "  " & ping2 & "ms"
        Label2.Visible = True
        Label5.Visible = True
      
       'COMMAND1 ENABLED
        
        Dim i As Integer
        i = 1
        While i < ListView2.ListItems.Count + 1 And maxplayers = 0
            If ListView2.ListItems.Item(i).Text = "sv_maxclients" Then maxplayers = ListView2.ListItems.Item(i).SubItems(1)
            i = i + 1
        Wend
        i = 1
        Dim hostname As String
        While i < ListView2.ListItems.Count + 1 And hostname = ""
            If ListView2.ListItems.Item(i).Text = "sv_hostname" Then hostname = ListView2.ListItems.Item(i).SubItems(1)
            i = i + 1
        Wend
        Label3.Caption = hostname
        Label6.Caption = hnparam
       ListView2.ColumnHeaders.Item(2).Width = ListView2.Width - 2250
      
        players = ListView1.ListItems.Count
        Label5.Caption = players & "/" & maxplayers
        If players = maxplayers Then
            Label5.ForeColor = vbRed
        Else
            Label5.ForeColor = vbBlack
        End If
        
        ListView2.HideSelection = True
        ListView1.HideSelection = True
        Label4.Visible = False
        'ListView1.Visible = True
        'ListView2.Visible = True
        If l1sortcolumn = 1 Then
            'yes this is ugly, but it wouldnt let me call the click function
          
        End If
    ElseIf (InStr(1, y, "infoResponse", vbBinaryCompare)) > 0 Then
        'theres nothing here we need, use for ping timeing only
    Else
        'its a bad packet
    End If




'RAD block
    If LstIP.Text = hostname Then
   ' do nothing if the HOSTNAME was already retrieved
      Else
      ' remove ip from lstip
      ' and add the hostname instead
      ' thus switching the ip over to lsthostip list
      
      LstIP.RemoveItem TxtIndex.Text
      LstHostIP.RemoveItem TxtIndex.Text
      LstIP.AddItem hostname, TxtIndex
      LstHostIP.AddItem hnparam, TxtIndex.Text
     
      End If





Exit Sub
error_handler:
    clearform
    Timer1.Enabled = False
    If Err.Number = 10054 Then
        Label4.Caption = "Winsock error 10054"
        Label4.Left = (ListView1.Left + ListView1.Width) - (Label4.Width / 2)
    Else
        Label4.Caption = "Received corrupt data"
        Label4.Left = (ListView1.Left + ListView1.Width) - (Label4.Width / 2)
    End If
End Sub

Private Sub Timer1_Timer()
    ctr = ctr + 10
    If ctr > 1200 Then
        'server timed out
        'disable the timer
        'close the winsock1 port
        Timer1.Enabled = False
        Winsock1.Close
        Label4.Caption = "Timed out"
       ' Label4.Left = (ListView1.Left + ListView1.Width) - (Label4.Width / 2)
        
        Label3.Caption = hnparam
        
        Label2.Visible = False
        Label5.Visible = False
      
        'Command1.SetFocus
    End If
End Sub


Private Sub Timer2_Timer()
    Timer2.Enabled = False
    
    Timer2.Interval = autorefreshtime * 1000
    Timer2.Enabled = True
End Sub



Private Sub lv1add(i1 As Variant, i2 As Variant, i3 As Variant)
        Dim itmx As ListItem
        Set itmx = ListView1.ListItems.Add(, , i1) 'playerping
        itmx.SubItems(1) = i2 'playerscore
        itmx.SubItems(2) = i3 'playername
        itmx.SubItems(3) = ListView1.ListItems.Count 'slot number
        Set itmx = Nothing
End Sub
Private Sub lv2add(i1 As Variant, i2 As Variant)
        Dim itmx As ListItem
        Set itmx = ListView2.ListItems.Add(, , i1)
        itmx.SubItems(1) = i2
        Set itmx = Nothing
End Sub

Private Sub clearform()
    Label2.Visible = False
    Label5.Visible = False
    Label2.Caption = ""
    Label5.Caption = ""
    Label4.Visible = True
   
End Sub

Private Function Q3_StripColours(name As String) As String
    Dim i As Integer
    i = 1
    Dim toBeReturned As String
    Dim temp As String
    Dim temp2 As String
    While i < Len(name) + 1
        temp = Mid$(name, i, 1)
        If i = Len(name) Then
            temp2 = i
        Else
            temp2 = Mid$(name, i + 1, 1)
        End If
        
        If temp = "^" And (temp2 >= "0" And temp2 <= "9") Then
            i = i + 1
        Else
            toBeReturned = toBeReturned & temp
        End If
        i = i + 1
    Wend
    Q3_StripColours = toBeReturned
    
End Function



Private Sub Q3_sendData()
    On Error GoTo error_handler
    Dim temp1 As String
    Dim temp2 As String
    temp1 = Left$(hnparam, Len(hnparam) - Abs((InStr(1, hnparam, ":", vbBinaryCompare) - Len(hnparam))) - 1)
    temp2 = Mid$(hnparam, InStr(1, hnparam, ":", vbBinaryCompare) + 1)
    Winsock1.RemoteHost = temp1
    Winsock1.RemotePort = temp2
    If Winsock1.State = 1 Then Winsock1.Close
    ctr = 0
    Winsock1.Bind
    If QueryPerformanceCounter(ping1) Then
        hptimer = True
    Else
        hptimer = False
    End If
    Timer1.Enabled = True
    Winsock1.SendData ("ÿÿÿÿgetstatus")
    'use for ping, redundant otherwise
    'Winsock1.sendData ("ÿÿÿÿgetinfo")
Exit Sub
error_handler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error"
End Sub
