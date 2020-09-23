VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo For Loading Data To Listview From Sequantial File !"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thaks For Vote To Me At www.Planet Source Code.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author   : Looi Tuck Wai
'Control  : 3 Command Button(cmdAdd,cmdSave,cmdLoad),3 Textbox
'(Text1,Text2,Text3),1 listview(lv)
Dim a() As String
Dim b() As String
Dim c() As String
Dim i As Integer
Private Sub Form_Load()
With lv
     .View = lvwReport
     .ColumnHeaders.Add , , "Name"
     .ColumnHeaders.Add , , "Age"
     .ColumnHeaders.Add , , "Lucky Number"
     .GridLines = True
     .FullRowSelect = True
End With
End Sub

Private Sub CmdAdd_Click()
Dim itm As ListItem
      Set itm = lv.ListItems.Add(, , Text1.Text)
      itm.SubItems(1) = Text2.Text
      itm.SubItems(2) = Text3.Text
End Sub

Private Sub CmdSave_Click()
    ItemCount = Form1.lv.ListItems.Count
    For i = 1 To ItemCount
        ReDim Preserve a(i) As String
        ReDim Preserve b(i) As String
        ReDim Preserve c(i) As String
        a(i) = Form1.lv.ListItems(i).Text
        b(i) = Form1.lv.ListItems(i).SubItems(1)
        c(i) = Form1.lv.ListItems(i).SubItems(2)
        Next i
        
    Open App.Path & "\test.lis" For Output As #1
    For i = 1 To ItemCount
        Write #1, a(i), b(i), c(i)
        Next i
        Close #1
End Sub

Private Sub CmdLoad_Click()
Dim itm As ListItem
Dim a, b, c As String
    Open App.Path & "\test.lis" For Input As #1
    Do Until EOF(1)
    Input #1, a, b, c
              Set itm = lv.ListItems.Add(, , a)
              itm.SubItems(1) = b
              itm.SubItems(2) = c
    Loop
    Close #1
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim k As Integer
k = MsgBox("Save Changed Before Exit", vbYesNo, "Save Change ?")
   If k = vbYes Then
       CmdSave_Click
       MsgBox "Data Saved", vbInformation, "Data Saved"
   End If
   If k = vbNo Then
       Unload Me
   End If
End Sub

'*Note that "test.lis" is a sample,
'you can change it to fit on your system
Private Sub lv_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
Text1.Text = lv.SelectedItem.Text
Text2.Text = lv.SelectedItem.SubItems(1)
Text3.Text = lv.SelectedItem.SubItems(2)
End Sub
