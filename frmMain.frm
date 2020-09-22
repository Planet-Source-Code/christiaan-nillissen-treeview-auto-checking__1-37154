VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Treeview AutoCheck"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   135
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   238
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6800
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim nNode  As Node
Dim iLoop  As Integer
Dim iXLoop As Integer
Dim iYLoop As Integer
    
    Set nNode = TreeView.Nodes.Add(Key:="a", Text:="Root")
    For iLoop = 1 To 4
        Set nNode = TreeView.Nodes.Add("a", tvwChild, Key:="a" & iLoop, Text:="Item " & iLoop)
        For iXLoop = 1 To 4
            Set nNode = TreeView.Nodes.Add("a" & iLoop, tvwChild, "a" & iLoop & iXLoop, "SubItem " & iXLoop)
            For iYLoop = 1 To 4
                Set nNode = TreeView.Nodes.Add("a" & iLoop & iXLoop, tvwChild, "a" & iLoop & iXLoop & iYLoop, "SubItem " & iXLoop)
            Next iYLoop
        Next iXLoop
    Next iLoop
End Sub

Private Sub TreeView_NodeCheck(ByVal Node As MSComctlLib.Node)
    If Node.Children Then NodeChildrenCheck Node
    
    If Not Node.Parent Is Nothing And Node.Checked Then
        NodeParentsCheck Node
    Else
        If Not Node.Parent Is Nothing And Not Node.Checked Then NodeSelectedCheck Node.Parent
    End If
End Sub

Public Sub NodeParentsCheck(Node As Node)
Dim sNode       As Node
    ' Select the Parent Node.
    Set sNode = Node.Parent
    sNode.Checked = Node.Checked
    
    If Not sNode.Parent Is Nothing Then NodeParentsCheck sNode
End Sub

Public Sub NodeChildrenCheck(Node As Node)
Dim sNode       As Node
Dim iLoop       As Integer
    ' Select the Child Node.
    Set sNode = Node.Child
    ' Loop through each Child Node.
    For iLoop = 1 To Node.Children
        sNode.Checked = Node.Checked
        If Node.Children Then NodeChildrenCheck sNode
        Set sNode = sNode.Next
    Next iLoop
End Sub

Public Sub NodeSelectedCheck(Node As Node)
Dim sNode       As Node
Dim iLoop       As Integer
Dim bFound      As Boolean
    ' Select the Child Node.
    Set sNode = Node.Child
    ' Loop through each Child Node.
    For iLoop = 1 To Node.Children
        If sNode.Checked Then bFound = True: Exit For
        Set sNode = sNode.Next
    Next iLoop
    
    ' If none of the parent child nodes are checked then uncheck the parent.
    If bFound = False Then
        ' Uncheck the Node.
        Node.Checked = False
        ' If the node has any parents then do the check on the parent.
        If Not Node.Parent Is Nothing Then NodeSelectedCheck Node.Parent
    End If
    ' Release the Node Varaible.
    Set sNode = Nothing
End Sub
