VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtendedTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Extended Treeview Class Module by RegX Copyright 2002
'
' This control adds some basic functionality to the treeview control such as Exist and AddChildNode
' This class also takes care of most of the error handling
' Things to add
' 1. Drag and drop support
' 2. Normal Raise error event type error handling
' 3. Catch image not in listview error
' 4. Get rid of public vars and use proper OO properties


' Usage
'(Form Declarations)
'   Public tv As ExtendedTreeView
'(Form_Load)
'   Set tv = New ExtendedTreeView
'   Set tv.TVctl = TreeView1

Option Explicit
Public WithEvents TVctl As TreeView
Attribute TVctl.VB_VarHelpID = -1
Public Sub maketestTree(Optional image As String, Optional selectedimage As String)
'If IsEmpty(image) = True Or TVctl.ImageList Is Nothing Then
' This just populates the treeview with data for testing
    Dim a As Long
    Dim b As Long
    Dim c As Long
For a = 1 To 5
    AddNode , , "node" & TVctl.PathSeparator & a, "node" & TVctl.PathSeparator & a, image, selectedimage
    For b = 1 To 5
        AddNode "node" & TVctl.PathSeparator & a, tvwChild, "node" & TVctl.PathSeparator & a & TVctl.PathSeparator & b, "node" & TVctl.PathSeparator & a & TVctl.PathSeparator & b, image, selectedimage
        For c = 1 To 5
                AddNode "node" & TVctl.PathSeparator & a & TVctl.PathSeparator & b, tvwChild, "node" & TVctl.PathSeparator & a & TVctl.PathSeparator & b & TVctl.PathSeparator & c, "node" & TVctl.PathSeparator & a & TVctl.PathSeparator & b & TVctl.PathSeparator & c, image, selectedimage
        Next c
    Next b
Next a

End Sub
Public Function exist(strKey As String) As Boolean
On Error GoTo bail
    Dim newnode As Node
    Set newnode = TVctl.Nodes.Add(, , strKey)
    TVctl.Nodes.Remove (newnode.Index)
    exist = False
Exit Function
bail:
    exist = True
End Function
Public Sub ExpandAll()
Dim Node As Node
For Each Node In TVctl.Nodes
    Node.Expanded = True
Next
End Sub
Public Sub CollapseAll()
Dim Node As Node
For Each Node In TVctl.Nodes
    Node.Expanded = False
Next
End Sub
Public Sub AddNode(Optional relationship As String, Optional child As TreeRelationshipConstants, Optional Key As String, Optional text As String, Optional image As String, Optional selectedimage As String)
' This function hendles most of the errors associated with adding nodes to a treeview control
On Error GoTo bail
    If exist(Key) = True Then
        MsgBox "Key " & Key & " already exist", vbOKOnly, "Error"
        Exit Sub
    End If
    If relationship <> "" And exist(relationship) = False Then
        MsgBox "Relative Key " & relationship & " doesn't exist", vbOKOnly, "Error"
        Exit Sub
    End If
    If relationship = "" Then
            If image <> "" And selectedimage <> "" Then
                TVctl.Nodes.Add , , Key, text, image, selectedimage
            ElseIf image <> "" Then
                TVctl.Nodes.Add , , Key, text, image
            Else
                TVctl.Nodes.Add , , Key, text
            End If
                
    Else
            If image <> "" And selectedimage <> "" Then
                TVctl.Nodes.Add relationship, child, Key, text, image, selectedimage
            ElseIf image <> "" Then
                TVctl.Nodes.Add relationship, child, Key, text, image
            Else
                TVctl.Nodes.Add relationship, child, Key, text
            End If
    End If
Exit Sub
bail:
MsgBox image & " or " & selectedimage & " doesn't exist in the listview control", vbOKOnly, "Error"
End Sub


Public Sub AddRootNode(Optional Key As String, Optional text As String, Optional image As String, Optional selectedimage As String)
    AddNode , , Key, text, image, selectedimage
End Sub
Public Sub AddChildNodeSelected(Optional Key As String, Optional text As String, Optional image As String, Optional selectedimage As String)
'This function adds a child node to the currently selected node
'and uses the selected nodes key plus the current treeview path seperator plus the key text for the new key
'This allows the same key to exist anywhere in the tree

TVctl.SelectedItem.Expanded = True
TVctl.Nodes.Add TVctl.Nodes(TVctl.SelectedItem.Index).Key, tvwChild, TVctl.SelectedItem.Key & TVctl.PathSeparator & Key, text, image, selectedimage

End Sub

Private Sub TVctl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Set TVctl.DropHighlight = TVctl.HitTest(x, y)
End Sub
Public Sub delete(Key As String)
    If exist(Key) = True Then
        TVctl.Nodes.Remove (TVctl.Nodes(Key).Index)
    Else
        MsgBox "Key " & Key & " doesn't exist", vbOKOnly, "Error"
    End If
End Sub
