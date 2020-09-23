VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Treeview Extender Class Example"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtselectedkey 
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6675
      Width           =   8055
   End
   Begin VB.CommandButton cmddeleteselected 
      Caption         =   "Delete selected node"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   5520
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "dirclosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27B4
            Key             =   "diropen"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdaddselectednode 
      Caption         =   "Add child node to selected"
      Height          =   300
      Left            =   5520
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddNode 
      Caption         =   "Add Node"
      Height          =   300
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Collapse All Nodes"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdExpandAll 
      Caption         =   "Expand All Nodes"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Text            =   "node-1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdExists 
      Caption         =   "Check if Node Exists"
      Height          =   300
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11456
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Node Key"
      Height          =   2175
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tv As ExtendedTreeView 'set reference to class
Private Sub Form_Load()
Set tv = New ExtendedTreeView 'Create object from class
Set tv.TVctl = TreeView1 'Set the tv.tvctl to the treeview control on your form
' Now all the extended functions are available through tv and the original
' object is available through tv.TVctl or the original object name in this case Treeview1
tv.maketestTree "dirclosed", "diropen" ' populate the treeview with test data

End Sub
Private Sub Form_Activate()
'add selected key to txtselected key
' This is only used to display the current key and so it is cut/pastable
' I added this because the KeyExist function checks for a node key not node text, and
' thought that without seeing the key this might confuse those new to the treeview control.
txtselectedkey = TreeView1.SelectedItem.Key
End Sub
Private Sub cmddelete_Click()
    tv.delete txtKey
End Sub

Private Sub cmddeleteselected_Click()
 tv.delete TreeView1.SelectedItem.Key
End Sub

Private Sub cmdAddNode_Click()
tv.AddRootNode txtKey, txtKey, "dirclosed", "diropen"
End Sub

Private Sub cmdaddselectednode_Click()
tv.AddChildNodeSelected txtKey, txtKey, "dirclosed", "diropen"
End Sub

Private Sub cmdExpandAll_Click()
tv.ExpandAll
End Sub

Private Sub Command1_Click()
tv.CollapseAll
End Sub



Private Sub cmdExists_Click()
If txtKey & "" = "" Then ' user didn't enter text
    MsgBox "Please type in a node name"
    Exit Sub
End If
MsgBox tv.exist(txtKey), vbOKOnly, "Key Exist?"

End Sub


Private Sub TreeView1_Click()
txtselectedkey = TreeView1.SelectedItem.Key
End Sub
