VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form XMLPresenter 
   BackColor       =   &H00404040&
   Caption         =   "XMLPresenter"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Validate 
      Caption         =   "Validate"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   960
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox XMLFile 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Browse 
      Caption         =   ".."
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5741
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Specify XML Document to be Opened"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "XMLPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fName As String
Dim XML_Tree As XMLToTree

Private Sub Browse_Click()

    ComDlg.CancelError = True
    On Error GoTo Error
    ComDlg.DialogTitle = "Open XML Document"
    ComDlg.Filter = "XML Document(*.XML)|*.xml"
    ComDlg.ShowOpen
    fName = ComDlg.FileName
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    XMLFile.Text = FSO.GetFileName(fName)
 
Error:
    Exit Sub
End Sub

Private Sub Clear_Click()
    TV.Nodes.Clear
    Load.Enabled = True
End Sub

Private Sub Exit_Click()
    Unload XMLPresenter
End Sub

Private Sub Load_Click()
    
    If XMLFile.Text <> "" Then
        Load.Enabled = False
        Clear.Enabled = True
        Set XML_Tree = New XMLToTree
        
        If XML_Tree.MakeXMLDocument(fName, Validate.Value) Then
            XML_Tree.PopulateTree TV
            TV.Nodes.Item(1).Expanded = True
            If TV.Nodes.Count > 1 Then
                TV.Nodes.Item(TV.Nodes.Count - 1).Expanded = True
            End If
        Else
             MsgBox "Error: Unable to Load XML Document"
        End If
    Else
        MsgBox "XML File not specified"
    End If
End Sub


