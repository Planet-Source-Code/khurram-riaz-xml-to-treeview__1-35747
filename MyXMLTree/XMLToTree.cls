VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLToTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NewTree As MSComctlLib.TreeView
Dim XMLDocument As MSXML2.DOMDocument30

Public Function MakeXMLDocument(XMLFile As String, Validate As Boolean) As Boolean
    Dim LoadCheck As Boolean
    Set XMLDocument = New MSXML2.DOMDocument30
    If Validate Then
        XMLDocument.validateOnParse = True
    Else
        XMLDocument.validateOnParse = False
    End If
    LoadCheck = XMLDocument.Load(XMLFile)
   
    If LoadCheck = False Then
        MsgBox XMLDocument.parseError.reason
        MakeXMLDocument = False
        Exit Function
    End If
    MakeXMLDocument = True
    
End Function

Public Sub PopulateTree(TreeView As Object)
    Dim root As IXMLDOMNode
    Set NewTree = TreeView
    Index = 0
    Set root = XMLDocument.documentElement
    NewTree.Nodes.Add , , root.nodeName & CStr(Index), root.nodeName
    InsertNodes CStr(Index), XMLDocument.documentElement
End Sub

Private Sub InsertNodes(pIndex As String, Node As IXMLDOMNode)
    Dim parentNode As IXMLDOMNode
    Dim childNode As IXMLDOMNode
    Dim newPIndex As String
    
    chIndex = 0
    Set parentNode = Node
    
    If parentNode.hasChildNodes Then
        For Each Item In parentNode.childNodes
            If Item.nodeType = NODE_ELEMENT Then
                Set childNode = Item
                newPIndex = pIndex & "_" & CStr(chIndex)
                NewTree.Nodes.Add parentNode.nodeName & pIndex, tvwChild, childNode.nodeName & newPIndex, childNode.nodeName
                chIndex = chIndex + 1
                InsertNodes newPIndex, childNode
            End If
        Next
    Else
        Exit Sub
    End If
    
End Sub
