VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBinaryTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' References/research links used
'http://msdn.microsoft.com/en-us/library/ms379572.aspx
'http://www.mactech.com/articles/mactech/Vol.06/06.08/BinaryTrees/index.html
'http://en.wikipedia.org/wiki/Binary_tree
'http://upload.wikimedia.org/wikipedia/en/c/c4/Tree_Rebalancing.gif



' THIS CLASS IS DESIGNED FOR BOTH KEY & VALUE
' Keys can be numeric and can be case-insensitive (see CaseSensitiveKeys property)
' When passing Index to any function...
'   -- If Index is a node index, not a key, pass Index as a numeric data type
'   -- If Index is a Key, not a node index, pass Index as a string data type
' Node indexes should never be cached. They can change whenever any node is deleted

' By default, all keys are case-sensitive (see CaseSensitiveKeys property)
' By default, duplicate keys are not allowed and will raise an error if one is attempted
'   -- You can override this by providing a True value for the optional AllowDuplicateKeys
'       parameter in the Add & ReKey functions
'   -- Should you be willing to allow duplicate keys, you must handle some extra work
'       when calling these two functions:
'       1. Delete.  Pass the numerical node Index, not the Key, as the function's Index parameter
'       2. ReKey. Pass the numerical node index, not the Key, as the function's Index parameter
'       3. Do not set a Key's value via the NodeKey property as it only allows non-duplicate keys
'       To get the correct index, when keys can be duplicated:
'           Index = Find(Key,0&)
'           If NodeValue(Index) = criteria Then you have the correct index
'           Else call Find again & again, passing the Index+1... until criteria is found
'           Index = Find(Key, Index + 1)

Private Type BTreeNode
    LeftChild As Long   ' reference to left child
    RightChild As Long  ' reference to right child
    Balance As Long     ' balance factor ranges from -2 to +2
    Value As Long       ' change to String, Variant or other data type as needed. Can be ObjPtr, StrPtr, etc
    Key As String       ' a unique key for the data, like in Collections
End Type

Private Type BTreeKeyStruct
    Count As Long
    Node() As BTreeNode
End Type

Private BTreeKey As BTreeKeyStruct      ' the tree
Private KeyCompareBinary As Boolean     ' option for Text or Binary key compare

Public Property Let CaseSensitiveKeys(ByVal NewValue As Boolean)
    ' by default, keys are case-sensitive. Change if needed
    ' Note: This can only be set when the tree is empty
    If NewValue Xor (KeyCompareBinary = vbBinaryCompare) Then
        If BTreeKey.Count = 0& Then
            If NewValue Then KeyCompareBinary = vbBinaryCompare Else KeyCompareBinary = vbTextCompare
        End If
    End If
End Property
Public Property Get CaseSensitiveKeys() As Boolean
    CaseSensitiveKeys = (KeyCompareBinary = vbBinaryCompare)
End Property

Public Function Add(ByVal Key As String, ByVal Value As Long, ByVal AllowDuplicateKeys As Boolean) As Boolean
    ' Function adds a key/value to the tree.
    ' Returns True if no errors occurred else raises an error
    
    If Key = vbNullString Then Exit Function
    If BTreeKey.Count = 0& Then
        Add = (CreateNode(0&, True, Key, Value) > 0&)
    Else
        Select Case LocalInsert(0&, Key, Value, AllowDuplicateKeys)
        Case -1: ' failed to add node; out of memory? some other error
            Err.Raise 5, "cBinaryTree.Add", "System Error Preventing Tree Modification"
        Case 4: ' failed to load because of key duplication
            Err.Raise 457, "cBinaryTree.Add" ' duplicate key
        Case Else: Add = True
        End Select
    End If
End Function

Public Function Find(ByVal Key As String, Optional ByVal FromIndex As Long = 0&) As Long
    ' Function returns the Node index of the node containing Key
    Find = ValidateIndex(Key, FromIndex)
End Function

Public Function ReKey(Index As Variant, ByVal NewKey As String, Optional ByVal AllowDuplicateKeys As Boolean = False) As Long
    
    ' Function changes the key of a node and returns its new Index if no error occurs
    ' Note: The NodeKey property also calls this function but defaults to no duplicate keys
    
    Dim Value As Long
    Dim nodeIndex As Long
    
    If BTreeKey.Count = 0& Then Exit Function ' no tree yet
    If NewKey = vbNullString Then
        Err.Raise 380, "cBinaryTree.ReKey" ' invalid property value
        Exit Function
    End If
    
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then
        If StrComp(BTreeKey.Node(nodeIndex).Key, NewKey, KeyCompareBinary) = 0& Then
            ReKey = nodeIndex
        Else
            If Not AllowDuplicateKeys Then
                If LocalFind(NewKey, BTreeKey.Node(0).LeftChild + 0&) Then
                    Err.Raise 457, "cBinaryTree.ReKey" ' duplicate key
                    Exit Function ' new key would cause duplicate key
                End If
            End If
            Value = BTreeKey.Node(nodeIndex).Value
            LocalDelete vbNullString, nodeIndex
            LocalInsert 0&, NewKey, Value, AllowDuplicateKeys
            ReKey = BTreeKey.Count
        End If
    Else
        Err.Raise 380, "cBinaryTree.ReKey" ' invalid property value
    End If
End Function

Public Sub Clear()
    ' sub clears the tree
    Erase BTreeKey.Node()
    BTreeKey.Count = 0
End Sub

Public Function NodeChild(Index As Variant, ByVal LeftChild As Boolean) As Long
    ' Property retrieves the Node's Left or Right child
    Dim nodeIndex As Long
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then
        If LeftChild Then
            NodeChild = BTreeKey.Node(nodeIndex).LeftChild
        Else
            NodeChild = BTreeKey.Node(nodeIndex).RightChild
        End If
    End If
End Function
Public Property Get NodeValue(Index As Variant) As Long
    ' Property sets/retrieves a Node's Value
    Dim nodeIndex As Long
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then NodeValue = BTreeKey.Node(nodeIndex).Value
End Property
Public Property Let NodeValue(Index As Variant, ByVal NewValue As Long)
    Dim nodeIndex As Long
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then BTreeKey.Node(nodeIndex).Value = NewValue
End Property
Public Property Get NodeKey(Index As Variant) As String
    ' Property sets/retrieves a Node's Key (see ReKey also)
    Dim nodeIndex As Long
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then NodeKey = BTreeKey.Node(nodeIndex).Key
End Property
Public Property Let NodeKey(Index As Variant, ByVal NewKey As String)
    ' does not allow duplicate keys. If duplicate keys are desired, use ReKey instead
    Dim nodeIndex As Long
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then Me.ReKey BTreeKey.Node(nodeIndex).Key, NewKey, False
End Property
Public Property Get NodeRootIndex() As Long
    ' Property retrieves the Index of the root
    If BTreeKey.Count Then NodeRootIndex = BTreeKey.Node(0).LeftChild
End Property
Public Property Get NodeCount() As Long
    ' Property retrieves a number of nodes in the tree
    NodeCount = BTreeKey.Count
End Property
Public Property Get NodeBalance(Index As Variant) As Long
    ' Property retrieves a Node's Balance factor
    Dim nodeIndex As Long
    nodeIndex = ValidateIndex(Index)
    If nodeIndex Then NodeBalance = BTreeKey.Node(nodeIndex).Balance
End Property

Public Function Delete(Index As Variant) As Boolean
    ' Function deletes a node
    ' If Index is a Key, pass as a string variable type
    ' If Index is a node index, pass a numeric variable type
    If VarType(Index) = vbString Then
        Delete = LocalDelete(CStr(Index), 0&)
    Else
        Dim nodeIndex As Long
        nodeIndex = ValidateIndex(Index)
        If nodeIndex Then Delete = LocalDelete(vbNullString, nodeIndex)
    End If
End Function

Private Function LocalInsert(Parent As Long, Key As String, Value As Long, AllowDupKeys As Boolean) As Long

    ' Converted from C++ example at http://www.mactech.com/articles/mactech/Vol.06/06.08/BinaryTrees/index.html

    ' The height of a node is the number of steps to its most distant descendant.
    ' The height of a tree (or a sub-tree) is the height of its root node.
    ' A binary tree is balanced if for every node in the tree, the height of the left sub-tree is within 1 of the height of the right sub-tree.
    ' Here’s another way to say the same thing: each node in a binary tree has a balance factor,
    ' which is equal to the height of the left sub-tree minus the height of the right sub-tree.
    ' A binary tree is balanced if every balance factor is either 0, 1 or -1.
    
    Dim Pivot As Long, Result As Long, dir As Long
    Dim insParent As Long, Child As Long, cmpResult As Long
    
    On Error GoTo EH
    If Parent = 0& Then ' root node always moves to the left
        dir = 1
        Pivot = BTreeKey.Node(0).LeftChild
    Else
        cmpResult = StrComp(Key, BTreeKey.Node(Parent).Key, KeyCompareBinary)
        If cmpResult = -1 Then
            dir = 1         ' move to the left
            Pivot = BTreeKey.Node(Parent).LeftChild
        ElseIf cmpResult = 1 Then
            dir = 2         ' move to the right
            Pivot = BTreeKey.Node(Parent).RightChild
        Else                ' match, allow only if AllowDupKeys is True
            If AllowDupKeys = False Then
                LocalInsert = 4&
                Exit Function
            End If
        End If
    End If
    If Pivot = 0& Then      ' add new node here
        CreateNode Parent, (dir = 1), Key, Value
        LocalInsert = dir
        Exit Function
    End If
    ' continue recursing down until open node is found
    Result = LocalInsert(Pivot, Key, Value, AllowDupKeys)
    If Result < 1& Or Result = 4& Then
        LocalInsert = Result
        Exit Function
    End If
    ' when recursion ends, the Result variable holds the path of the last
    ' 3 previous recursions (left to right bit order) in 4 bit blocks
    
    If (Result And &HF) = 1 Then ' moving up from the left else right
        BTreeKey.Node(Pivot).Balance = BTreeKey.Node(Pivot).Balance + 1
    Else
        BTreeKey.Node(Pivot).Balance = BTreeKey.Node(Pivot).Balance - 1
    End If
    
    Select Case BTreeKey.Node(Pivot).Balance
    Case 0: Exit Function ' balanced suBTreeKey, no action needed
    Case 1, -1 ' shift Result and add this node's direction
        LocalInsert = (Result And &HFFF) * &H10 Or dir
        Exit Function
    Case Else: ' 2,-2 requires rebalancing
    End Select
    
    With BTreeKey
    
        If (Result And &HF) = 1 Then insParent = .Node(Pivot).LeftChild Else insParent = .Node(Pivot).RightChild
        
        Select Case (Result And &HFF)
        Case &H11 ' left rotation
            .Node(Pivot).LeftChild = .Node(insParent).RightChild
            .Node(insParent).RightChild = Pivot
            If dir = 1 Then .Node(Parent).LeftChild = insParent Else .Node(Parent).RightChild = insParent
            .Node(Pivot).Balance = 0
            .Node(insParent).Balance = 0
        Case &H12 ' right then left rotations
            Child = .Node(insParent).LeftChild
            .Node(Pivot).RightChild = .Node(Child).LeftChild
            .Node(insParent).LeftChild = .Node(Child).RightChild
            If dir = 1 Then .Node(Parent).LeftChild = Child Else .Node(Parent).RightChild = Child
            .Node(Child).LeftChild = Pivot
            .Node(Child).RightChild = insParent
            .Node(Child).Balance = 0
            .Node(insParent).Balance = 0
            .Node(Pivot).Balance = 0
            Result = Result And &HF00
            If Result = &H100 Then
                .Node(insParent).Balance = -1
            ElseIf Result = &H200 Then
                .Node(Pivot).Balance = 1
            End If
        Case &H21 ' left then right rotations
            Child = .Node(insParent).RightChild
            .Node(insParent).RightChild = .Node(Child).LeftChild
            .Node(Pivot).LeftChild = .Node(Child).RightChild
            If dir = 1 Then .Node(Parent).LeftChild = Child Else .Node(Parent).RightChild = Child
            .Node(Child).LeftChild = insParent
            .Node(Child).RightChild = Pivot
            .Node(Child).Balance = 0
            .Node(insParent).Balance = 0
            .Node(Pivot).Balance = 0
            Result = Result And &HF00
            If Result = &H100 Then
                .Node(Pivot).Balance = -1
            ElseIf Result = &H200 Then
                .Node(insParent).Balance = 1
            End If
        Case &H22 ' right rotation
            .Node(Pivot).RightChild = .Node(insParent).LeftChild
            .Node(insParent).LeftChild = Pivot
            If dir = 1 Then .Node(Parent).LeftChild = insParent Else .Node(Parent).RightChild = insParent
            .Node(Pivot).Balance = 0
            .Node(insParent).Balance = 0
        End Select
    End With

EH:
    If Err Then ' possible error? Out of Memory maybe
        Err.Clear
        LocalInsert = -1&
    End If
End Function

Private Function LocalDelete(Key As String, Index As Long) As Boolean

    ' Deleting a node is relatively simple. Another node slides into the deleted node's position.
    ' Tracking and balancing deletions is far more complicated than insertions. Why?  With insertions
    ' only, you are always inserting at the lowest level of a suBTreeKey. With deletions, the deleted node
    ' can come from anywhere in the tree and may actually involve additional nodes not near the node
    ' to be deleted (Case 3 deletions described below).
    
    ' However, when maintaining height-balanced trees, it gets even more complicated because of
    ' two primary reasons when the deletion causes tree to become unbalanced:
    '   1) We navigated in one direction to delete the node, but rebalancing needs to come
    '      from the other direction which has not been navigated
    '   2) We also need to adjust balance factors for both sides after deletion & rebalancing

    Dim Parent As Long, Child As Long
    Dim GrandChild As Long, Successor As Long
    Dim pathCount As Long, cmpResult As Long
    Dim cTraverseDir(1 To 35) As Byte ' used to cache path direction for balance factor updating
    Dim cTraversePtr(0 To 34) As Long ' used to cache node per path step for balance factor updating
    ' the above arrays are max'd out for the largest possible tree size 2^32
    
    If BTreeKey.Count = 0& Then Exit Function      ' nothing to delete
    If BTreeKey.Count = 1& Then
        Me.Clear
        LocalDelete = True
        Exit Function
    End If
    
    If Key = vbNullString Then Key = BTreeKey.Node(Index).Key
    
    Child = BTreeKey.Node(0).LeftChild             ' root index
    pathCount = 1&: cTraverseDir(pathCount) = 1 ' path begins by moving left
    
    Do Until Child = 0&                         ' find the node based on passed value
        cmpResult = StrComp(Key, BTreeKey.Node(Child).Key, KeyCompareBinary)
        If cmpResult = 0& Then Exit Do
        cTraversePtr(pathCount) = Child         ' set pathcount's Node ID
        Parent = Child                          ' set parent, child is next in path
        pathCount = pathCount + 1&              ' increment pathcount & set its direction
        If cmpResult = -1& Then
            cTraverseDir(pathCount) = 1         ' next path is left
            Child = BTreeKey.Node(Parent).LeftChild
        Else
            cTraverseDir(pathCount) = 2         ' next path is right
            Child = BTreeKey.Node(Parent).RightChild
        End If
    Loop
    If Child = 0& Then Exit Function            ' requested value not found in tree; abort
    
    ' When deletion occurs, there are 3 specific cases that must be addressed; from easiest to hardest
    ' Case 1: The deleted node has no right child; left child (if any) replaces deleted node
    ' Case 2: The deleted node's right child has no left child; deleted node's right child replaces deleted node
    ' Case 3: The deleted node's right child has both left & right children
    '       This requires finding either the In-Order successor or predecessor and replacing the deleted node
    '       with the successor/predecessor. This routine uses successor vs predecessor - personal choice
    '       What is an In-Order successor? The next node that has a value > deleted node, but less than the
    '       deleted node's right child's value. This is performed by moving to the deleted node's right, then
    '       from that node, navigating only left-side decendents until last one is found.
    With BTreeKey
        If .Node(Child).RightChild = 0& Then        ' deleted node has no right child: Case 1
            If cTraverseDir(pathCount) = 1 Then     ' deleted node is to Parent's left
                .Node(Parent).LeftChild = .Node(Child).LeftChild
            Else
                .Node(Parent).RightChild = .Node(Child).LeftChild
            End If
        Else                                        ' deleted node has right child, does it's child have a right child?
            GrandChild = .Node(Child).RightChild
            cTraversePtr(pathCount) = Child
            cTraversePtr(0&) = pathCount            ' cache child's position in array
            pathCount = pathCount + 1&: cTraverseDir(pathCount) = 2 ' next potential path direction is right
            If .Node(GrandChild).LeftChild = 0& Then ' no left great grandchild: Case 2
                ' simply slide the deleted node's right child into the deleted node's position
                .Node(GrandChild).LeftChild = .Node(Child).LeftChild ' slide into position
                .Node(GrandChild).Balance = .Node(Child).Balance     ' carryover the balance factor
                If cTraverseDir(cTraversePtr(0&)) = 1 Then                ' deleted node is to Parent's left
                    .Node(Parent).LeftChild = GrandChild
                Else '
                    .Node(Parent).RightChild = GrandChild
                End If
                cTraversePtr(cTraversePtr(0&)) = GrandChild ' in the recursion array, replace deleted node too
            Else    ' Case 3 deletion
                Do  ' find successor
                    cTraversePtr(pathCount) = GrandChild
                    pathCount = pathCount + 1&: cTraverseDir(pathCount) = 1
                    Successor = BTreeKey.Node(GrandChild).LeftChild
                    If BTreeKey.Node(Successor).LeftChild = 0& Then Exit Do
                    GrandChild = Successor
                Loop
                .Node(GrandChild).LeftChild = .Node(Successor).RightChild ' replace deleted node with its In-Order successor
                .Node(Successor).LeftChild = .Node(Child).LeftChild
                .Node(Successor).RightChild = .Node(Child).RightChild
                .Node(Successor).Balance = .Node(Child).Balance           ' carryover balance factor
                If cTraverseDir(cTraversePtr(0&)) = 1 Then                ' deleted node is to Parent's left
                    .Node(Parent).LeftChild = Successor
                Else
                    .Node(Parent).RightChild = Successor
                End If
                cTraversePtr(cTraversePtr(0&)) = Successor                ' in the recursion array, replace deleted node too
            End If
            cTraversePtr(0&) = 0&                                         ' reset
        End If
        .Node(0).RightChild = Child
        .Count = .Count - 1&                                              ' decrement number of nodes in tree
    End With
    
    ' the above, minus the pathtracking, is all that is needed to delete a node from a binary tree
    ' However, since we want to maintain height balancing, path tracking was performed as will
    ' be the following section which rebalances/shifts suBTreeKeys & updates balance factors
    
    For pathCount = pathCount - 1 To 1 Step -1
        If cTraverseDir(pathCount + 1) = 1 Then                    ' deletion came from the left
            BTreeKey.Node(cTraversePtr(pathCount)).Balance = BTreeKey.Node(cTraversePtr(pathCount)).Balance - 1
        Else                                                       ' deletion came from the right
            BTreeKey.Node(cTraversePtr(pathCount)).Balance = BTreeKey.Node(cTraversePtr(pathCount)).Balance + 1
        End If
        Select Case BTreeKey.Node(cTraversePtr(pathCount)).Balance
        Case 1, -1: Exit For    ' no lower level was removed; tree still in balance; done
        Case 0&                 ' tree may have become unbalanced; keep adjusting balance up the tree
        Case Else               ' tree is now unbalanced
            With BTreeKey
                Parent = cTraversePtr(pathCount - 1&)              ' get parent of unbalanced node
                Child = cTraversePtr(pathCount)                    ' cache balanced node
                
                If cTraverseDir(pathCount + 1) = 1 Then            ' deletion came from left; rebalance from Right
                ' Knowing the tree was height balanced & using the Balance factor's we can make some presumptions
                ' Based on the Child's right child (GrandChild) balance value
                '   Balance Factor = -1: GrandChild's right side has more levels than it's left. Use GrandChild to shift from
                '   Balance Factor =  1: GrandChild's left side has more levels than it's right. Use Grandchild's Left to shift from
                '   Balance Factor =  0: GrandChild has same amount of left/right levels (may be zero). Use GrandChild to shift from
                ' Knowing this, we either rebalance from the GrandChild or its left child
                    
                    GrandChild = .Node(Child).RightChild           ' cache deleted node's right child; will always have one
                    If .Node(GrandChild).Balance = 1 Then          ' rebalance from GrandChild's Right to avoid additional rebalancing
                        Successor = .Node(GrandChild).LeftChild    ' cache GrandChild's left; this will be the node that is inserted before unbalanced node
                        .Node(GrandChild).LeftChild = .Node(Successor).RightChild   ' do the shifting
                        .Node(Successor).RightChild = .Node(Child).RightChild
                        .Node(Child).RightChild = .Node(Successor).LeftChild
                        .Node(Successor).LeftChild = Child
                        Select Case .Node(Successor).Balance       ' fix up the balance factors
                            Case 0&
                                .Node(Child).Balance = 0&
                                .Node(GrandChild).Balance = 0&
                            Case 1&
                                .Node(Child).Balance = 1&
                                .Node(GrandChild).Balance = 0&
                            Case -1&
                                .Node(Child).Balance = 0&
                                .Node(Successor).Balance = 0&
                                .Node(GrandChild).Balance = -1&
                        End Select
                        GrandChild = Successor                     ' set for parent assignment & exit comparison
                    Else                                           ' rebalance off the GrandChild vs GrandChild's Left
                        If .Node(GrandChild).Balance = 0& Then     ' fix up the balance factors
                            .Node(GrandChild).Balance = 1&
                            .Node(Child).Balance = -1
                        Else
                            .Node(GrandChild).Balance = 0&
                            .Node(Child).Balance = 0
                        End If                                     ' do the shifting
                        .Node(Child).RightChild = .Node(GrandChild).LeftChild
                        .Node(GrandChild).LeftChild = Child
                    End If
    
                Else                                               ' rebalance from right (opposite from above)
                    ' Knowing the tree was height balanced & using the Balance factor's we can make some presumptions
                    ' Based on the Child's left child (GrandChild) balance value
                    '   Balance Factor = -1: GrandChild's right side has more levels than it's left. Use Grandchild's Right to shift from
                    '   Balance Factor =  1: GrandChild's left side has more levels than it's right. Use GrandChild to shift from
                    '   Balance Factor =  0: GrandChild has same amount of left/right levels (may be zero). Use GrandChild to shift from
                    ' Knowing this, we either rebalance from the GrandChild or its right child
                    GrandChild = .Node(Child).LeftChild
                    If .Node(GrandChild).Balance = -1 Then
                        Successor = .Node(GrandChild).RightChild
                        .Node(GrandChild).RightChild = .Node(Successor).LeftChild
                        .Node(Successor).LeftChild = .Node(Child).LeftChild
                        .Node(Child).LeftChild = .Node(Successor).RightChild
                        .Node(Successor).RightChild = Child
                        Select Case .Node(Successor).Balance
                            Case 0&
                                .Node(Child).Balance = 0&
                                .Node(GrandChild).Balance = 0&
                            Case 1&
                                .Node(Child).Balance = -1&
                                .Node(GrandChild).Balance = 0&
                            Case -1&
                                .Node(Child).Balance = 0&
                                .Node(Successor).Balance = 0&
                                .Node(GrandChild).Balance = 1&
                        End Select
                        GrandChild = Successor
                    Else
                        If .Node(GrandChild).Balance = 0& Then
                            .Node(GrandChild).Balance = -1&
                            .Node(Child).Balance = 1
                        Else
                            .Node(GrandChild).Balance = 0&
                            .Node(Child).Balance = 0
                        End If
                        .Node(Child).LeftChild = .Node(GrandChild).RightChild
                        .Node(GrandChild).RightChild = Child
                    End If
                End If
                If cTraverseDir(pathCount) = 1 Then                ' assign parent the shifted node as its child
                    .Node(Parent).LeftChild = GrandChild
                Else
                    .Node(Parent).RightChild = GrandChild
                End If
            End With
            If BTreeKey.Node(GrandChild).Balance Then Exit For        ' if the shifted node lost a level, keep checking else abort
        End Select
    Next
    
    ' Array management
    On Error Resume Next
    With BTreeKey   ' swap the last node in array with the one just deleted
        If .Node(0).RightChild <= .Count Then
            ' need to find the last node in the array first
            LocalFind .Node(.Count + 1).Key, .Node(0).LeftChild + 0&, Parent
            .Node(.Node(0).RightChild) = .Node(.Count + 1)
            If .Node(Parent).RightChild > .Count Then
                .Node(Parent).RightChild = .Node(0).RightChild
            Else
                .Node(Parent).LeftChild = .Node(0).RightChild
            End If
            .Node(.Count + 1).Key = vbNullString
        Else
            .Node(.Node(0).RightChild).Key = vbNullString
        End If
    End With
    If UBound(BTreeKey.Node) > BTreeKey.Count + 20& Then
        ReDim Preserve BTreeKey.Node(0 To BTreeKey.Count + 5)
    End If
    If Err Then Err.Clear
    LocalDelete = True
    
End Function

Private Function LocalFind(Key As String, FromIndex As Long, Optional ParentIndex As Long) As Long
    
    ' Returns the index of a node and optionally the parent node (LocalDelete uses the ParentIndex)
    Dim cmpResult As Long
    ParentIndex = 0&
    Do
        cmpResult = StrComp(Key, BTreeKey.Node(FromIndex).Key, KeyCompareBinary)
        If cmpResult = 0& Then Exit Do
        ParentIndex = FromIndex
        If cmpResult = -1& Then
            FromIndex = BTreeKey.Node(FromIndex).LeftChild
        Else
            FromIndex = BTreeKey.Node(FromIndex).RightChild
        End If
    Loop Until FromIndex = 0&
    LocalFind = FromIndex
    
End Function


Private Function CreateNode(Parent As Long, LeftChild As Boolean, Key As String, Value As Long) As Long
    
    ' Returns the index of the newly created node
    If BTreeKey.Count = 0& Then
        ReDim BTreeKey.Node(0 To 10)
    ElseIf BTreeKey.Count = UBound(BTreeKey.Node) Then
        ReDim Preserve BTreeKey.Node(0 To BTreeKey.Count + 10)
    End If
                
    BTreeKey.Count = BTreeKey.Count + 1&
    If LeftChild Then
        BTreeKey.Node(Parent).LeftChild = BTreeKey.Count
    Else
        BTreeKey.Node(Parent).RightChild = BTreeKey.Count
    End If
    With BTreeKey.Node(BTreeKey.Count)
        .LeftChild = 0&
        .RightChild = 0&
        .Balance = 0&
        .Value = Value
        .Key = Key
    End With
    CreateNode = BTreeKey.Count
    
End Function

Private Function ValidateIndex(Index As Variant, Optional FromIndex As Long) As Long
    ' Returns the Index of the passed Key or simply the index if Index is numeric
    ' Raises an error if Index is invalid
    If BTreeKey.Count Then
        If VarType(Index) = vbString Then
            ValidateIndex = LocalFind(CStr(Index), Me.NodeRootIndex, FromIndex)
        Else
            On Error GoTo EH
            If CLng(Index) = 0& Then
                ValidateIndex = BTreeKey.Node(0).LeftChild
            Else
                If Index > 0& And Index <= BTreeKey.Count Then
                    ValidateIndex = Index
                Else
                    Err.Raise 380
                End If
            End If
        End If
    End If
EH:
    If Err Then Err.Raise Err.Number, "cBinaryTree"
End Function


Private Sub FYI_TraversalRoutines()
'   The following are here only to keep them with the class.
'   Should you need to traverse the tree, you can copy & paste one of
'   these routines to your project and unrem the lines

' From within your project, assuming your tree class is called myBinaryTree....

'    Call InOrderAsc(myBinaryTree.NodeRootIndex)
'    Private Sub InOrderAsc(Index As Long)
'        If Index Then
'            Dim rNode As Long
'            rNode = myBinaryTree.NodeChild(Index, True)
'            Call InOrderAsc(rNode)
'            ' maybe add key or value to a listbox?
'            List1.AddItem myBinaryTree.NodeKey(Index)
'            rNode = myBinaryTree.NodeChild(Index, False)
'            Call InOrderAsc(rNode)
'        End If
'    End Sub


'    Call InOrderDesc(myBinaryTree.NodeRootIndex)
'    Private Sub InOrderDesc(Index As Long)
'        If Index Then
'            Dim rNode As Long
'            rNode = myBinaryTree.NodeChild(Index, False)
'            Call InOrderDesc(rNode)
'            ' maybe add key or value to a listbox?
'            List1.AddItem myBinaryTree.NodeKey(Index)
'            rNode = myBinaryTree.NodeChild(Index, True)
'            Call InOrderDesc(rNode)
'        End If
'    End Sub


'    Call PreOrder(myBinaryTree.NodeRootIndex)
'    Private Sub PreOrder(Index As Long)
'        Dim rNode As Long
'        If Index Then
'            ' maybe add key or value to a listbox?
'            List1.AddItem myBinaryTree.NodeKey(Index)
'            rNode = myBinaryTree.NodeChild(Index, True)
'            Call PreOrder(rNode)
'            rNode = myBinaryTree.NodeChild(Index, False)
'            Call PreOrder(rNode)
'        End If
'    End Sub


'    Call PostOrder(myBinaryTree.NodeRootIndex)
'    Private Sub PostOrder(Index As Long)
'        If Index Then
'            Dim rNode As Long
'            rNode = myBinaryTree.NodeChild(Index, True)
'            Call PostOrder(rNode)
'            rNode = myBinaryTree.NodeChild(Index, False)
'            Call PostOrder(rNode)
'            ' maybe add key or value to a listbox?
'            List1.AddItem myBinaryTree.NodeKey(Index)
'        End If
'    End Sub

End Sub
