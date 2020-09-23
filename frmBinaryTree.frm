VERSION 5.00
Begin VB.Form frmBinaryTree 
   Caption         =   "Self-Balancing Binary Tree"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Height Balanced?"
      Height          =   585
      Left            =   9255
      TabIndex        =   4
      Top             =   2850
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0024CACE&
      FillColor       =   &H0024CACE&
      Height          =   2400
      Left            =   90
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   693
      TabIndex        =   1
      Top             =   315
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Tree"
      Height          =   585
      Left            =   165
      TabIndex        =   0
      Top             =   2820
      Width           =   2550
   End
   Begin VB.Label Label1 
      Caption         =   "Furthest Right on right side is highest sorted Index."
      Height          =   285
      Index           =   1
      Left            =   3735
      TabIndex        =   3
      Top             =   3150
      Width           =   4185
   End
   Begin VB.Label Label1 
      Caption         =   "Furthest Left on left side is lowest sorted Index."
      Height          =   285
      Index           =   0
      Left            =   3735
      TabIndex        =   2
      Top             =   2910
      Width           =   4380
   End
   Begin VB.Label Label1 
      Caption         =   "Left click on nodes for information and options, double dclick node to delete it."
      Height          =   285
      Index           =   2
      Left            =   225
      TabIndex        =   5
      Top             =   90
      Width           =   5850
   End
   Begin VB.Menu mnuPU 
      Caption         =   "popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Balance"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Value"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Delete this Node"
         Index           =   3
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Insert a New Node"
         Index           =   4
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Change Node Key"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmBinaryTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NODEREF
    X As Long
    y As Long
    ID As Long
End Type
Private mMouseHotSpots(1 To 26) As NODEREF ' node locations on display
Private Const NodeCx As Long = 20

Private myBinaryTree As cBinaryTree

Private Sub Command1_Click()
        
    ' create a binary tree of random height
    Dim X As Long
    If myBinaryTree Is Nothing Then Set myBinaryTree = New cBinaryTree
    
    X = myBinaryTree.NodeCount
    Do Until Abs(X - myBinaryTree.NodeCount)
        X = Int(Rnd * 24) + 3
    Loop
    
    myBinaryTree.Clear
    
    For X = 65 To X + 64
        myBinaryTree.Add Chr$(X), Int(Rnd * vbWhite), False
    Next
    RenderTree
    
End Sub

Private Sub Command2_Click()
    Dim sMsg As String
    sMsg = "A tree is height-balanced if the heights of the left and right subtree's of each node are within 1." & vbCrLf & _
        "A tree is perfectly balanced if the left and right subtrees of any node are the same height."
    MsgBox sMsg, vbOKOnly Or vbInformation, "Definition"
End Sub

Private Sub Form_Load()
    Dim X As Long
    With Picture1
        .AutoRedraw = True
        .Font.Size = 8
        .Font.Bold = True
        .FillStyle = vbSolid
        .BackColor = &H808000
    End With
    Show
    Randomize
    Call Command1_Click
    MsgBox "The tree display is limited to only 6 rows, the actual tree's limitation is 2^32-1 nodes"

End Sub
Private Sub RenderTree()
    
    ' a crude display of the tree
    ' The display is limited to 6 rows of nodes, simply because trying to display
    ' more would result in very large display areas. Too large to display on screen.
    
    Dim Colors(0 To 7)      ' node colors per row
    Colors(0) = vbWhite
    Colors(1) = vbBlack
    Colors(2) = vbBlue
    Colors(3) = vbMagenta
    Colors(4) = vbRed
    Colors(5) = RGB(92, 92, 92)
    Colors(6) = RGB(128, 128, 128)
    'Colors(7) never used
    
    Dim Ht As Long, Spacer As Long
    Dim X As Long, y As Long
    Dim Row As Long, Col As Long
    Dim RightLeft As Long
    Dim Index As Long, NodeCount As Long
    Dim theOrder() As Long
    
    If myBinaryTree.NodeCount = 0& Then
        Picture1.Cls
        Exit Sub
    End If
    
    Ht = 6 ' max nr of rows to display
    
    
    ' reset forecolor
    Picture1.ForeColor = Colors(0)
    
    ' resize array & get the nodes in a specific order
    ' The order for display purposes is a modified BinaryTree InOrder traverse
    ' If tree looked like this: ( _ = no node)
    '         A (root is never returned in the array)
    '      B
    '   C    D
    ' E  F     G
    ' the returned order would be: B, C, D, E, F, _, G
    
    ReDim theOrder(0 To 2 ^ (Ht + 1) + 1)
    ' prime array with the root node's left child
    theOrder(1) = myBinaryTree.NodeChild(myBinaryTree.NodeRootIndex, True)
    
    Erase mMouseHotSpots()
    NodeCount = 1
    Picture1.Cls
    For RightLeft = 0 To 1
        If RightLeft = 1 Then ' do the right half of the tree
            ReDim theOrder(0 To UBound(theOrder))
            theOrder(1) = myBinaryTree.NodeChild(myBinaryTree.NodeRootIndex, False)
        End If
        theOrder(0) = 1 ' indicate how many nodes are in array
        BuildPreOrderArray theOrder, 1, 1 ' call recursion function to fill array
        theOrder(theOrder(0) + 1) = -1       ' identify stopping point
        ' initiailize which array member to start with, which display row & col to start with
        Index = 1: Col = 1: Row = 2
        Do
            Spacer = ((2 ^ (Ht - Row) - 0) * NodeCx) ' spacing between nodes for this row
            X = (Spacer \ 2) + Spacer * ((2 ^ (Row - 2)) * RightLeft) ' first node's X coordinate
            y = Row * NodeCx                         ' nodes Y coordinate
            Picture1.FillColor = Colors(Row - 1) ' set circle's backcolor
            ' loop thru the row's nodes
            For Index = Index To Index + Col - 1
                If theOrder(Index) Then ' if zero, don't display circle
                    ' draw circle, draw node Index value
                    Picture1.Circle (X + NodeCx \ 2, y + NodeCx \ 2), NodeCx \ 2
                    Picture1.CurrentX = X + 4: Picture1.CurrentY = y + 3
                    Picture1.Print myBinaryTree.NodeKey(theOrder(Index))
                    NodeCount = NodeCount + 1
                    mMouseHotSpots(NodeCount).X = X
                    mMouseHotSpots(NodeCount).y = y
                    mMouseHotSpots(NodeCount).ID = theOrder(Index)
                End If
                X = X + Spacer  ' set X coordinate for next node
            Next
            If theOrder(Index) = -1 Then Exit Do ' if exit flag encountered, exit loop
            Col = Col * 2       ' set position in array for next row's nodes
            Row = Row + 1       ' set next row
        Loop
    Next
    ' draw the root node & add extra info to the display
    X = Picture1.ScaleWidth \ 2
    For RightLeft = 0 To 0
        y = NodeCx \ 2
        Picture1.FillColor = Colors(0)
        Picture1.Circle (X + NodeCx \ 2, y + NodeCx \ 2), NodeCx \ 2
        Picture1.ForeColor = Colors(1)
        Picture1.CurrentX = X + 4: Picture1.CurrentY = y + 3
        Picture1.Print myBinaryTree.NodeKey(myBinaryTree.NodeRootIndex);
        Picture1.ForeColor = Colors(0)
        Picture1.Print "      Root"
        mMouseHotSpots(1).X = X
        mMouseHotSpots(1).y = y
    Next
End Sub

Private Sub BuildPreOrderArray(outArray() As Long, Index As Long, Count As Long)
    
    ' recursive routine for the RenderTree function
    
    Dim bFilled As Boolean
    For Index = Index To Index + Count - 1
        outArray(0) = outArray(0) + 1   ' (0) element tracks array entries
        If outArray(Index) Then         ' add non-zero node Indexes (Left children)
            outArray(outArray(0)) = myBinaryTree.NodeChild(outArray(Index), True)
            bFilled = True
        End If
        outArray(0) = outArray(0) + 1   ' (0) element tracks array entries
        If outArray(Index) Then         ' add non-zero node Indexes (Right children)
            outArray(outArray(0)) = myBinaryTree.NodeChild(outArray(Index), False)
            bFilled = True
        End If
    Next
    If bFilled Then ' do next row
        BuildPreOrderArray outArray(), Index, Count * 2
    Else            ' abort, reset to position after last entry
        outArray(0) = outArray(0) - Count * 2
    End If
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Select Case Index
    Case 3 ' delete
        myBinaryTree.Delete mnuPU.Tag
        RenderTree
    Case 4, 5 ' append, rekey
        Dim Key As String
        Key = LTrim$(InputBox("Enter a Key (A-Z), duplicates are not permitted", "New Node"))
        If Len(Key) Then
            If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase$(Key)) Then
                On Error GoTo EH
                If Index = 4 Then ' append
                    If myBinaryTree.Add(UCase$(Key), Int(Rnd * vbWhite), False) Then RenderTree
                Else
                    If myBinaryTree.ReKey(mnuPU.Tag, UCase$(Key)) Then RenderTree
                End If
            Else
                MsgBox "Invalid Key for this sample project. Enter a new key between A-Z", vbInformation + vbOKOnly, "Invalid Key"
            End If
        End If
    End Select
EH:
    If Err Then
        MsgBox "Error: " & Err.Description, vbOKOnly, Err.Source
        Err.Clear
    End If
End Sub

Private Sub Picture1_DblClick()
    ' allow dbl clicking on a node to delete the node
    If Len(mnuPU.Tag) Then
        myBinaryTree.Delete mnuPU.Tag
        mnuPU.Tag = vbNullString
        RenderTree
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' simple method to determine if mouse clicked in node or not
    Dim I As Long
    mnuPU.Tag = vbNullString
    For I = 1 To 26
        If X > mMouseHotSpots(I).X Then
            If X < mMouseHotSpots(I).X + NodeCx Then
                If y > mMouseHotSpots(I).y Then
                    If y < mMouseHotSpots(I).y + NodeCx Then
                        mnuPU.Tag = myBinaryTree.NodeKey(mMouseHotSpots(I).ID)
                        If Button = vbRightButton Then
                            mnuPopup(0).Caption = "Balance Factor is " & myBinaryTree.NodeBalance(mMouseHotSpots(I).ID)
                            mnuPopup(1).Caption = "Value is " & myBinaryTree.NodeValue(mMouseHotSpots(I).ID)
                            PopupMenu mnuPU, , , , mnuPopup(3)
                            mnuPU.Tag = vbNullString
                        End If
                        Exit For
                    End If
                End If
            End If
        End If
    Next
End Sub
