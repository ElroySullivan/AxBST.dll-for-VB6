VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BST_LngKey_VarVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Version 2.00.     by Elroy Sullivan, PhD
'
' See notes in VbForums.Com (or associated RTF file) on how to use this:
'   https://www.vbforums.com/showthread.php?903908-Homegrown-Self-Balancing-Binary-Search-Tree-(BST)
'
Option Explicit
Option Compare Binary                               ' The default but just to be explicit.
'
' This variable (set in Class_Initialize) determines what the "NoKey" indicator value should be.
' If possible 0& is the best one used.  However, in some situations, we may want the 0& as a valid key.
' Another good alternative is &h80000000, as that's a number very unlikely to be used.
' It's set during design-time because it can't be changed once a tree is being used.
Private mNoKey As Long
'
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (a() As Any) As Long
Private Declare Sub SafeArrayAllocDescriptor Lib "oleaut32.dll" (ByVal cDims As Long, ByRef psaInOut As Long)
'
Public Enum BstAscendDescend_LngKey_VarVal
    bstAscend_LngKey_VarVal
    bstDescend_LngKey_VarVal
End Enum
'
' Things for NextKey & PrevKey to work.
Private IterateBranch(40&)      As Long             ' 40 is deeper than the tree can ever possibly get.
Private IterateLeftRigt(40&)    As Integer          ' 40 is deeper than the tree can ever possibly get.
Private BranchNodeCount         As Long             ' Always the NEXT one available (not the last one used).
Private Iterating               As Boolean
'
' Our node type and stuff for the index management.
Private Type NodeType
    Key         As Long
    Value       As Variant
    LeftChild   As Long
    RigtChild   As Long
    Height      As Integer
    Padding     As Integer      ' This will be here, whether explicit or implicit, so may as well be explicit.
End Type
'
Private Type GroupsType
    Nodes() As NodeType
End Type
Private Groups()                As GroupsType
Private Const NODES_PER_GROUP   As Long = 20&       ' This effectively tells us how often we ReDim when inserting (which takes time).  But too large wastes memory.
Private DeletedIndexes()        As Long
Private Const DELETED_BUMP      As Long = 100&
Private miDeletedIndexCount     As Long
'
' The root and count.
Private pRoot           As Long
Private miNodeCount     As Long
'
' It's important to note that "Index" values are ONE-based.
' This is done because it's much easier to have ZERO indicate that no index has been assigned.
'


Private Sub Class_Initialize()
    mNoKey = 0&     ' See note where this is declared for more explanation.
    Clear
End Sub
Public Property Get NoKey() As Long
    NoKey = mNoKey  ' We create this property to make sure that mNoKey is read-only during runtime.
End Property

' *******************************************************************
' *******************************************************************
'
'   Exposed production procedures.
'
' *******************************************************************
' *******************************************************************

Public Function Insert(Key As Long, Value As Variant, Optional ChangeValueIfExists As Boolean) As Boolean
    ' Returns TRUE if successful.
    ' FALSE indicates we attempted to add a duplicate.
    '
    If Key = mNoKey Then Exit Function
    Dim pNode As Long
    pNode = FindNodeRecurse(pRoot, Key)
    If pNode <> 0& Then
        If ChangeValueIfExists Then
            LetSet Groups(Group(pNode)).Nodes(GNode(pNode)).Value, Value
            Insert = True                           ' If we're not changing value, leave return as False.
        End If
        Exit Function
    End If
    '
    If Iterating Then CancelIterating               ' Because of stored iteration branch, we can't continue iterating after insert/delete.
    pRoot = InsertRecurse(pRoot, Key, Value)        ' Add the key.
    Insert = True
End Function

Public Function Delete(Key As Long) As Boolean
    ' Returns TRUE if successful.
    ' FALSE indicates we couldn't find the key.
    '
    If Not KeyExists(Key) Then Exit Function        ' Because of the way the DeleteRecurse works, we're forced to do this here.
    If Iterating Then CancelIterating               ' Because of stored iteration branch, we can't continue iterating after insert/delete.
    pRoot = DeleteRecurse(pRoot, Key)               ' Delete the key.
    Delete = True
End Function

Public Property Get Value(Key As Long) As Variant
    ' Returns Empty if the key doesn't exist.
    ' But, even if the key does exist, the value could still be Empty.
    ' So, if this is a concern, check first with KeyExists.
    '
    Dim pNode As Long
    pNode = FindNodeRecurse(pRoot, Key)
    If pNode Then LetSet Value, Groups(Group(pNode)).Nodes(GNode(pNode)).Value
End Property

Public Property Let Value(Key As Long, Value As Variant)
    ' This provides a way to change values without requiring Delete/Insert.
    ' It does nothing if the key doesn't exist.
    '
    Dim pNode As Long
    pNode = FindNodeRecurse(pRoot, Key)
    If pNode Then Groups(Group(pNode)).Nodes(GNode(pNode)).Value = Value        ' We don't use LetSet on this one.
End Property

Public Property Set Value(Key As Long, Value As Variant)
    ' This provides a way to change values without requiring Delete/Insert.
    ' It does nothing if the key doesn't exist.
    '
    Dim pNode As Long
    pNode = FindNodeRecurse(pRoot, Key)
    If pNode Then Set Groups(Group(pNode)).Nodes(GNode(pNode)).Value = Value    ' We don't use LetSet on this one.
End Property

Public Function KeyExists(Key As Long) As Boolean
    KeyExists = ExistsRecurse(pRoot, Key)
End Function

Public Function Count() As Long
    Count = miNodeCount
End Function

Public Sub Clear()
    ' Just clears all nodes out of the tree.
    '
    CancelIterating
    pRoot = 0&
    miNodeCount = 0&
    miDeletedIndexCount = 0&
    '
    ReDim Groups(0&)
    ReDim Groups(0&).Nodes(NODES_PER_GROUP - 1&)
    ReDim DeletedIndexes(DELETED_BUMP - 1&)
End Sub

Public Sub Compact()
    ' It's not necessary to do this, and the only time it's useful is when
    ' there are a lot of deletions without insertions to follow.
    ' If insertions are to follow, they'll fill-in the deletions.
    ' Calling this resets any NextKey/PrevKey.
    '
    ' Make sure there's a need.
    If miDeletedIndexCount = 0& Then Exit Sub
    '
    ' Save a local copy of all the keys.
    Dim LocalKeys() As Long:        LocalKeys = KeyArray
    Dim LocalValues() As Variant:   LocalValues = ValueArray
    '
    ' Reset everything.
    Clear
    '
    ' Rebuild everything.
    Dim i As Long
    For i = 0& To UBound(LocalKeys)
        pRoot = InsertRecurse(pRoot, LocalKeys(i), LocalValues(i))
    Next
End Sub

Public Function KeyArray(Optional AscendDescend As BstAscendDescend_LngKey_VarVal = bstAscend_LngKey_VarVal) As Long()
    ' Returns zero-based SORTED array of all the keys.
    ' Returns 0 to -1 array if the tree is empty.
    '
    BuildKeyArray KeyArray, AscendDescend
End Function

Public Function ValueArray(Optional AscendDescend As BstAscendDescend_LngKey_VarVal = bstAscend_LngKey_VarVal) As Variant()
    ' Returns zero-based array of all values, SORTED based on the keys.
    ' If both KeyArray and ValueArray are fetched, they will be synchronized.
    ' Returns 0 to -1 array if the tree is empty.
    '
    BuildValueArray ValueArray, AscendDescend
End Function

Public Function NextKey() As Long
    ' This is slower than getting the KeyArray and then looping, but sometimes this way is desired.
    ' mNoKey is returned when all the keys have been iterated.  Calling again will start over.
    ' This gets reset on any Insert/Delete.
    ' Call the Reset method to start over.
    '
    NextKey = mNoKey
    If miNodeCount = 0& Then Exit Function
    '
    Dim pThis As Long                                               ' Always temporary reference of existing node, and just falls out of scope.
    If Not Iterating Then
        Iterating = True
        ' Crawl down to bottom of left branch, and return that key, accumulating branch as we go.
        pThis = pRoot
        IterateBranch(BranchNodeCount) = pThis
        ' IterateLeftRigt(0) is always 0 (to indicate the root).
        BranchNodeCount = BranchNodeCount + 1&
        Do While Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
            pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
            IterateBranch(BranchNodeCount) = pThis
            IterateLeftRigt(BranchNodeCount) = -1
            BranchNodeCount = BranchNodeCount + 1&
        Loop
        NextKey = Groups(Group(pThis)).Nodes(GNode(pThis)).Key
        Exit Function
    End If
    '
    ' If we get down here, we're actively iterating (past the minimum).
    ' Paths:
    '   Right, then to bottom of left.
    '   Go up until we find a parent that's "greater than" the one we're starting it.
    '   No greater than parent, we're all done.
    '
    pThis = IterateBranch(BranchNodeCount - 1&)                     ' Always the one we previously returned.
    '
    If Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild Then      ' Right, then to bottom of left.
        pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
        IterateBranch(BranchNodeCount) = pThis
        IterateLeftRigt(BranchNodeCount) = 1
        BranchNodeCount = BranchNodeCount + 1&
        Do While Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
            pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
            IterateBranch(BranchNodeCount) = pThis
            IterateLeftRigt(BranchNodeCount) = -1
            BranchNodeCount = BranchNodeCount + 1&
        Loop
        NextKey = Groups(Group(pThis)).Nodes(GNode(pThis)).Key      ' Found next one down branch.
        Exit Function
    Else
        Do
            BranchNodeCount = BranchNodeCount - 1&
            If BranchNodeCount = 0& Then                            ' We're done, no more.
                CancelIterating
                Exit Function
            End If
            pThis = IterateBranch(BranchNodeCount - 1&)
            If IterateLeftRigt(BranchNodeCount) = -1 Then           ' This parent is greater than where we were.
                NextKey = Groups(Group(pThis)).Nodes(GNode(pThis)).Key
                Exit Function
            End If
        Loop
    End If
End Function

Public Function PrevKey() As Long
    ' This is the "descending" version of NextKey.  See those comments for more explanation.
    ' As a note, you can call NextKey, and then call PrevKey to backup.  They will work together like that.
    '
    PrevKey = mNoKey
    If miNodeCount = 0& Then Exit Function
    '
    Dim pThis As Long                                               ' Always temporary reference of existing node, and just falls out of scope.
    If Not Iterating Then
        Iterating = True
        ' Crawl down to bottom of right branch, and return that key, accumulating branch as we go.
        pThis = pRoot
        IterateBranch(BranchNodeCount) = pThis
        ' IterateLeftRigt(0) is always 0 (to indicate the root).
        BranchNodeCount = BranchNodeCount + 1&
        Do While Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
            pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
            IterateBranch(BranchNodeCount) = pThis
            IterateLeftRigt(BranchNodeCount) = 1
            BranchNodeCount = BranchNodeCount + 1&
        Loop
        PrevKey = Groups(Group(pThis)).Nodes(GNode(pThis)).Key
        Exit Function
    End If
    '
    ' If we get down here, we're actively iterating (past the minimum).
    ' Paths:
    '   Left, then to bottom of right.
    '   Go up until we find a parent that's "less than" the one we're starting it.
    '   No less than parent, we're all done.
    '
    pThis = IterateBranch(BranchNodeCount - 1&)                     ' Always the one we previously returned.
    '
    If Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild Then      ' Left, then to bottom of right.
        pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
        IterateBranch(BranchNodeCount) = pThis
        IterateLeftRigt(BranchNodeCount) = -1
        BranchNodeCount = BranchNodeCount + 1&
        Do While Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
            pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
            IterateBranch(BranchNodeCount) = pThis
            IterateLeftRigt(BranchNodeCount) = 1
            BranchNodeCount = BranchNodeCount + 1&
        Loop
        PrevKey = Groups(Group(pThis)).Nodes(GNode(pThis)).Key      ' Found next one down branch.
        Exit Function
    Else
        Do
            BranchNodeCount = BranchNodeCount - 1&
            If BranchNodeCount = 0& Then                            ' We're done, no more.
                CancelIterating
                Exit Function
            End If
            pThis = IterateBranch(BranchNodeCount - 1&)
            If IterateLeftRigt(BranchNodeCount) = 1 Then            ' This parent is less than where we were.
                PrevKey = Groups(Group(pThis)).Nodes(GNode(pThis)).Key
                Exit Function
            End If
        Loop
    End If
End Function

Public Sub Reset()
    CancelIterating
End Sub

' *******************************************************************
' *******************************************************************
'
'   Exposed procedures for debugging/testing.
'
' *******************************************************************
' *******************************************************************

Friend Function TheKey(Index As Long) As Long
    If Index = 0& Then Exit Function
    TheKey = Groups(Group(Index)).Nodes(GNode(Index)).Key
End Function

Friend Function TheHeight(Index As Long) As Integer
    If Index = 0& Then Exit Function
    TheHeight = Groups(Group(Index)).Nodes(GNode(Index)).Height
End Function

Friend Function TheRoot() As Long
    TheRoot = pRoot
End Function

Friend Function TheLeftChild(Index As Long) As Long
    If Index = 0& Then Exit Function
    TheLeftChild = Groups(Group(Index)).Nodes(GNode(Index)).LeftChild
End Function

Friend Function TheRigtChild(Index As Long) As Long
    If Index = 0& Then Exit Function
    TheRigtChild = Groups(Group(Index)).Nodes(GNode(Index)).RigtChild
End Function

Friend Function FindNode(Key As Long) As Long
    ' Returns 0 if not found.
    FindNode = FindNodeRecurse(pRoot, Key)
End Function

Friend Sub DumpKeysAndValuesInOrder(Optional AscendDescend As BstAscendDescend_LngKey_VarVal = bstAscend_LngKey_VarVal, Optional pNode As Long, Optional bRecursing As Boolean)
    ' Don't ever supply pNode or bRecursing, as they're just for recursion.
    '
    If Not bRecursing Then Debug.Print "-------- In Order Tree Dump --------"
    If pRoot = 0& Then Debug.Print "NO TREE !!!": Exit Sub
    '
    ' See if we've got a node.
    If pNode = 0& Then
        If bRecursing Then
            Exit Sub                                            ' We hit a leaf.
        Else
            pNode = pRoot                                       ' We're just getting started.
        End If
    End If
    '
    Dim Value As Variant
    If AscendDescend = bstDescend_LngKey_VarVal Then
        DumpKeysAndValuesInOrder AscendDescend, Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, True    ' Go through right branches.
        LetSet Value, Groups(Group(pNode)).Nodes(GNode(pNode)).Value
        Select Case True
        Case IsObject(Value):   Debug.Print Groups(Group(pNode)).Nodes(GNode(pNode)).Key, "Value: Object type: " & TypeName(Value)
        Case IsArray(Value):    Debug.Print Groups(Group(pNode)).Nodes(GNode(pNode)).Key, "Value: Array"
        Case Else:              Debug.Print Groups(Group(pNode)).Nodes(GNode(pNode)).Key, Value
        End Select
        DumpKeysAndValuesInOrder AscendDescend, Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, True    ' Go through left branches.
    Else
        DumpKeysAndValuesInOrder AscendDescend, Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, True    ' Go through left branches.
        LetSet Value, Groups(Group(pNode)).Nodes(GNode(pNode)).Value
        Select Case True
        Case IsObject(Value):   Debug.Print Groups(Group(pNode)).Nodes(GNode(pNode)).Key, "Value: Object type: " & TypeName(Value)
        Case IsArray(Value):    Debug.Print Groups(Group(pNode)).Nodes(GNode(pNode)).Key, "Value: Array"
        Case Else:              Debug.Print Groups(Group(pNode)).Nodes(GNode(pNode)).Key, Value
        End Select
        DumpKeysAndValuesInOrder AscendDescend, Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, True    ' Go through right branches.
    End If
End Sub

Friend Function NodeCount(Optional ByVal pNode As Long, Optional bRecursing As Boolean) As Long
    ' Don't ever supply pNode or bRecursing, as they're just for recursion.
    ' This counts the nodes the "hard way", not using miNodeCount.  It should ALWAYS equal miNodeCount.
    '
    Static iCount As Long
    If pRoot = 0& Then Exit Function                                    ' NodeCount = 0.
    '
    ' See if we've got a node.
    If pNode = 0& Then
        If bRecursing Then
            Exit Function                                               ' We hit a leaf.
        Else
            pNode = pRoot                                               ' We're just getting started.
            iCount = 0&
        End If
    End If
    '
    NodeCount Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, True  ' Go through left branches.
    iCount = iCount + 1&
    NodeCount Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, True  ' Go through right branches.
    '
    If Not bRecursing Then NodeCount = iCount
End Function

' *******************************************************************
' *******************************************************************
'
'   Helper functions --- WITH recursion.
'
' *******************************************************************
' *******************************************************************

Private Function InsertRecurse(ByVal pNode As Long, Key As Long, Value As Variant) As Long
    ' Upon return of first call, the return is what should be the new tree's root.
    '
    ' Find our node, probably recursing to do it.
    If pNode = 0& Then
        pNode = AllocIndex                                              ' We found the bottom, so make a new leaf.  This increments miNodeCount.
        Groups(Group(pNode)).Nodes(GNode(pNode)).Key = Key
        LetSet Groups(Group(pNode)).Nodes(GNode(pNode)).Value, Value
        InsertRecurse = pNode                                           ' For now, it'll be our root, further adjusted below.
    Else
        Select Case Groups(Group(pNode)).Nodes(GNode(pNode)).Key
        Case Is > Key:  Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild = InsertRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, Key, Value)   ' Recurse left, keep looking for empty spot.
        Case Is < Key:  Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild = InsertRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, Key, Value)   ' Recurse right, looking for empty spot.
        Case Else:      Exit Function                                           ' Duplicate keys not allowed.  Return 0.
        End Select
    End If
    '
    ' Update height of current node.  Updating all as recursion unwinds.
    Groups(Group(pNode)).Nodes(GNode(pNode)).Height = Max(Height(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild), Height(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)) + 1
    '
    ' Balance unbalanced [Abs(Balance) > 1] node(s).  Rebalancing all as recursion unwinds.
    Select Case BalanceFactor(pNode)
    Case Is > 1     ' Note that >0 means we MUST have a left child.
        Dim pLeft As Long: pLeft = Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild
        If Key < Groups(Group(pLeft)).Nodes(GNode(pLeft)).Key Then      ' Left  Left  Case.
            InsertRecurse = RotateRight(pNode)
        Else                                                            ' Left  Right  Case.
            Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild = RotateLeft(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild)
            InsertRecurse = RotateRight(pNode)
        End If
    Case Is < -1    ' Note that <0 means we MUST have a right child.
        Dim pRigt As Long: pRigt = Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild
        If Key > Groups(Group(pRigt)).Nodes(GNode(pRigt)).Key Then      ' Right  Right  Case.
            InsertRecurse = RotateLeft(pNode)
        Else                                                            ' Right  Left  Case.
            Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild = RotateRight(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)
            InsertRecurse = RotateLeft(pNode)
        End If
    Case Else
        InsertRecurse = pNode                                           ' It was balanced.
    End Select
End Function

Private Function DeleteRecurse(ByVal pNode As Long, Key As Long) As Long
    ' Upon return of first call, the return is what should be the new tree's root.
    ' On each recursion, it returns the new root for the branch being traversed.
    ' Basically, the deletion is accomplished via many node replacements (to maintain balance).
    '
    If pNode = 0& Then Exit Function
    '
    ' Find our node to delete.  BUT, it might be the inorder successor we're working on.
    Select Case Groups(Group(pNode)).Nodes(GNode(pNode)).Key
    Case Is > Key
        Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild = DeleteRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, Key) ' Still looking for node (or inorder successor) to delete.
    Case Is < Key
        Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild = DeleteRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, Key) ' Still looking for node (or inorder successor) to delete.
    Case Else
        Dim pTemp As Long
        ' Found a match, but we keep recursing until we find at least one empty child.
        If Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild <> 0& And Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild <> 0& Then
            '
            ' Both children are present, so we replace pNode.RigtChild (choice) with inorder successor, and then delete it.
            ' We will ONLY get into this code ONCE or not at all.
            '
            pTemp = MinKeyNode(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)                             ' Get the inorder successor (smallest in the right subtree).
            '
            ' ************************************************************
            ' ************************************************************
            '
            '   Below is the only place where we actually move/copy data between nodes.
            '   Use whatever "MOVE" method is fastest for the Key and Value.
            '
            ' ************************************************************
            ' ************************************************************
            '
            Groups(Group(pNode)).Nodes(GNode(pNode)).Key = Groups(Group(pTemp)).Nodes(GNode(pTemp)).Key             ' Copy the inorder successor's data to this node.
            LetSet Groups(Group(pNode)).Nodes(GNode(pNode)).Value, Groups(Group(pTemp)).Nodes(GNode(pTemp)).Value   ' And the value.
            '
            ' Delete the inorder successor (via replacement).
            Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild = DeleteRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, _
                                                                               Groups(Group(pTemp)).Nodes(GNode(pTemp)).Key)
        Else
            ' We definitely (and ONLY) get into this code once (for node OR inorder successor to delete).
            ' Only one or no child.  We finally actually get to do our deletion, and then unwind.
            ' We must figure out which child there is, if there is one.
            ' It's either the node to delete, or the inorder successor, which was found above.
            '
            ' If we have a RigtChild, take it.  Else, take LeftChild whether we have it or not.
            If Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild = 0& Then
                pTemp = Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild
            Else
                pTemp = Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild
            End If
            '
            ' HERE, in both "no child" or "one child" case, the pNode reference WILL be released.
            ' Using an INDEXED approach, we could deallocate the index here, setting pNode=0.
            ' In the "one child" case, it will take on another index value.
            '
            DeAllocIndex pNode                                              ' This decrements miNodeCount.
            If pTemp = 0& Then                                              ' No child case.
                pNode = 0&                                                  ' Do deletion (no child case).  pNode will be returned to caller.
            Else                                                            ' One child case.
                pNode = pTemp                                               ' Use the non-empty child.  pNode will be returned to caller (possibly rotated).
            End If
        End If
    End Select
    '
    ' No more recursion (just unwinding) once we get down here.
    '
    If pNode = 0& Then Exit Function                                        ' A leaf, so no balancing necessary.  Just return nothing and keep unwinding.
    '                                                                       ' This line isn't absolutely necessary, but it may speed things up a bit.
    '                                                                       ' And, this can't be true unless we're currently still unwinding.
    ' Update height of current node.  Updating all as recursion unwinds.
    Groups(Group(pNode)).Nodes(GNode(pNode)).Height = Max(Height(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild), _
                                                          Height(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)) + 1
    '
    ' Balance unbalanced [Abs(Balance) > 1] node(s).  Rebalancing all as recursion unwinds.
    Select Case BalanceFactor(pNode)
    Case Is > 1     ' Note that >0 means we MUST have a left child.
        If BalanceFactor(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild) >= 0 Then  ' Left  Left  Case.
            DeleteRecurse = RotateRight(pNode)
        Else ' BalanceFactor(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild) < 0    ' Left  Right  Case.
            Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild = RotateLeft(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild)
            DeleteRecurse = RotateRight(pNode)
        End If
    Case Is < -1    ' Note that <0 means we MUST have a right child.
        If BalanceFactor(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild) <= 0 Then  ' Right  Right  Case.
            DeleteRecurse = RotateLeft(pNode)
        Else ' BalanceFactor(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild) > 0    ' Right  Left  Case.
            Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild = RotateRight(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)
            DeleteRecurse = RotateLeft(pNode)
        End If
    Case Else
        DeleteRecurse = pNode                                                           ' Node was balanced.
    End Select
End Function

Private Function ExistsRecurse(ByVal pNode As Long, Key As Long) As Boolean
    ' Recursive function to search for existence of a given key in the tree.
    ' pNode is typically set to pRoot to start.
    '
    If pNode = 0& Then Exit Function                                        ' Hit bottom and didn't find it.
    Select Case Groups(Group(pNode)).Nodes(GNode(pNode)).Key
    Case Is > Key:      ExistsRecurse = ExistsRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, Key)
    Case Is < Key:      ExistsRecurse = ExistsRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, Key)
    Case Else:          ExistsRecurse = True                                ' Found it, so unwind, returning True.
    End Select
End Function

Private Function FindNodeRecurse(ByVal pNode As Long, Key As Long) As Long
    ' Recursive function to find the node of a given key in the tree.
    ' pNode is typically set to pRoot to start.
    '
    If pNode = 0& Then Exit Function                                        ' Hit bottom and didn't find it.
    Select Case Groups(Group(pNode)).Nodes(GNode(pNode)).Key
    Case Is > Key:      FindNodeRecurse = FindNodeRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, Key)
    Case Is < Key:      FindNodeRecurse = FindNodeRecurse(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, Key)
    Case Else:          FindNodeRecurse = pNode                             ' Found it, so unwind and return it.
    End Select
End Function

Private Sub BuildKeyArrayAscendRecurse(pNode As Long, TheKeyArray() As Long, iAccum As Long)
    If pNode = 0& Then Exit Sub                                             ' See if we've got a node.
    BuildKeyArrayAscendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, TheKeyArray(), iAccum       ' Go through left branches.
    TheKeyArray(iAccum) = Groups(Group(pNode)).Nodes(GNode(pNode)).Key
    iAccum = iAccum + 1&
    BuildKeyArrayAscendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, TheKeyArray(), iAccum       ' Go through right branches.
End Sub

Private Sub BuildKeyArrayDescendRecurse(pNode As Long, TheKeyArray() As Long, iAccum As Long)
    If pNode = 0& Then Exit Sub                                             ' See if we've got a node.
    BuildKeyArrayDescendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, TheKeyArray(), iAccum      ' Go through right branches.
    TheKeyArray(iAccum) = Groups(Group(pNode)).Nodes(GNode(pNode)).Key
    iAccum = iAccum + 1&
    BuildKeyArrayDescendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, TheKeyArray(), iAccum      ' Go through left branches.
End Sub

Private Sub BuildValueArrayAscendRecurse(pNode As Long, TheValueArray() As Variant, iAccum As Long)
    If pNode = 0& Then Exit Sub                                                                                 ' See if we've got a node.
    BuildValueArrayAscendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, TheValueArray(), iAccum    ' Go through left branches.
    LetSet TheValueArray(iAccum), Groups(Group(pNode)).Nodes(GNode(pNode)).Value
    iAccum = iAccum + 1&
    BuildValueArrayAscendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, TheValueArray(), iAccum    ' Go through right branches.
End Sub

Private Sub BuildValueArrayDescendRecurse(pNode As Long, TheValueArray() As Variant, iAccum As Long)
    If pNode = 0& Then Exit Sub                                                                                 ' See if we've got a node.
    BuildValueArrayDescendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild, TheValueArray(), iAccum   ' Go through right branches.
    LetSet TheValueArray(iAccum), Groups(Group(pNode)).Nodes(GNode(pNode)).Value
    iAccum = iAccum + 1&
    BuildValueArrayDescendRecurse Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild, TheValueArray(), iAccum   ' Go through left branches.
End Sub

' *******************************************************************
' *******************************************************************
'
'   Helper functions --- WITHOUT recursion.
'
' *******************************************************************
' *******************************************************************

Private Sub CancelIterating()
    Erase IterateBranch
    Erase IterateLeftRigt
    BranchNodeCount = 0&
    Iterating = False
End Sub

Private Function RotateRight(pNode As Long) As Long
    ' Perform single right rotation.
    '
    Dim pLeft As Long
    pLeft = Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild
    Dim pRigt As Long
    pRigt = Groups(Group(pLeft)).Nodes(GNode(pLeft)).RigtChild
    '
    ' Perform rotation.
    Groups(Group(pLeft)).Nodes(GNode(pLeft)).RigtChild = pNode
    Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild = pRigt
    '
    ' Update heights
    Groups(Group(pNode)).Nodes(GNode(pNode)).Height = Max(Height(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild), _
                                                          Height(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)) + 1
    Groups(Group(pLeft)).Nodes(GNode(pLeft)).Height = Max(Height(Groups(Group(pLeft)).Nodes(GNode(pLeft)).LeftChild), _
                                                          Height(Groups(Group(pLeft)).Nodes(GNode(pLeft)).RigtChild)) + 1
    '
    ' Return new root.
    RotateRight = pLeft
End Function

Private Function RotateLeft(pNode As Long) As Long
    ' Perform single left rotation.
    '
    Dim pRigt As Long
    pRigt = Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild
    Dim pLeft As Long
    pLeft = Groups(Group(pRigt)).Nodes(GNode(pRigt)).LeftChild
    '
    ' Perform rotation.
    Groups(Group(pRigt)).Nodes(GNode(pRigt)).LeftChild = pNode
    Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild = pLeft
    '
    ' Update heights.
    Groups(Group(pNode)).Nodes(GNode(pNode)).Height = Max(Height(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild), _
                                                          Height(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)) + 1
    Groups(Group(pRigt)).Nodes(GNode(pRigt)).Height = Max(Height(Groups(Group(pRigt)).Nodes(GNode(pRigt)).LeftChild), _
                                                          Height(Groups(Group(pRigt)).Nodes(GNode(pRigt)).RigtChild)) + 1
    '
    ' Return new root.
    RotateLeft = pRigt
End Function

Private Function BalanceFactor(pNode As Long) As Integer
    ' <0 indicates right heavy.  >0 indicates left heavy.
    '
    If pNode = 0& Then
        BalanceFactor = 0
    Else
        BalanceFactor = Height(Groups(Group(pNode)).Nodes(GNode(pNode)).LeftChild) - _
                        Height(Groups(Group(pNode)).Nodes(GNode(pNode)).RigtChild)
    End If
End Function

Private Function Height(pNode As Long) As Integer
    ' A height of -1 indicates that there's no node at all.
    ' Just the node and no children, Height=0.
    '
    If pNode = 0& Then Height = -1 Else Height = Groups(Group(pNode)).Nodes(GNode(pNode)).Height
End Function

Private Function Max(i1 As Integer, i2 As Integer) As Integer
    If i1 > i2 Then Max = i1 Else Max = i2
End Function

Private Function MinKeyNode(pNode As Long) As Long
    ' Find node with minimum value in subtree with pNode as root of that subtree.
    ' Set pNode to pRoot to find overall minimum key.
    '
    Dim pThis As Long
    pThis = pNode
    '
    ' Loop down to find the leftmost leaf.
    Do While Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
        pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).LeftChild
    Loop
    '
    MinKeyNode = pThis
End Function

Private Function MaxKeyNode(pNode As Long) As Long
    ' Find node with maximum value in subtree with pNode as root of that subtree.
    ' Set pNode to pRoot to find overall maximum key.
    '
    Dim pThis As Long
    pThis = pNode
    '
    ' Loop down to find the rightmost leaf.
    Do While Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
        pThis = Groups(Group(pThis)).Nodes(GNode(pThis)).RigtChild
    Loop
    '
    MaxKeyNode = pThis
End Function

Private Sub BuildKeyArray(TheKeyArray() As Long, AscendDescend As BstAscendDescend_LngKey_VarVal)
    If miNodeCount = 0& Then
        MakeZeroToNegOneArray ArrPtr(TheKeyArray)
        Exit Sub
    End If
    '
    Dim iAccum As Long
    ReDim TheKeyArray(miNodeCount - 1&)
    If AscendDescend = bstDescend_LngKey_VarVal Then
        BuildKeyArrayDescendRecurse pRoot, TheKeyArray, iAccum
    Else
        BuildKeyArrayAscendRecurse pRoot, TheKeyArray, iAccum
    End If
End Sub

Private Sub BuildValueArray(TheValueArray() As Variant, AscendDescend As BstAscendDescend_LngKey_VarVal)
    If miNodeCount = 0& Then
        MakeZeroToNegOneArray ArrPtr(TheValueArray)
        Exit Sub
    End If
    '
    Dim iAccum As Long
    ReDim TheValueArray(miNodeCount - 1&)
    If AscendDescend = bstDescend_LngKey_VarVal Then
        BuildValueArrayDescendRecurse pRoot, TheValueArray, iAccum
    Else
        BuildValueArrayAscendRecurse pRoot, TheValueArray, iAccum
    End If
End Sub

Private Sub MakeZeroToNegOneArray(pArray As Long, Optional cDims As Long = 1&)
    SafeArrayAllocDescriptor cDims, ByVal pArray
End Sub

Private Sub LetSet(vDest As Variant, vSource As Variant)
    If IsObject(vSource) Then
        Set vDest = vSource
    Else
        vDest = vSource
    End If
End Sub

' *******************************************************************
' *******************************************************************
'
'   Index management procedures --- no recursion.
'
' *******************************************************************
' *******************************************************************

Private Function Group(Index As Long) As Long
    ' This points at the group the index is in.
    Group = (Index - 1&) \ NODES_PER_GROUP
End Function

Private Function GNode(Index As Long) As Long
    ' This points at the node within a group.
    GNode = (Index - 1&) Mod NODES_PER_GROUP
End Function

Private Sub DeAllocIndex(Index As Long)
    ' Index MUST be a valid index, although it is possible to be > miNodeCount.
    ' Here, we just add it to the DeletedIndexes() array.
    ' This also decrements our miNodeCount.
    ' The Index is also set to ZERO upon return, so it must be ByRef.
    '
    If UBound(DeletedIndexes) + 1& = miDeletedIndexCount Then
        ReDim Preserve DeletedIndexes(UBound(DeletedIndexes) + DELETED_BUMP)
    End If
    DeletedIndexes(miDeletedIndexCount) = Index
    miDeletedIndexCount = miDeletedIndexCount + 1&
    miNodeCount = miNodeCount - 1&
    Index = 0&                                      ' This sets the passed in Index pointer to ZERO.
End Sub

Private Function AllocIndex() As Long
    ' This also increments our miNodeCount.
    ' Don't forget that Index is ONE-based.
    '
    miNodeCount = miNodeCount + 1&                                          ' We increment first, as it makes the following a bit easier.
    '
    If miDeletedIndexCount Then                                             ' See if we can re-use a deleted slot.
        miDeletedIndexCount = miDeletedIndexCount - 1&
        AllocIndex = DeletedIndexes(miDeletedIndexCount)
    Else                                                                    ' No deleted slot, so assign node from end.
        If (UBound(Groups) + 1&) * NODES_PER_GROUP - miNodeCount < 0& Then  ' See if we need more memory.
            ReDim Preserve Groups(UBound(Groups) + 1&)
            ReDim Groups(UBound(Groups)).Nodes(NODES_PER_GROUP - 1&)
        End If
        AllocIndex = miNodeCount                                            ' Our new index is the end.
    End If
End Function


