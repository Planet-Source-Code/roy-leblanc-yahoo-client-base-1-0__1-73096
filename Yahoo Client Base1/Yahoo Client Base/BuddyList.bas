Attribute VB_Name = "BuddyList"
'Yahoo Client Base Example 1.0
'Written By : EliteRoy(Roy LeBlanc)
'My Website : http://eliteroy.com
'Credits Go To -
'C-4 BuddyList Parse Method
'Expulsion, Adam, and Dubee For Login Method

Public Sub ParseBuddyList(Data As String, TV As TreeView)

    Dim strGroup()           As String
    Dim strBuddy()           As String
    Dim lngGroup             As Long
    Dim lngBuddy             As Long
    Dim strCurrentGroup      As String
    Dim Nodx                 As Node
    
On Error GoTo ErrHandler
      
      strGroup = Split(Data, "65��")
      TV.Nodes.Clear
        
    For lngGroup = 1 To UBound(strGroup)
        strCurrentGroup = strGroup(lngGroup)
        strGroup(lngGroup) = Split(strGroup(lngGroup), "��302��")(0)
        Set Nodx = TV.Nodes.Add(, , strGroup(lngGroup), strGroup(lngGroup), 2)
        Nodx.Expanded = True
        strBuddy = Split(strCurrentGroup, "��7��")
        
        For lngBuddy = 1 To UBound(strBuddy)
            strBuddy(lngBuddy) = Split(strBuddy(lngBuddy), "��301��")(0)
            Set Nodx = TV.Nodes.Add(strGroup(lngGroup), tvwChild, , strBuddy(lngBuddy), 1)
        Next lngBuddy
        
    Next lngGroup
    
ErrHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbOKOnly, "Error: " & Err.Number: Exit Sub
End Sub
