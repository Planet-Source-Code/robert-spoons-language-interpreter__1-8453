Attribute VB_Name = "ModDateAndTime"
'This module is just a straight passthrough

Public Function ReturnDate() As String
Dim myDate As Date
    myDate = Date
    ReturnDate = CStr(myDate)
End Function

Public Function ReturnTime() As String
Dim myTime
    myTime = Time
    ReturnTime = CStr(myTime)
End Function
