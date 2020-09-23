Attribute VB_Name = "ModLanguageRANDOM"
Public Sub RandNumber(plint As Integer)
Dim i As Integer
Dim j As Integer
Dim lR1 As String
Dim lR2 As String
Dim dReturn As Integer
Dim compNum As Integer

    compNum = StructureLen
   
    If compNum > 3 Then
        If CheckVal(3) Then
            lR1 = Val(AllItems(3))
        Else
            lR1 = FindItNow(3)
        End If
        If CheckVal(4) Then
            lR2 = Val(AllItems(4))
        Else
            lR2 = FindItNow(4)
        End If
    
        dReturn = Int(Rnd(1) * lR2) + lR1
    
    Else
        If compNum > 2 Then
            If AllItems(3) <> "" Then
                If CheckVal(3) Then
                    lR1 = Val(AllItems(3))
                Else
                    lR1 = FindItNow(3)
                End If
                dReturn = Int(Rnd(1) * lR1)
            Else
                dReturn = Int(Rnd(1) * 101)
            End If
        End If
    End If
    v = 5
    StoreItNow 0, dReturn
    
End Sub

