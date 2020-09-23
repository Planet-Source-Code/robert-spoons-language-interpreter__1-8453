Attribute VB_Name = "ModLngArithmeticStacks"
Public Stack1(100) As String
Public Stack2(100) As Integer
Private m_Stack1Top As Integer
Private m_Stack2Top As Integer
Public Type MyOps
    ops(100) As Integer
End Type

Public Function Pop1() As String
If m_Stack1Top > -1 Then
    Pop1 = Stack1(m_Stack1Top)
    m_Stack1Top = m_Stack1Top - 1
End If
End Function
Public Function Pop2() As Integer
If m_Stack2Top > -1 Then
    Pop2 = Stack2(m_Stack2Top)
    m_Stack2Top = m_Stack2Top - 1
End If
End Function
Public Function Push1(lStr As String)
If m_Stack1Top < 101 Then
    m_Stack1Top = m_Stack1Top + 1
    Stack1(m_Stack1Top) = lStr
End If
End Function
Public Function Push2(lStr As Integer)
If m_Stack2Top < 101 Then
    m_Stack2Top = m_Stack2Top + 1
    Stack2(m_Stack2Top) = lStr
End If
End Function
Public Function StackTop1() As Integer
    StackTop1 = m_Stack1Top
End Function
Public Function StackTop2() As Integer
    StackTop2 = m_Stack2Top
End Function
Public Function Stack1TopItem() As String
    Stack1TopItem = Stack1(m_Stack1Top)
End Function
Public Function Stack2TopItem() As Integer
    Stack2TopItem = Stack2(m_Stack1Top)
End Function
Public Sub ClearStacks()
    m_Stack1Top = 0
    m_Stack2Top = 0
End Sub
