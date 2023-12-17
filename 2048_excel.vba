Public Const B_LEFT  As Integer = 2
Public Const B_DOWN  As Integer = 6
Public Const B_RIGHT  As Integer = 5
Public Const B_UP  As Integer = 3
Public Const MSG_CELL = "B2"
Public Const COLOR_DELTA As Integer = 25
Public GAME_OVER As Boolean
Public has_changed As Boolean


Public Function IsInBorders(x, y)
    If x < B_LEFT Or x > B_RIGHT Or y < B_UP Or y > B_DOWN Then
        IsInBorders = False
    Else
        IsInBorders = True
    End If
End Function


Public Sub Click_Right()
    has_changed = False
    If Not GAME_OVER Then
        For y = B_UP To B_DOWN
            For x = B_RIGHT To B_LEFT Step -1
                If IsFree(y, x) = False Then
                    MoveToBorder 0, y, x '0 = right 1 = down 2 = left 3 = up
                End If
            Next
        Next
        AddValue
        RecolorField
    End If
End Sub

Public Sub Click_Left()
    has_changed = False
    If Not GAME_OVER Then
        For y = B_UP To B_DOWN
            For x = B_LEFT To B_RIGHT
                If IsFree(y, x) = False Then
                    MoveToBorder 2, y, x '0 = right 1 = down 2 = left 3 = up
                End If
            Next
        Next
        AddValue
        RecolorField
    End If
End Sub

Public Sub Click_Up()
    has_changed = False
    If Not GAME_OVER Then
        For x = B_LEFT To B_RIGHT
            For y = B_UP To B_DOWN
                If IsFree(y, x) = False Then
                    MoveToBorder 3, y, x '0 = right 1 = down 2 = left 3 = up
                End If
            Next
        Next
        AddValue
        RecolorField
    End If
End Sub

Public Sub Click_Down()
    has_changed = False
    If Not GAME_OVER Then
        For x = B_LEFT To B_RIGHT
            For y = B_DOWN To B_UP Step -1
                If IsFree(y, x) = False Then
                    MoveToBorder 1, y, x '0 = right 1 = down 2 = left 3 = up
                End If
            Next
        Next
        AddValue
        RecolorField
    End If
End Sub

Public Sub MoveToBorder(side, y, x)
    'Moves number to the limit and sums it if can
    '0 = right 1 = down 2 = left 3 = up
    Select Case side
        Case 0
            deltaX = 1
            deltaY = 0
        Case 1
            deltaX = 0
            deltaY = 1
        Case 2
            deltaX = -1
            deltaY = 0
        Case 3
            deltaX = 0
            deltaY = -1
    End Select
    While (IsFree(y + deltaY, x + deltaX) And IsInBorders(x + deltaX, y + deltaY))
        Cells(y + deltaY, x + deltaX).value = Cells(y, x).value
        Cells(y, x).value = ""
        x = x + deltaX
        y = y + deltaY
        has_changed = True
    Wend
    If IsInBorders(x + deltaX, y + deltaY) And Cells(y + deltaY, x + deltaX).value = Cells(y, x).value Then
        SumCells y, x, deltaY, deltaX
        has_changed = True
    End If
End Sub

Public Sub Start()
    has_changed = True
    GAME_OVER = False
    Clear
    AddValue
    AddValue
    delta = 0
End Sub

Public Sub Clear()
    'Clears all cells and message cell
    Range(MSG_CELL).value = ""
    For x = B_LEFT To B_RIGHT
        For y = B_UP To B_DOWN
            Cells(y, x) = ""
            Cells(y, x).Interior.Color = RGB(255, 255, 255)
            Cells(y, x).Borders(xlEdgeRight).LineStyle = xlContinuous
            Cells(y, x).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Cells(y, x).Borders(xlEdgeTop).LineStyle = xlContinuous
            Cells(y, x).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Next
    Next
End Sub

Public Sub AddValue()
    'Puts 2 in a random cell or ends the game if can`t
    If Not GameOver Then
        If has_changed Then
            Do
                y = 3 + Int(4 * Rnd)
                x = 2 + Int(4 * Rnd)
            Loop Until Not Cells(y, x).value <> ""
            Cells(y, x) = 2
        End If
    Else
        Range(MSG_CELL).value = "GAME OVER >:)"
    End If
End Sub

Public Function GameOver()
    'If there is at least one empty cell game is not over
    For y = B_LEFT To B_RIGHT
        For x = B_UP To B_DOWN
            If Cells(x, y).value = "" Then
                GameOver = False
                Exit Function
            End If
        Next
    Next
    GAME_OVER = True
    GameOver = True
End Function

Public Function IsFree(y, x) As Boolean
    IsFree = Cells(y, x) = ""
End Function

Public Function SameValue(y1, x1, y2, x2)
    SameValue = Cells(y1, x1).value = Cells(y2, x2).value
End Function

Public Sub SumCells(y, x, deltaY, deltaX)
    Cells(y + deltaY, x + deltaX).value = Cells(y + deltaY, x + deltaX).value * 2
    Cells(y, x).value = ""
End Sub

Public Function powerOfTwo(number, power)
    If number = 1 Then
        powerOfTwo = power
    End If
    If number > 1 Then
        powerOfTwo = powerOfTwo(Int(number / 2), power + 1)
    End If
End Function


Public Sub RecolorField()
    Dim power As Integer
    If has_changed Then
        For x = B_LEFT To B_RIGHT
            For y = B_UP To B_DOWN
                power = powerOfTwo(Cells(y, x).value, 0)
                Cells(y, x).Interior.Color = RGB(255, 255 - COLOR_DELTA * power, 255 - COLOR_DELTA * power)
                Cells(y, x).Borders(xlEdgeRight).LineStyle = xlContinuous
                Cells(y, x).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Cells(y, x).Borders(xlEdgeTop).LineStyle = xlContinuous
                Cells(y, x).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Next
        Next
    End If

End Sub
