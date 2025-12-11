Attribute VB_Name = "Module1"
Sub SecretSantaOptimized()
    Dim ws As Worksheet, wsResult As Worksheet
    Set ws = ThisWorkbook.Sheets("Participants")
    Set wsResult = ThisWorkbook.Sheets("Results")
    
    wsResult.Cells.Clear
    wsResult.Range("A1:C1").Value = Array("Giver Name", "Email ID", "Receiver Name")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim givers() As String, receivers() As String, courses() As String
    Dim interests() As Variant
    Dim i As Long, j As Long
    
    ReDim givers(1 To lastRow - 1)
    ReDim receivers(1 To lastRow - 1)
    ReDim courses(1 To lastRow - 1)
    ReDim interests(1 To lastRow - 1, 1 To 3)
    
    ' Load participant data
    For i = 2 To lastRow
        givers(i - 1) = ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value
        receivers(i - 1) = givers(i - 1)
        courses(i - 1) = ws.Cells(i, 4).Value
        interests(i - 1, 1) = ws.Cells(i, 5).Value
        interests(i - 1, 2) = ws.Cells(i, 6).Value
        interests(i - 1, 3) = ws.Cells(i, 7).Value
    Next i
    
    Dim bestScore As Long, bestShuffle() As String
    bestScore = -99999
    
    Dim attempt As Long, maxAttempts As Long: maxAttempts = 5000
    Dim valid As Boolean, totalScore As Long
    
    Dim shuffled() As String
    
    For attempt = 1 To maxAttempts
        shuffled = ShuffleArray(receivers)
        valid = True
        totalScore = 0
        
        For i = 1 To UBound(givers)
            ' Hard constraint: no self-assignment
            If givers(i) = shuffled(i) Then
                valid = False
                Exit For
            End If
            
            ' Soft constraint: prefer avoiding same course
            If courses(i) = courses(Application.Match(shuffled(i), givers, 0)) Then
                totalScore = totalScore - 1 ' penalize same-course slightly
            End If
            
            ' Interest complementarity: fewer shared interests = higher score
            totalScore = totalScore - CountSharedInterests(i, shuffled(i), givers, interests)
        Next i
        
        If valid And totalScore > bestScore Then
            bestScore = totalScore
            bestShuffle = shuffled
        End If
    Next attempt
    
    If bestScore = -99999 Then
        MsgBox "Could not find valid pairings."
        Exit Sub
    End If
    
    ' Output final pairings
    For i = 1 To UBound(givers)
        wsResult.Cells(i + 1, 1).Value = givers(i)
        wsResult.Cells(i + 1, 2).Value = ws.Cells(i + 1, 3).Value
        wsResult.Cells(i + 1, 3).Value = bestShuffle(i)
    Next i
    
    MsgBox "Secret Santa pairings generated!"
End Sub

' Function to shuffle array randomly
Function ShuffleArray(arr() As String) As String()
    Dim temp As String, i As Long, j As Long
    Dim n As Long: n = UBound(arr)
    Randomize
    For i = n To 1 Step -1
        j = Int(Rnd() * i) + 1
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
    ShuffleArray = arr
End Function

' Function to count shared interests between giver and receiver
Function CountSharedInterests(giverIndex As Long, receiverName As String, givers() As String, interests() As Variant) As Long
    Dim receiverIndex As Long
    receiverIndex = Application.Match(receiverName, givers, 0)
    
    Dim count As Long
    count = 0
    Dim k As Long
    For k = 1 To 3
        If interests(giverIndex, k) = interests(receiverIndex, k) Then count = count + 1
    Next k
    CountSharedInterests = count
End Function

