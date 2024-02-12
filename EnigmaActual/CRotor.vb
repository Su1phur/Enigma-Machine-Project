Class CRotor
	Public RotorKVPDict As New Dictionary(Of Char, Char)
	Public ReversedRotorKVPDict As New Dictionary(Of Char, Char)
	Public RotorPositionQueue As New Queue(Of KeyValuePair(Of Char, Char))
	Public ReversedRotorPositionQueue As New Queue(Of KeyValuePair(Of Char, Char))
	Public RotorID As String
	Public RotorOffset As Integer
	Public TurnoverPosition As Char
	Sub LoadRotors(ByRef RotorNumber As String)
		Dim items() As String
		Dim textline As String
		Dim Noofitems As Integer
		Dim x As Integer = 0

		If RotorNumber = 0 Or RotorNumber = "I" Then
			FileOpen(1, "Rotor1KVP.txt", OpenMode.Input)
			TurnoverPosition = "Q"
		ElseIf RotorNumber = 1 Or RotorNumber = "II" Then
			FileOpen(1, "Rotor2KVP.txt", OpenMode.Input)
			TurnoverPosition = "E"
		ElseIf RotorNumber = 2 Or RotorNumber = "III" Then
			FileOpen(1, "Rotor3KVP.txt", OpenMode.Input)
			TurnoverPosition = "V"
		ElseIf RotorNumber = 3 Or RotorNumber = "IV" Then
			FileOpen(1, "Rotor4KVP.txt", OpenMode.Input)
			TurnoverPosition = "J"
		ElseIf RotorNumber = 4 Or RotorNumber = "V" Then
			FileOpen(1, "Rotor5KVP.txt", OpenMode.Input)
			TurnoverPosition = "Z"
		ElseIf RotorNumber = 9 Then
			FileOpen(1, "UKW_KVP.txt", OpenMode.Input)
		End If

		RotorID = RotorNumber

		textline = LineInput(1)
		items = Split(textline, ",") ' Use to splice lines of text to load and unload 
		Noofitems = items.Length

		Do Until x >= Noofitems
			RotorKVPDict.Add(items(x), items(x + 1)) ' ie, A,E means A is connected to E
			x += 2
		Loop

		FileClose(1)

		ReverseRotorKVP() ' to reverse such that the KVP goes both ways

		SetupRotorQueue()

		RotorOffset = 0
	End Sub
	Sub RotateRotor()
		Dim TempQueueItem As KeyValuePair(Of Char, Char)
		Dim TempReverseQueueItem As KeyValuePair(Of Char, Char)

		TempQueueItem = RotorPositionQueue(0)
		RotorPositionQueue.Dequeue()
		RotorPositionQueue.Enqueue(TempQueueItem)

		TempReverseQueueItem = ReversedRotorPositionQueue(0)
		ReversedRotorPositionQueue.Dequeue()
		ReversedRotorPositionQueue.Enqueue(TempReverseQueueItem)

		AdjustRotorOffset(False)
	End Sub
	Sub ReverseRotateRotor()
		Dim TempQueue1 As New Queue(Of KeyValuePair(Of Char, Char))
		Dim TempQueue2 As New Queue(Of KeyValuePair(Of Char, Char))

		For Counter = 25 To 0 Step -1       ' Reverses Queue
			TempQueue1.Enqueue(RotorPositionQueue(Counter))
		Next

		TempQueue1.Enqueue(TempQueue1(0))   ' Duplicates first item and puts it at the end
		TempQueue1.Dequeue()                ' Removes first item

		RotorPositionQueue.Clear()          ' Empties Queue

		For Counter = 25 To 0 Step -1       ' Queues stuff in reverse order
			RotorPositionQueue.Enqueue(TempQueue1(Counter))
		Next

		TempQueue2.Enqueue(ReversedRotorPositionQueue(25))

		For Counter = 0 To 24
			TempQueue2.Enqueue(ReversedRotorPositionQueue(Counter))
		Next

		ReversedRotorPositionQueue = TempQueue2

		AdjustRotorOffset(True)
	End Sub
	Sub ReverseRotorKVP()
		Dim TempDict As New Dictionary(Of Char, Char)
		For Each Pair In RotorKVPDict
			TempDict.Add(Pair.Value, Pair.Key)  ' redundancy/using same code for simplicity
			ReversedRotorKVPDict = TempDict
		Next
	End Sub
	Sub SetupRotorQueue()
		For Each pair In RotorKVPDict ' to make it into a queue
			RotorPositionQueue.Enqueue(pair)
		Next
		For Each pair In ReversedRotorKVPDict  ' to do the reverse
			ReversedRotorPositionQueue.Enqueue(pair)
		Next
		SortQueueAlphabetically(ReversedRotorPositionQueue)
	End Sub
	Sub SortQueueAlphabetically(ByRef InputQueue As Queue(Of KeyValuePair(Of Char, Char)))
		Dim TempQueue As New Queue(Of KeyValuePair(Of Char, Char))
		Dim tempKVP As KeyValuePair(Of Char, Char)
		For z = 0 To 25
			If InputQueue(z).Key = "A" Then
				tempKVP = InputQueue(z)
				TempQueue.Enqueue(tempKVP)
			End If
		Next

		For x = 0 To 25
			For y = 0 To 25
				If Asc(InputQueue(y).Key) - Asc(TempQueue(x).Key) = 1 Then
					TempQueue.Enqueue(InputQueue(y))
				End If
			Next
		Next
		InputQueue = TempQueue
	End Sub
	Sub AdjustRotorOffset(ByVal Clockwise As Boolean)
		If Clockwise = True Then
			RotorOffset -= 1
		Else
			RotorOffset += 1
		End If

		If RotorOffset < 0 Then
			RotorOffset += 26
		End If
		If RotorOffset > 26 Then
			RotorOffset -= 26
		End If
	End Sub
	Sub LoadAndSetRotor(InputID, RotorKey)
		If InputID <> RotorID Then
			RotorKVPDict.Clear()
			ReversedRotorKVPDict.Clear()
			RotorPositionQueue.Clear()
			ReversedRotorPositionQueue.Clear()

			LoadRotors(InputID)
		End If

		Do Until RotorPositionQueue(0).Key = RotorKey
			RotateRotor()
		Loop

	End Sub
End Class
