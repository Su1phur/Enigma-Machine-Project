Imports System.IO
Imports System.Convert
Class CEnigma
	Const NumberOfRotors = 2 ' 0 is also used, so 2 means there are 3 rotors - 0 is the rightmost
	Structure ButtonPair
		Dim Button As Button
		Dim Paired As Boolean
	End Structure
	' Representation of the Enigma parts
	Public Rotors(NumberOfRotors) As CRotor
	Public Reflector As CRotor  ' Umkehrwalze/UKW Reflector plate
	Public Plugboard As CPlugboard
	' Computational parts needed
	Public Colors(25) As Color
	Public EnableKeypress As Boolean = False
	Public TxBInput As TextBox
	Public LSTMethod As ListBox
	Public LSTOutput As ListBox
	Public ButtonList(25) As ButtonPair
	Sub New(ByRef InputBox As TextBox, ByRef MethodBox As ListBox, ByRef OutputBox As ListBox)
		TxBInput = InputBox
		LSTMethod = MethodBox
		LSTOutput = OutputBox

		EnigmaSetup()    ' Default Setup
		EnableKeypress = True
	End Sub

	' Enigma Set Up
	Sub EnigmaSetup()
		Dim counter As Integer = 0
		SetupColors()

		Do
			Dim TempRotor As New CRotor
			TempRotor.LoadRotors(counter)
			Rotors(counter) = TempRotor
			counter += 1
		Loop Until counter = 3 ' 0 is also used

		Dim TempReflector As New CRotor
		TempReflector.LoadRotors(9)
		Reflector = TempReflector

		Dim TempPlugboard As New CPlugboard()
		Plugboard = TempPlugboard
	End Sub
	Sub SelectRotors(ByVal Combobox As ComboBox, ByRef Rotor As CRotor, ByRef choice As String)
		Dim TempRotor As New CRotor

		TempRotor.LoadRotors(choice)
		Rotor = TempRotor
	End Sub
	Sub ImportEnigma(ByRef aform As Form, ByRef filename As String)
		Dim items() As String
		Dim textline As String
		Dim Noofitems As Integer
		Dim x As Integer = 0

		FileOpen(2, filename, OpenMode.Input)

		textline = LineInput(2)
		items = Split(textline, ",") ' Use to splice lines of text to load and unload 
		Noofitems = items.Length

		For x = 0 To Noofitems - 1
			If x = 0 Then
				Rotors(0).LoadAndSetRotor(items(0), items(x + 1))

			ElseIf x = 2 Then
				Rotors(1).LoadAndSetRotor(items(2), items(x + 1))

			ElseIf x = 4 Then
				Rotors(2).LoadAndSetRotor(items(4), items(x + 1))

			ElseIf x > 5 And x < Noofitems Then
				For Each ctl In aform.Controls
					If TypeOf ctl Is Button Then
						If ctl.text = items(x) Then
							SetTempButton(ctl)
						End If
					End If
				Next
			End If
		Next

		FileClose(2)
	End Sub
	Sub ExportEnigmaSettings()
		Dim Settings As String = ""
		Dim Counter As Integer = 0

		FileOpen(1, "Import.txt", OpenMode.Output)

		Do Until Counter = 3
			Settings = Settings & Rotors(Counter).RotorID & "," & Rotors(Counter).RotorPositionQueue(0).Key & ","
			Counter += 1
		Loop

		For Counter2 = 0 To 25
			If Not ButtonList(Counter2).Button Is Nothing Then
				Settings = Settings & ButtonList(Counter2).Button.Text & ","
			End If
		Next

		PrintLine(1, Settings)
		FileClose(1)

		MsgBox("Settings exported.")
	End Sub

	' Enigma Encryption
	Function AppendProcess(ByRef Process As String, ByVal character As Char, ByVal SignificantStep As Boolean) ' For debugging
		If SignificantStep = True Then
			Process = Process & " - " & character
			Return Process
		Else
			Process = Process & character
			Return Process
		End If
	End Function
	Sub EncryptionProcess(UserInput)
		Dim Output As Char
		Dim TempStorage As Char
		Dim StepNumber As Integer
		Dim Process As String = ""

		TempStorage = UserInput
		AppendProcess(Process, UserInput, False)

		TempStorage = PlugboardEncryption(TempStorage)
		AppendProcess(Process, TempStorage, True)

		Rotors(0).RotateRotor()
		CheckRotorStepping(Rotors(0), Rotors(1), Rotors(2)) ' Checks if the rotors need to step

		For StepNumber = 0 To 2 ' Do this process 3 times
			TempStorage = FirstWayEncryption(TempStorage, Rotors(StepNumber)) ' Electric signal from press goes through the 3 rotors
			If StepNumber = 0 Then
				AppendProcess(Process, TempStorage, True)
			Else
				AppendProcess(Process, TempStorage, False)
			End If
		Next

		TempStorage = ReflectorEncryption(TempStorage, Reflector) ' Electric signal passes through the Reflector
		AppendProcess(Process, TempStorage, True)

		For StepNumber = 2 To 0 Step -1                         ' Electric signal passes from Reflector back through the rotors
			TempStorage = SecondWayEncryption(TempStorage, Rotors(StepNumber))
			If StepNumber = 2 Then
				AppendProcess(Process, TempStorage, True)
			Else
				AppendProcess(Process, TempStorage, False)
			End If
		Next

		TempStorage = PlugboardEncryption(TempStorage)
		Output = TempStorage
		AppendProcess(Process, TempStorage, True)

		OutputEncryptedCharacter(Process, Output)
	End Sub
	Function PlugboardEncryption(ByVal Input As Char)
		Dim TempKey As Char
		Dim Output As Char
		TempKey = Input
		If Plugboard.PlugboardKVPDict.ContainsKey(TempKey) = True Then
			Output = Plugboard.PlugboardKVPDict(TempKey)
			Return Output
		ElseIf Plugboard.ReversedPlugboardKVPDict.ContainsKey(TempKey) = True Then
			Output = Plugboard.ReversedPlugboardKVPDict(TempKey)
			Return Output
		Else
			Return Input
		End If
	End Function
	Function FirstWayEncryption(ByRef Input As Char, ByRef Rotor As CRotor)
		Dim Tempchar As Char
		Dim TempIndex As Integer
		Dim tempindex2 As Integer

		TempIndex = FindValue(Input)
		Tempchar = Rotor.RotorPositionQueue(TempIndex).Value

		tempindex2 = Asc(Tempchar) - Rotor.RotorOffset
		If tempindex2 >= Asc("A") And tempindex2 <= Asc("Z") Then
			Tempchar = Chr(tempindex2)
		Else
			tempindex2 += 26
			Tempchar = Chr(tempindex2)
		End If

		Return Tempchar
	End Function
	Sub CheckRotorStepping(ByRef Rotor1 As CRotor, ByRef Rotor2 As CRotor, ByRef Rotor3 As CRotor)
		If Rotor1.RotorPositionQueue.Peek.Key = Rotor1.TurnoverPosition Then
			Rotor2.RotateRotor() ' outsourced to a different subroutine 
		End If

		If Rotor2.RotorPositionQueue.Peek.Key = Rotor2.TurnoverPosition Then
			Rotor3.RotateRotor() ' outsourced to a different subroutine 
		End If
	End Sub
	Function FindValue(TempChar)
		Dim index As Integer
		Dim Asciinumber As Integer
		Asciinumber = Asc(TempChar)

		index = Asciinumber - Asc("A")

		Return index
	End Function
	Function ReflectorEncryption(ByRef Input As Char, ByRef Reflector As CRotor) ' Reflector Rotor Reflects the signal
		'Dim tempindex As Integer
		'Dim Output As Char

		For Each KVP As KeyValuePair(Of Char, Char) In Reflector.RotorPositionQueue
			If Input = KVP.Key Then
				Return KVP.Value
			End If
		Next
	End Function
	Function SecondWayEncryption(ByRef Input As Char, ByRef Rotor As CRotor) ' Electric signal passes bothways
		Dim Tempindex As Integer
		Dim TempChar As Char
		Dim tempindex2 As Integer

		Tempindex = FindValue(Input)
		TempChar = Rotor.ReversedRotorPositionQueue(Tempindex).Value

		tempindex2 = Asc(TempChar) - Rotor.RotorOffset
		If tempindex2 >= Asc("A") And tempindex2 <= Asc("Z") Then
			TempChar = Chr(tempindex2)
		Else
			tempindex2 += 26
			TempChar = Chr(tempindex2)
		End If

		Return TempChar
	End Function
	Sub OutputEncryptedCharacter(Method, Output)
		LSTMethod.Items.Add(Method)
		LSTOutput.Items.Add(Output)
	End Sub

	' Button Interactions
	Function CheckButtonPressed(TempButton As Button)
		If ButtonList IsNot Nothing Then                ' Check if button list exists
			For Each ButtonPair In ButtonList           ' Declaration to do this for every buttonpair
				If ButtonPair.Button IsNot Nothing Then ' If there is a button in the button pair, check
					If ButtonPair.Button Is TempButton Then
						Return True
					End If
				End If
			Next
		End If
		Return False
	End Function
	Function CheckQueueEmptySpace()
		If Not ButtonList(0).Button Is Nothing Then
			For Counter = 0 To 25
				If Not (ButtonList(Counter).Button Is Nothing) Then
					' Nothing happens lol
				Else
					Return Counter
				End If
			Next
		End If
		Return 0
	End Function
	Sub SetupColors()
		Colors(0) = Color.Red
		Colors(1) = Color.Red
		Colors(2) = Color.Blue
		Colors(3) = Color.Blue
		Colors(4) = Color.Orange
		Colors(5) = Color.Orange
		Colors(6) = Color.Green
		Colors(7) = Color.Green
		Colors(8) = Color.Indigo
		Colors(9) = Color.Indigo
		Colors(10) = Color.Violet
		Colors(11) = Color.Violet
		Colors(12) = Color.Yellow
		Colors(13) = Color.Yellow
		Colors(14) = Color.Lime
		Colors(15) = Color.Lime
		Colors(16) = Color.Plum
		Colors(17) = Color.Plum
		Colors(18) = Color.Purple
		Colors(19) = Color.Purple
		Colors(20) = Color.Cyan
		Colors(21) = Color.Cyan
		Colors(22) = Color.DarkBlue
		Colors(23) = Color.DarkBlue
		Colors(24) = Color.Teal
		Colors(25) = Color.Teal
	End Sub
	Sub SetTempButton(TempButton As Button)
		Dim Counter As Integer
		Dim TempButtonPair As New ButtonPair

		Counter = CheckQueueEmptySpace()

		TempButtonPair.Button = TempButton
		TempButtonPair.Paired = False
		TempButtonPair.Button.BackColor = Colors(Counter)
		ButtonList(Counter) = TempButtonPair

		If Counter Mod 2 <> 0 And Counter <> 0 Then ' If Counter is not even or 0 / For every other Counter
			PairButton(Counter)
		End If

	End Sub
	Sub PairButton(ByVal Counter As Integer)
		Dim Newkey As Char
		Dim NewValue As Char

		ButtonList(Counter - 1).Paired = True
		ButtonList(Counter).Paired = True
		Newkey = ButtonList(Counter - 1).Button.Text
		NewValue = ButtonList(Counter).Button.Text
		Plugboard.NewPlugboardPair(Newkey, NewValue)
	End Sub
	Sub RemoveTempButton(Tempbutton As Button)
		For Counter = 0 To 25
			If ButtonList(Counter).Button Is Tempbutton Then
				Plugboard.RemovePlugboardPair(ButtonList(Counter).Button.Text)
				ResetButtonList(Counter)
				' Put in a check here just in case
				If Counter = 0 Or Counter Mod 2 = 0 Then
					ResetButtonList(Counter + 1)    '	Counter is Even, the next one is the pair
				Else
					ResetButtonList(Counter - 1)    '	Counter is Odd, the previous one is the pair
				End If
			End If
		Next
	End Sub
	Sub ResetButtonList(Input As Integer)
		ButtonList(Input).Button.BackColor = Button.DefaultBackColor
		ButtonList(Input).Button = Nothing
		ButtonList(Input).Paired = False
	End Sub

	' Program Functions
	Sub BackSpace()
		Dim Counter = LSTMethod.Items.Count - 1
		If Counter >= 0 Then
			LSTMethod.Items.RemoveAt(Counter)
			LSTOutput.Items.RemoveAt(Counter)
			Rotors(0).ReverseRotateRotor()  ' unfortunately it would make it far too complex for it to remember what it did last step
		End If

	End Sub
	Sub ExportTextToFile()
		Dim textline As String = ""

		FileOpen(1, "Export.txt", OpenMode.Output)

		For x = 0 To LSTOutput.Items.Count - 1
			If x Mod 5 = 0 And x <> 0 Then  ' Every 5 characters, make a space
				textline = textline & " " & LSTOutput.Items(x).ToString
			Else
				textline = textline & LSTOutput.Items(x).ToString
			End If
		Next

		PrintLine(1, textline)
		FileClose(1)
	End Sub
	Function ExportTextToEmail()
		Dim Textline As String = ""

		For x = 0 To LSTOutput.Items.Count - 1
			If x Mod 5 = 0 And x <> 0 Then  ' Every 5 characters, make a space
				Textline = Textline & " " & LSTOutput.Items(x).ToString
			Else
				Textline = Textline & LSTOutput.Items(x).ToString
			End If
		Next

		Return Textline
	End Function
End Class