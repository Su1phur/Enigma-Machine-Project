Class CPlugboard
	Public PlugboardKVPDict As New Dictionary(Of Char, Char)
	Public ReversedPlugboardKVPDict As New Dictionary(Of Char, Char)
	' A plugboard is just a dictionary, so no need for a queue
	Sub New()

	End Sub
	Sub NewPlugboardPair(TempKeyChar, TempValueChar)
		PlugboardKVPDict.Add(TempKeyChar, TempValueChar)
		ReversedPlugboardKVPDict.Add(TempValueChar, TempKeyChar)
	End Sub
	Sub RemovePlugboardPair(Input)
		If PlugboardKVPDict.ContainsKey(Input) = True Then ' Input is a Key
			Dim TempValueChar As Char

			TempValueChar = PlugboardKVPDict(Input)
			PlugboardKVPDict.Remove(Input)
			ReversedPlugboardKVPDict.Remove(TempValueChar)
		ElseIf ReversedPlugboardKVPDict.ContainsKey(Input) = True Then ' Input is a Value
			Dim TempKeyChar As Char

			TempKeyChar = ReversedPlugboardKVPDict(Input)
			PlugboardKVPDict.Remove(TempKeyChar)
			ReversedPlugboardKVPDict.Remove(Input)
		End If
	End Sub
End Class
