Public Class Form1
	Dim Enigma As CEnigma
	Dim SetupComplete As Boolean = False
	Dim Lengthbefore As Integer = 0
	Dim Lengthafter As Integer = 0
	Sub Main()

	End Sub
	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Enigma = New CEnigma(TxBInput, LSTMethod, LSTOutput)
		LBLR0Display.Text = Enigma.Rotors(0).RotorPositionQueue(0).Key
		LBLR1Display.Text = Enigma.Rotors(1).RotorPositionQueue(0).Key
		LBLR2Display.Text = Enigma.Rotors(2).RotorPositionQueue(0).Key
		ComboBoxSetup(CmBR0Selection, "I")
		ComboBoxSetup(CMBR1Selection, "II")
		ComboBoxSetup(CmBR2Selection, "III")

		CreateDynamicButtons()

		SetupComplete = True

		UpdateDisplayText()
	End Sub
	Private Sub TxBInput_TextChanged(sender As Object, e As EventArgs) Handles TxBInput.TextChanged
		Dim achar As Char

		Lengthafter = TxBInput.Text.Length
		If Lengthafter - Lengthbefore > 1 Then
			For x = 0 To Lengthafter - 1
				achar = UCase(TxBInput.Text(x))
				If Asc(achar) >= Asc("A") And Asc(achar) <= Asc("Z") Then
					Enigma.EncryptionProcess(UCase(TxBInput.Text(x)))
				End If
			Next
			Lengthbefore = Lengthafter
		ElseIf Lengthafter - Lengthbefore < 0 Then
			Enigma.BackSpace()
		Else
			achar = UCase(TxBInput.Text(Lengthafter - 1))
			If Asc(achar) >= Asc("A") And Asc(achar) <= Asc("Z") Then
				Enigma.EncryptionProcess(achar)
			End If
			Lengthbefore += 1
		End If

		UpdateDisplayText()
	End Sub
	Private Sub UpdateDisplayText()
		LBLR0DisplayPrev.Text = Enigma.Rotors(0).RotorPositionQueue(25).Key
		LBLR0Display.Text = Enigma.Rotors(0).RotorPositionQueue(0).Key
		LBLR0DisplayNext.Text = Enigma.Rotors(0).RotorPositionQueue(1).Key

		LBLR1DisplayPrev.Text = Enigma.Rotors(1).RotorPositionQueue(25).Key
		LBLR1Display.Text = Enigma.Rotors(1).RotorPositionQueue(0).Key
		LBLR1DisplayNext.Text = Enigma.Rotors(1).RotorPositionQueue(1).Key

		LBLR2DisplayPrev.Text = Enigma.Rotors(2).RotorPositionQueue(25).Key
		LBLR2Display.Text = Enigma.Rotors(2).RotorPositionQueue(0).Key
		LBLR2DisplayNext.Text = Enigma.Rotors(2).RotorPositionQueue(1).Key
	End Sub
	Private Sub BTNR0Increase_Click(sender As Object, e As EventArgs) Handles BTNR0Increase.Click
		Enigma.Rotors(0).RotateRotor()
		LBLR0Display.Text = Enigma.Rotors(0).RotorPositionQueue(0).Key
		UpdateDisplayText()
	End Sub

	Private Sub BTNR0Decrease_Click(sender As Object, e As EventArgs) Handles BTNR0Decrease.Click
		Enigma.Rotors(0).ReverseRotateRotor()
		LBLR0Display.Text = Enigma.Rotors(0).RotorPositionQueue(0).Key
		UpdateDisplayText()
	End Sub
	Private Sub BTNR1Increase_Click(sender As Object, e As EventArgs) Handles BTNR1Increase.Click
		Enigma.Rotors(1).RotateRotor()
		LBLR1Display.Text = Enigma.Rotors(1).RotorPositionQueue(0).Key
		UpdateDisplayText()
	End Sub
	Private Sub BTNR1Decrease_Click(sender As Object, e As EventArgs) Handles BTNR1Decrease.Click
		Enigma.Rotors(1).ReverseRotateRotor()
		LBLR1Display.Text = Enigma.Rotors(1).RotorPositionQueue(0).Key
		UpdateDisplayText()
	End Sub
	Private Sub BTNR2Increase_Click(sender As Object, e As EventArgs) Handles BTNR2Increase.Click
		Enigma.Rotors(2).RotateRotor()
		LBLR2Display.Text = Enigma.Rotors(2).RotorPositionQueue(0).Key
		UpdateDisplayText()
	End Sub
	Private Sub BTNR2Decrease_Click(sender As Object, e As EventArgs) Handles BTNR2Decrease.Click
		Enigma.Rotors(2).ReverseRotateRotor()
		LBLR2Display.Text = Enigma.Rotors(2).RotorPositionQueue(0).Key
		UpdateDisplayText()
	End Sub
	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		TxBInput.Clear()
		LSTMethod.Items.Clear()
		LSTOutput.Items.Clear()
	End Sub
	Private Sub CreateDynamicButtons()
		Dim butarr(25) As Button
		Dim letters As String = "QWERTYUIOPASDFGHJKLZXCVBNM"
		Dim temploc As Point
		Dim tempsize As Drawing.Size
		Dim tempmargin As Padding
		tempsize.Width = 62
		tempsize.Height = 20
		temploc.X = 17
		temploc.Y = 425
		tempmargin.All = 3

		For Counter = 0 To 25
			Dim tempbut = New Button
			tempbut.Text = letters(Counter)
			tempbut.Size = tempsize
			tempbut.Location = temploc
			tempbut.Margin = tempmargin

			temploc.X = temploc.X + 62
			If letters(Counter) = "P" Then
				temploc.X = 49
				temploc.Y += 23
			ElseIf letters(Counter) = "L" Then
				temploc.X = 65
				temploc.Y += 23
			End If

			AddHandler tempbut.Click, AddressOf letterbuts
			butarr(Counter) = tempbut
			Controls.Add(tempbut)
		Next
	End Sub
	Sub letterbuts(ByVal sender As Object, ByVal e As System.EventArgs)
		Dim buttonpressed As Boolean = False
		buttonpressed = Enigma.CheckButtonPressed(sender)
		If buttonpressed = False Then
			Enigma.SetTempButton(sender)
		ElseIf buttonpressed = True Then
			Enigma.RemoveTempButton(sender)
			sender.BackColor = DefaultBackColor
		End If
	End Sub
	Private Sub CmBR0Selection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmBR0Selection.SelectedIndexChanged
		If SetupComplete = True Then
			Enigma.SelectRotors(CmBR0Selection, Enigma.Rotors(0), CmBR0Selection.SelectedIndex)
		End If
	End Sub
	Private Sub ComboBoxSetup(Combobox As ComboBox, Selection As String)
		Dim index As Integer
		Combobox.Items.Add("I")
		Combobox.Items.Add("II")
		Combobox.Items.Add("III")
		Combobox.Items.Add("IV")
		Combobox.Items.Add("V")

		index = Combobox.FindString(Selection)
		Combobox.SelectedIndex = index
	End Sub
	Private Sub CMBR1Selection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMBR1Selection.SelectedIndexChanged
		If SetupComplete = True Then
			Enigma.SelectRotors(CMBR1Selection, Enigma.Rotors(1), CMBR1Selection.SelectedIndex)
		End If
	End Sub

	Private Sub CmBR2Selection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmBR2Selection.SelectedIndexChanged
		If SetupComplete = True Then
			Enigma.SelectRotors(CmBR2Selection, Enigma.Rotors(2), CmBR2Selection.SelectedIndex)
		End If
	End Sub

	Private Sub BTNExport_Click(sender As Object, e As EventArgs) Handles BTNExport.Click
		Enigma.ExportTextToFile()
	End Sub

	Private Sub BTNBackspace_Click(sender As Object, e As EventArgs) Handles BTNBackspace.Click
		Enigma.BackSpace()

		LBLR0Display.Text = Enigma.Rotors(0).RotorPositionQueue(0).Key
	End Sub
	Private Sub BTNExportToEmail_Click(sender As Object, e As EventArgs) Handles BTNExportToEmail.Click
		Form2.Show()
		Form2.TextBox4.Text = Enigma.ExportTextToEmail()
	End Sub

	Private Sub BTNImportSettings_Click(sender As Object, e As EventArgs) Handles BTNImportSettings.Click
		Enigma.ImportEnigma(Me, "Import.txt")
		UpdateDisplayText()
		MsgBox("Settings Imported.")
	End Sub
	Private Sub BTNHelp_Click(sender As Object, e As EventArgs) Handles BTNHelp.Click
		Dim NormalSize As New Size(685, 570)
		Dim HelpSize As New Size(910, 570)
		If Size = NormalSize Then
			Size = HelpSize
		ElseIf Size = HelpSize Then
			Size = NormalSize
		Else
			Size = HelpSize
		End If
	End Sub

	Private Sub BTNExportSettings_Click(sender As Object, e As EventArgs) Handles BTNExportSettings.Click
		Enigma.ExportEnigmaSettings()
	End Sub
End Class
