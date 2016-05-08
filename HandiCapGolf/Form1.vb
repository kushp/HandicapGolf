'############################'
'# Kush Patel               #'
'# kushpatel35@hotmail.com  #'
'# V1.2.0                   #'
'############################'

Public Class Form1

    'Generate a file name to use for saving, RTF file
    Dim fileName As String = "HandiCapGolf-" + Date.Now.ToLongDateString.Replace(" ", "") + "-" + Date.Now.ToLongTimeString.Replace(":", "_").Replace(" ", "") + ".rtf"

    'This is used to save generated data, this is useful for sorting later on, id really isn't used for anything. 
    Dim savedData As ArrayList = New ArrayList()
    Dim id As Integer

    'Save Data File
    Dim dataSaveFileName As String

    'This is important, this is all the number of worse holes and adjust by values for each gross score. 
    'Format: {Gross Score, Number of Worse Holes To Remove, Adjust By} 
    'Lowest is 66 and highest is 165
    Dim data(,) As Double = { _
        {66, 0.5, -2}, {67, 0.5, -1}, {68, 0.5, 0}, {69, 0.5, 1}, {70, 0.5, 2}, {71, 1, -2}, {72, 1, -1}, {73, 1, 0}, {74, 1, 1}, {75, 1, 2}, _
        {76, 1.5, -2}, {77, 1.5, -1}, {78, 1.5, 0}, {79, 1.5, 1}, {80, 1.5, 2}, {81, 2, -2}, {82, 2, -1}, {83, 2, 0}, {84, 2, 1}, {85, 2, 2}, _
        {86, 2.5, -2}, {87, 2.5, -1}, {88, 2.5, 0}, {89, 2.5, 1}, {90, 2.5, 2}, {91, 3, -2}, {92, 3, -1}, {93, 3, 0}, {94, 3, 1}, {95, 3, 2}, _
        {96, 3.5, -2}, {97, 3.5, -1}, {98, 3.5, 0}, {99, 3.5, 1}, {100, 3.5, 2}, {101, 4, -2}, {102, 4, -1}, {103, 4, 0}, {104, 4, 1}, {105, 4, 2}, _
        {106, 4.5, -2}, {107, 4.5, -1}, {108, 4.5, 0}, {109, 4.5, 1}, {110, 4.5, 2}, {111, 5, -2}, {112, 5, -1}, {113, 5, 0}, {114, 5, 1}, {115, 5, 2}, _
        {116, 5.5, -2}, {117, 5.5, -1}, {118, 5.5, 0}, {119, 5.5, 1}, {120, 5.5, 2}, {121, 6, -2}, {122, 6, -1}, {123, 6, 0}, {124, 6, 1}, {125, 6, 2}, _
        {126, 6.5, -2}, {127, 6.5, -1}, {128, 6.5, 0}, {129, 6.5, 1}, {130, 6.5, 2}, {131, 7, -2}, {132, 7, -1}, {133, 7, 0}, {134, 7, 1}, {135, 7, 2}, _
        {136, 7.5, -2}, {137, 7.5, -1}, {138, 7.5, 0}, {139, 7.5, 1}, {140, 7.5, 2}, {141, 8, -2}, {142, 8, -1}, {143, 8, 0}, {144, 8, 1}, {145, 8, 2}, _
        {146, 8.5, -2}, {147, 8.5, -1}, {148, 8.5, 0}, {149, 8.5, 1}, {150, 8.5, 2}, {151, 9, -2}, {152, 9, -1}, {153, 9, 0}, {154, 9, 1}, {155, 9, 2}, _
        {156, 9.5, -2}, {157, 9.5, -1}, {158, 9.5, 0}, {159, 9.5, 1}, {160, 9.5, 2}, {161, 10, -2}, {162, 10, -1}, {163, 10, 0}, {164, 10, 1}, {165, 10, 2} _
    }

    'This had to be made after the update of splitting the one table that was used for the holes to 2 tables
    Private Function findControl(ByVal name As String, ByVal hole As Integer) As Control
        'Basically, if the hole number is more than 9, then it'll get the data from the second table, if
        'it's less than 9, then the data will be gotten from the first table
        If hole > 9 Then
            Return TableLayoutPanel2.Controls.Find(name, False)(0)
        Else
            Return TableLayoutPanel1.Controls.Find(name, False)(0)
        End If
    End Function

    Private Sub clearPlayerScoresBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clearPlayerScoresBtn.Click
        'This will loop through all the player score counters and set them to 0
        For i As Integer = 18 To 1 Step -1
            findControl("nud" & i, i).Text = "0"
        Next
    End Sub

    Private Sub doCalculationsBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles doCalculationsBtn.Click
        'Create all the Integer variables we'll be using 
        Dim parTotal, grossScore, adjustBy, handicap, worseHolesInt, netScore, front9, back9 As Integer

        'Create the amount of worse holes to be used variable, must be double because it's a decimal 
        Dim worseHolesToBeUsed As Double

        'Create a true/false variable that will tell us if worse holes has a .5
        Dim hasHalf As Boolean

        'These will be used to record the spreads, the clone will be used to do sort them so original is unaffected
        Dim spreads(16), spreadClone(16) As Integer
        Dim spreadDex As Integer = 0

        'This will give values to gross score, part total and calculate some spreads
        For i As Integer = 1 To 18
            grossScore += Integer.Parse(findControl("nud" & i, i).Text)
            parTotal += Integer.Parse(findControl("NumericUpDown" & i, i).Text)
            If i <= 9 Then
                front9 += Integer.Parse(findControl("nud" & i, i).Text)
            Else
                back9 += Integer.Parse(findControl("nud" & i, i).Text)
            End If
            If i <> 18 Then
                spreads(spreadDex) = Integer.Parse(findControl("nud" & i, i).Text) - Integer.Parse(findControl("NumericUpDown" & i, i).Text)
                spreadDex += 1
            End If
        Next

        'Exit out of the sub if gross score is under 66 or over 165
        If grossScore < 66 Or grossScore > 165 Then
            MsgBox("Your gross score is: " & grossScore & ". Sorry this program does not support that score")
            Return
        End If

        'Get the adjust by from the data
        adjustBy = data(grossScore - 66, 2)
        'Get the amount of worse holes to be used from the data
        worseHolesToBeUsed = data(grossScore - 66, 1)

        'Check if worse holes to be used is a decimal, if it is then enable the half boolean
        If worseHolesToBeUsed.ToString.Contains(".") Then hasHalf = True
        'Math.Ceiling will round up the worse number of holes to be used
        worseHolesInt = Math.Ceiling(worseHolesToBeUsed)

        'Clone the spreads array then sort the clone, we will use the original later on
        spreadClone = spreads.Clone
        Array.Sort(spreadClone)

        'This is where only the worse spreads will be stored
        Dim worseSpreads(worseHolesInt - 1) As Integer
        spreadDex = 0

        'The highest will be closer to the end of the array when sorted, so to find the worse spreads, we take whatever amount we need from the end of the array
        For i As Integer = 16 To 16 - (worseHolesInt - 1) Step -1
            worseSpreads(spreadDex) = spreadClone(i)
            spreadDex += 1
        Next

        'This is where worse holes will be stored
        Dim worseHoles(worseHolesInt - 1), worseHoleScores(worseHolesInt - 1) As Integer
        Dim worseHoleDex As Integer

        'This will calculate the worse holes by using the worse spreads and then using the normal spread array as well. 
        'It also uses the isWorseHoleBySpread method.
        For i As Integer = 0 To worseSpreads.Length - 1
            For j As Integer = 0 To spreads.Length - 1
                If worseSpreads(i) = spreads(j) And isWorseHoleBySpread(worseSpreads(i), worseHoles, j + 1) Then
                    'The hole is j + 1, so that is what will be added to worse holes array when the worst hole is found
                    worseHoles(worseHoleDex) = j + 1
                    'Find Worse Hole Score
                    worseHoleScores(worseHoleDex) = Integer.Parse(findControl("nud" & (j + 1), (j + 1)).Text)
                    'Since we added new data, increment the index counter
                    worseHoleDex += 1
                    'Set to -100 so this will not be used again
                    spreads(j) = -100
                    'This will only exit the current loop, not both
                    Exit For
                End If
            Next
        Next

        'This will add to the handicap by using the data that was collected
        Dim toAdd As Integer
        For i As Integer = 0 To worseSpreads.Length - 1
            If (Integer.Parse(findControl("nud" & worseHoles(i), worseHoles(i)).Text)) > ((Integer.Parse(findControl("NumericUpDown" & worseHoles(i), worseHoles(i)).Text) * 2) + 1) Then
                toAdd = ((Integer.Parse(findControl("NumericUpDown" & worseHoles(i), worseHoles(i)).Text) * 2) + 1)
                If i = (worseSpreads.Length - 1) And hasHalf Then toAdd = Math.Ceiling(toAdd / 2)
                handicap += toAdd
            Else
                toAdd = Integer.Parse(findControl("nud" & worseHoles(i), worseHoles(i)).Text)
                If i = (worseSpreads.Length - 1) And hasHalf Then toAdd = Math.Ceiling(toAdd / 2)
                handicap += toAdd
            End If
        Next

        'Add adjust by to the handicap 
        handicap += adjustBy

        'Calculate final net score with the handicap
        netScore = grossScore - handicap

        'Output everything
        If outBox.Text <> "" Then
            outBox.AppendText(vbNewLine & "Player Name: " & playerNameTxt.Text)
        Else
            outBox.AppendText("Player Name: " & playerNameTxt.Text)
        End If
        outBox.AppendText(vbNewLine & "Course Par Total: " & parTotal)
        outBox.AppendText(vbNewLine & "Front 9 Gross Score: " & front9)
        outBox.AppendText(vbNewLine & "Back 9 Gross Score: " & back9)
        outBox.AppendText(vbNewLine & "18 Hole Gross Score: " & grossScore)
        outBox.AppendText(vbNewLine & "Amount of bad holes to be used: " & worseHolesToBeUsed)
        outBox.AppendText(vbNewLine & "Worse Holes: " & arrayToString(worseHoles))
        outBox.AppendText(vbNewLine & "Worse Hole Scores: " & arrayToString(worseHoleScores))
        outBox.AppendText(vbNewLine & "Worse Spreads: " & arrayToString(worseSpreads))
        outBox.AppendText(vbNewLine & "Adjust By: " & adjustBy)
        outBox.AppendText(vbNewLine & "Handicap: " & handicap)
        outBox.AppendText(vbNewLine & "Net Score: " & netScore)
        outBox.AppendText(vbNewLine & "-")

        'Scroll outbox all the way down
        outBox.ScrollToCaret()

        'Save output box text to file 
        outBox.SaveFile(fileName)

        'Add to save data. Since net score is first then gross score, when it sorts, it will sort by lowest net score and ties will be broken by the gross score
        savedData.Add(netScore & "|" & grossScore & "|" & id & "|" & playerNameTxt.Text & "|" & parTotal & "|" & worseHolesToBeUsed & "|" & arrayToString(worseHoles) & "|" & arrayToString(worseSpreads) & "|" & adjustBy & "|" & handicap & "|" & front9 & "|" & back9 & "|" & arrayToString(worseHoleScores))
        id += 1

        'If there is a save file name, then add the * to the title bar.
        If dataSaveFileName <> Nothing Then Me.Text = "Casual Handi-Cap Golf Tool 2011 | Kush Patel | (*" & dataSaveFileName & ")"
    End Sub

    'This function will return if the current hole is the worse possible hole to add to the worse holes at the moment
    'au stands for already used
    Function isWorseHoleBySpread(ByVal spread As Integer, ByVal alreadyUsedArr As Array, ByVal holeToCheck As Integer)
        Dim au As Boolean
        For i As Integer = 1 To 17
            If (spread = (Integer.Parse(findControl("nud" & i, i).Text) - Integer.Parse(findControl("NumericUpDown" & i, i).Text))) _
                And Integer.Parse(findControl("nud" & i, i).Text) > Integer.Parse(findControl("nud" & holeToCheck, holeToCheck).Text) Then
                For j As Integer = 0 To alreadyUsedArr.Length - 1
                    If i = alreadyUsedArr(j) Then
                        au = True
                    End If
                Next
                If i <> holeToCheck And au = False Then
                    Return False
                End If
            End If
            au = False
        Next
        Return True
    End Function

    'This function will convert all the items in an array to a string by seperating items by a comma
    Function arrayToString(ByVal array As Array) As String
        Dim newString As String = ""
        For i As Integer = 0 To array.Length - 1
            newString &= array(i)
            If i <> array.Length - 1 Then newString += ", "
        Next
        Return newString
    End Function

    'Clear Name
    Private Sub clearNameBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clearNameBtn.Click
        'Clears the player name
        playerNameTxt.Text = ""
    End Sub

    'Sort
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        'First check if there is data saved before continuing
        If savedData.Count > 0 Then
            'We don't want to modify the original data, so clone it to another variable
            Dim outData As ArrayList = savedData.Clone
            'Now Sort it, it works because of how the data is formatted with Net Score then | then Gross Score
            outData.Sort()
            'Clear the outbox
            outBox.Clear()
            'Loop through all the data strings and log all the data, it will now log in order
            For Each s As String In outData
                Dim strArr As Array = s.Split("|")
                If outBox.Text <> "" Then
                    outBox.AppendText(vbNewLine & "Player Name: " & strArr.GetValue(3))
                Else
                    outBox.AppendText("Player Name: " & strArr.GetValue(3))
                End If
                outBox.AppendText(vbNewLine & "Course Par Total: " & strArr.GetValue(4))
                outBox.AppendText(vbNewLine & "Front 9 Gross Score: " & strArr.GetValue(10))
                outBox.AppendText(vbNewLine & "Back 9 Gross Score: " & strArr.GetValue(11))
                outBox.AppendText(vbNewLine & "18 Hole Gross Score: " & strArr.GetValue(1))
                outBox.AppendText(vbNewLine & "Amount of bad holes to be used: " & strArr.GetValue(5))
                outBox.AppendText(vbNewLine & "Worse Holes: " & strArr.GetValue(6))
                outBox.AppendText(vbNewLine & "Worse Hole Scores: " & strArr.GetValue(12))
                outBox.AppendText(vbNewLine & "Worse Spreads: " & strArr.GetValue(7))
                outBox.AppendText(vbNewLine & "Adjust By: " & strArr.GetValue(8))
                outBox.AppendText(vbNewLine & "Handicap: " & strArr.GetValue(9))
                outBox.AppendText(vbNewLine & "Net Score: " & strArr.GetValue(0))
                outBox.AppendText(vbNewLine & "-")
            Next
            'Scroll outbox down
            outBox.ScrollToCaret()
            'Save everything in the output box to the file
            outBox.SaveFile(fileName)
            'Let user know sorting is complete
            MsgBox("Sorting process is complete! Check the output box and the saved RTF file!")
        Else
            'Let the user know there's no data saved
            MsgBox("Sorry, the program has not recorded any information and therefore cannot sort at the moment, please do calculations before trying to sort again")
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem2.Click
        'Make sure there is saved data in the program
        If savedData.Count = 0 Then
            MsgBox("There is nothing to save!")
            Exit Sub
        End If
        'Display a save dialog with the only extension as hdf
        Dim saveBox As New SaveFileDialog
        saveBox.InitialDirectory = FileIO.FileSystem.CurrentDirectory
        saveBox.CreatePrompt = True
        saveBox.OverwritePrompt = True
        saveBox.DefaultExt = "hdf"
        saveBox.Filter = "Handi-Cap Golf Data Files (*.hdf)|*.hdf"
        Dim diagResult As DialogResult = saveBox.ShowDialog()
        'If the user didn't cancel out, then proceed
        If diagResult = Windows.Forms.DialogResult.Yes Or diagResult = Windows.Forms.DialogResult.OK Then
            'Get the file name/path that the user gave to the save box
            dataSaveFileName = saveBox.FileName
            'If the same file exists, delete it.
            If FileIO.FileSystem.FileExists(dataSaveFileName) Then FileIO.FileSystem.DeleteFile(dataSaveFileName)
            'Save the data
            Dim f As New IO.StreamWriter(dataSaveFileName, True)
            For Each s As String In savedData
                f.WriteLine(s)
            Next
            'Windows can see changes
            f.Flush()
            'Close the writer
            f.Close()
            'Let the user know the save is complete
            MsgBox("Save Complete!")
            'Update the title bar
            Me.Text = "Casual Handi-Cap Golf Tool 2011 | Kush Patel | (" & dataSaveFileName & ")"
            'Enable the save option under the File menu
            SaveAsToolStripMenuItem1.Enabled = True
        End If
    End Sub

    'This bundle of code is ran when the form is being closed
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'The * means there is unsaved data, so if the title contains an *, we will run the following code before closing
        If Me.Text.Contains("*") Then
            'We will ask the user if they want to save the un-saved data.
            Dim m As MsgBoxResult = MsgBox("You have un-saved data, would you like to save ?", MsgBoxStyle.YesNo, "Un-saved Data!")
            'if they say yes, then we will run the save file code.
            If m = MsgBoxResult.Yes Then
                'Delete old
                FileIO.FileSystem.DeleteFile(dataSaveFileName)
                'Make new
                Dim f As New IO.StreamWriter(dataSaveFileName, True)
                'Write data
                For Each s As String In savedData
                    f.WriteLine(s)
                Next
                'This will display changes in windows
                f.Flush()
                'Close the writer
                f.Close()
                'Let the user know, save has been completed
                MsgBox("Save Complete!")
                'Update title bar (Seems pointless since program is closing after this line anyways, but oh well)
                Me.Text = "Casual Handi-Cap Golf Tool 2011 | Kush Patel | (" & dataSaveFileName & ")"
            End If
        End If
    End Sub

    'File -> Save
    Private Sub SaveAsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem1.Click
        'First we have to delete the current file.
        FileIO.FileSystem.DeleteFile(dataSaveFileName)
        'Then we create it again and write everything in saved data into the newly created file
        Dim f As New IO.StreamWriter(dataSaveFileName, True)
        For Each s As String In savedData
            f.WriteLine(s)
        Next
        'Flush will make the changes appear in windows
        f.Flush()
        'Must always close
        f.Close()
        'Change the title bar text to get rid of the * in title bar which means unsaved data, now it's all saved so we don't need it
        Me.Text = "Casual Handi-Cap Golf Tool 2011 | Kush Patel | (" & dataSaveFileName & ")"
    End Sub

    Private Sub openToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles openToolStripMenuItem.Click
        'Open a new open dialog, set the only file extension to hdf
        Dim openBox As New OpenFileDialog
        openBox.InitialDirectory = FileIO.FileSystem.CurrentDirectory
        openBox.DefaultExt = "hdf"
        openBox.Filter = "Handi-Cap Golf Data Files (*.hdf)|*.hdf"
        Dim openBoxResult As DialogResult = openBox.ShowDialog
        'If the user didn't cancel out of it, then proceed to run the open code
        If openBoxResult = Windows.Forms.DialogResult.OK Then
            'First save the file name from what the user opened
            dataSaveFileName = openBox.FileName
            'We have to construct the array list again so if there's anything in it then it is emptied.
            savedData = New ArrayList()
            'Each line of the file = a save data line, so we just read a line and add that read line to the saved data
            Dim reader As New IO.StreamReader(dataSaveFileName)
            While Not reader.EndOfStream
                savedData.Add(reader.ReadLine)
            End While
            'Must close the reader to prevent memory leak
            reader.Close()
            'Clear the out put box
            outBox.Clear()
            'Loop through all the data strings and log all the data that was just opened from the save file
            For Each s As String In savedData
                Dim strArr As Array = s.Split("|")
                If outBox.Text <> "" Then
                    outBox.AppendText(vbNewLine & "Player Name: " & strArr.GetValue(3))
                Else
                    outBox.AppendText("Player Name: " & strArr.GetValue(3))
                End If
                outBox.AppendText(vbNewLine & "Course Par Total: " & strArr.GetValue(4))
                outBox.AppendText(vbNewLine & "Front 9 Gross Score: " & strArr.GetValue(10))
                outBox.AppendText(vbNewLine & "Back 9 Gross Score: " & strArr.GetValue(11))
                outBox.AppendText(vbNewLine & "18 Hole Gross Score: " & strArr.GetValue(1))
                outBox.AppendText(vbNewLine & "Amount of bad holes to be used: " & strArr.GetValue(5))
                outBox.AppendText(vbNewLine & "Worse Holes: " & strArr.GetValue(6))
                outBox.AppendText(vbNewLine & "Worse Hole Scores: " & strArr.GetValue(12))
                outBox.AppendText(vbNewLine & "Worse Spreads: " & strArr.GetValue(7))
                outBox.AppendText(vbNewLine & "Adjust By: " & strArr.GetValue(8))
                outBox.AppendText(vbNewLine & "Handicap: " & strArr.GetValue(9))
                outBox.AppendText(vbNewLine & "Net Score: " & strArr.GetValue(0))
                outBox.AppendText(vbNewLine & "-")
            Next
            'Set the title of the form to include the save file name
            Me.Text = "Casual Handi-Cap Golf Tool 2011 | Kush Patel | (" & dataSaveFileName & ")"
            'Enable the save tool item under File menu. Apparantly I didn't rename it to save from SaveAs 
            SaveAsToolStripMenuItem1.Enabled = True
        End If
    End Sub
End Class
