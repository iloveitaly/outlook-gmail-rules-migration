Sub GetAllRules()
    Dim st As Outlook.Store
    Dim myRules As Outlook.Rules
    Dim rl As Outlook.Rule
    Dim count As Integer
    Dim ruleList As String

    Set st = Application.Session.DefaultStore

    ' get rules
    Set myRules = st.GetRules
    Dim ruleInfo As String

    ' iterate all the rules
    For Each rl In myRules
        If rl.Conditions.From.Recipients.count > 0 Then
            ' Condition: From Email
            Dim adType
            adType = rl.Conditions.From.Recipients.Item(1).AddressEntry.AddressEntryUserType
            If adType = olExchangeRemoteUserAddressEntry Or adType = olExchangeUserAddressEntry Then
	            Dim exUserr
	            exUserr = rl.Conditions.From.Recipients.Item(1).AddressEntry.GetExchangeUser.PrimarySmtpAddress
	            ruleList = ruleList & exUserr & ","
	            'ruleList = ruleList & rl.Conditions.From.Recipients.Item(1).AddressEntry.Name
	            '.GetExchangeUser
	            'ruleList = ruleList & "," & exUser.PrimarySmtpAddress
	            'Dim FromEmail As String = exUser.PrimarySmtpAddress
	            'MessageBox.Show(ìFROM:  î & FromEmail)
            ElseIf adType = olSmtpAddressEntry Then
				ruleList = ruleList & rl.Conditions.From.Recipients.Item(1).AddressEntry.Address & ","
            End If
        ElseIf rl.Conditions.SentTo.Enabled Then
            ' CONDITION: TO EMAIL
            'Dim exUser As Outlook.ExchangeUser = rl.Conditions.To.Recipients.Item(1).AddressEntry.GetExchangeUser
            'Dim ToEmail As String = exUser.PrimarySmtpAddress
            'MessageBox.Show(ìTO: î & ToEmail)
        ElseIf rl.Conditions.Subject.Enabled Then
            ' Condition: Subject
            'MessageBox.Show(ìSubject: î & olRule.Conditions.Subject.Text)
        End If

        ' ACTION: MOVE TO FOLDER
        'If rl.RuleType = olRuleReceive
        'Debug.Print rl.Actions.Item(1).ActionType = olRuleActionMoveToFolder

        If rl.Actions.Item(1).ActionType = olRuleActionMoveToFolder Then
            If rl.Actions.MoveToFolder.Enabled Then
               ruleList = ruleList & rl.Actions.MoveToFolder.Folder.FolderPath & vbCrLf
            End If
            'MessageBox.Show ("Move To Folder: " & rl.Actions.Item(1).MoveToFolder.Folder.FolderPath)
        End If

        ' determine if it's an Inbox rule
        'If rl.RuleType = olRuleReceive Then
            ' if so, run it
         '   count = count + 1
          '  ruleList = ruleList & vbCrLf & rl.Name
           ' If rl.Conditions.From.Recipients.count > 0 Then
          '      ruleInfo = ruleInfo & rl.Conditions.From.Recipients.Item(1).Name
           ' End If
        'End If
        'Exit For
    Next

    ' tell the user what you did
    'ruleList = "These rules were executed against the Inbox: " & vbCrLf & ruleList
    'MsgBox ruleList, vbInformation, "Macro: RunAllInboxRules"
    MsgBox ruleList
    Set rl = Nothing
    Set st = Nothing
    Set myRules = Nothing

	'Write information to Text File
	Dim Stuff, myFSO, WriteStuff
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	Set WriteStuff = myFSO.OpenTextFile("c:\rules.csv", 8, True)
	WriteStuff.WriteLine (ruleList)
	WriteStuff.Close
	Set WriteStuff = Nothing
	Set myFSO = Nothing
End Sub
