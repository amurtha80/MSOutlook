Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'
'Paste this macro into your "ThisOutlookSession" module.
'DISCLAIMER: This was only tested on Outlook 2003
'
'ADDITIONAL CREDITS:
'This macro was created by modifying Mark Bird's original code which
'can be found here: http://mark.bird.googlepages.com/, fixing some
'issues related to search strings on replies/forwards, and also
'merging in additional features/checks taken from examples on Sue
'Mosher's OutlookCode.com. See examples here:
'http://www.outlookcode.com/codedetail.aspx?id=539 and here:
'http://www.outlookcode.com/codedetail.aspx?id=1278
'
Dim m As Variant
Dim strBody As String
Dim intIn, intLength As Long
Dim intAttachCount As Integer, intStandardAttachCount As Integer
'
On Error GoTo ErrorHandler
'
'You may have a picture or vCard in your email signature that you
'don't want to be counted when checking for attachments. If so, then
'edit the following line to make intStandardAttachCount equal the
'number of files attached in your signature.
intStandardAttachCount = 0
'
'CHECK #1: Check for a blank subject line
If Item.Subject = "" Then
'Extra spaces added to the messages just to
'keep them centered and pretty
m = MsgBox("The subject line is blank... " & _
vbNewLine & vbNewLine & _
"Do you still want to send this message? ", _
vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground, "Blank Subject")
If m = vbNo Then
Cancel = True
GoTo ExitSub
End If
End If
'
'CHECK #2: Check for a missing attachment
intIn = 0
strBody = LCase(Item.Subject) & LCase(Item.Body)
'If the message is a reply or forward, then the macro will
'not search for the strings in the original message. Anything
'below the "from:" line is ignored
intLength = InStr(1, strBody, "from:")
If intLength = 0 Then intLength = Len(strBody)
'
'Add lines for every string you want to check, including other
'languages, etc. Partial strings are fine. For example, "attach"
'will match "attached" & "attachment"
If intIn = 0 Then intIn = InStr(1, Left(strBody, intLength), "attach")
If intIn = 0 Then intIn = InStr(1, Left(strBody, intLength), "enclosed")
'
intAttachCount = Item.Attachments.Count
If intIn > 0 And intAttachCount <= intStandardAttachCount Then
m = MsgBox("It looks like you forgot to attach a file... " _
& vbNewLine & vbNewLine & _
"Do you still want to send this message? ", _
vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground, "Attachment Missing?")
If m = vbNo Then
Cancel = True
GoTo ExitSub
End If
End If
'
'CHECK #3: Check for meeting requests with no location
If Item.Class = olMeetingRequest Then
If InStr(1, Item.Body, "Where:", vbTextCompare) = 0 Then
m = MsgBox("The meeting location is blank... " _
& vbNewLine & vbNewLine & _
"Do you still want to send this meeting invite? ", _
vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground, "Blank Location")
If m = vbNo Then
Cancel = True
GoTo ExitSub
End If
End If
End If
'
ExitSub:
Set Item = Nothing
strBody = ""
Exit Sub
'
ErrorHandler:
MsgBox "Send Checker" & vbCrLf & vbCrLf _
& "Error Code: " & Err.Number & vbCrLf & Err.Description
Err.Clear
GoTo ExitSub
End Sub
