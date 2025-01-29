Public WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    Set InboxItems = Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    
    If TypeOf Item Is MeetingItem Then
        Dim meetingRequest As Outlook.MeetingItem
        Set meetingRequest = Item
        
        ' -- Skip if the subject contains "Canceled" (case-insensitive)
        If InStr(1, meetingRequest.Subject, "Canceled", vbTextCompare) > 0 Then
            ' It's a canceled meeting request; do nothing
            Exit Sub
        End If
        
        ' Get the proposed AppointmentItem
        Dim proposedAppt As Outlook.AppointmentItem
        Set proposedAppt = meetingRequest.GetAssociatedAppointment(True)
        
        If Not proposedAppt Is Nothing Then
            ' Check for conflict
            If ConflictExists(proposedAppt.Start, proposedAppt.End) Then
                ' Send your automated response
                SendConflictNotification meetingRequest
            End If
        End If
    End If
End Sub

Function ConflictExists(startTime As Date, endTime As Date) As Boolean
    ' Query your calendar to find overlapping items
    Dim oCalendar As Outlook.Folder
    Set oCalendar = Session.GetDefaultFolder(olFolderCalendar)

    Dim oItems As Outlook.Items
    Set oItems = oCalendar.Items

    ' Restrict items to those that might overlap
    oItems.Sort "[Start]"
    oItems.IncludeRecurrences = True
    
    Dim strFilter As String
    strFilter = "[Start] < '" & Format(endTime, "ddddd hh:mm") & "' AND [End] > '" & Format(startTime, "ddddd hh:mm") & "'"
    
    Dim conflictedItems As Outlook.Items
    Set conflictedItems = oItems.Restrict(strFilter)
    
    If conflictedItems.Count > 0 Then
        ConflictExists = True
    Else
        ConflictExists = False
    End If
End Function

Sub SendConflictNotification(meetingReq As Outlook.MeetingItem)
    Dim mail As Outlook.MailItem
    Dim proposedAppt As Outlook.AppointmentItem
    Dim meetingSubject As String
    
    ' Get the AppointmentItem object from the meeting request
    Set proposedAppt = meetingReq.GetAssociatedAppointment(True)
    
    If Not proposedAppt Is Nothing Then
        meetingSubject = proposedAppt.Subject
    Else
        ' Fallback: use the meeting request subject if we can't get the appointment
        meetingSubject = meetingReq.Subject
    End If
    
    Set mail = Application.CreateItem(olMailItem)
    mail.To = meetingReq.SenderEmailAddress
    mail.Subject = "[Auto-Notification] Meeting Invitation Conflict"

    ' DONT FORGET TO ADD YOUR ACTUAL EMAIL HERE INSTEAD OF email@email.com
    mail.Body = _
        "Your requested meeting, """ & meetingSubject & """, conflicts with an existing appointment on this calendar (email@email.com)." & vbCrLf & vbCrLf & _
        "Please use Outlook's Scheduling Assistant to view available time slots for all meeting participants and propose a different time to avoid overlap." & vbCrLf & vbCrLf & _
        "This is an automated notification from Outlook." & vbCrLf & _
        "No further action is required on your part until a new meeting request is sent."
    
    mail.Send
End Sub

