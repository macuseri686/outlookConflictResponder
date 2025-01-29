Private Const WORK_START_HOUR As Integer = 9  ' 9 AM
Private Const WORK_END_HOUR As Integer = 17   ' 5 PM
Private mConflictType As String

Public WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    Set InboxItems = Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler
    
    Debug.Print "New item received"
    
    If TypeOf Item Is MeetingItem Then
        Dim meetingRequest As Outlook.MeetingItem
        Set meetingRequest = Item
        
        Debug.Print "Processing meeting request: " & meetingRequest.Subject
        
        ' -- Skip if the subject contains "Canceled" (case-insensitive)
        If InStr(1, meetingRequest.Subject, "Canceled", vbTextCompare) > 0 Then
            Debug.Print "Skipping canceled meeting"
            Exit Sub
        End If
        
        ' Get the proposed AppointmentItem
        Dim proposedAppt As Outlook.AppointmentItem
        Set proposedAppt = meetingRequest.GetAssociatedAppointment(True)
        
        If Not proposedAppt Is Nothing Then
            Debug.Print "Checking for conflicts..."
            Dim hasConflict As Boolean
            hasConflict = ConflictExists(proposedAppt.Start, proposedAppt.End, meetingRequest)
            
            If hasConflict Then
                Debug.Print "Conflict found, sending notification"
                SendConflictNotification meetingRequest
            Else
                Debug.Print "No conflict found"
            End If
        Else
            Debug.Print "Could not get associated appointment"
        End If
    End If
    Exit Sub

ErrorHandler:
    Debug.Print "Error in ItemAdd: " & Err.Description
End Sub

Function ConflictExists(startTime As Date, endTime As Date, meetingRequest As Outlook.MeetingItem) As Boolean
    On Error GoTo ErrorHandler
    
    ' First check if meeting is outside working hours (9 AM - 5 PM Pacific Time, Monday-Friday)
    If IsOutsideWorkingHours(startTime, endTime) Then
        ConflictExists = True
        ' Store conflict type in module-level variable instead of Tag
        SetConflictType "OutsideHours"
        Exit Function
    End If
    
    ' Reset conflict type
    SetConflictType "Schedule"
    
    Debug.Print "----------------------------------------"
    Debug.Print "Checking conflicts for meeting: " & meetingRequest.Subject
    Debug.Print "Time range: " & Format(startTime, "mm/dd/yyyy hh:mm AMPM") & " to " & Format(endTime, "mm/dd/yyyy hh:mm AMPM")
    
    ' Initialize conflict flag
    ConflictExists = False
    
    ' Get calendar items
    Dim oCalendar As Outlook.Folder
    Set oCalendar = Session.GetDefaultFolder(olFolderCalendar)
    
    Dim oItems As Outlook.Items
    Set oItems = oCalendar.Items
    oItems.IncludeRecurrences = True
    
    ' Filter for items on the same day
    Dim sFilter As String
    sFilter = "[Start] >= '" & Format(DateSerial(Year(startTime), Month(startTime), Day(startTime)), "mm/dd/yyyy") & _
              "' AND [Start] < '" & Format(DateSerial(Year(startTime), Month(startTime), Day(startTime) + 1), "mm/dd/yyyy") & "'"
              
    Debug.Print "Using filter: " & sFilter
    
    Dim oFilteredItems As Outlook.Items
    Set oFilteredItems = oItems.Restrict(sFilter)
    
    Debug.Print "Found " & oFilteredItems.Count & " items on this day"
    
    ' Check each appointment
    Dim oAppt As Object
    For Each oAppt In oFilteredItems
        If TypeOf oAppt Is AppointmentItem Then
            On Error Resume Next  ' Handle any property access errors
            
            Dim appt As AppointmentItem
            Set appt = oAppt
            
            Debug.Print vbNewLine & "Checking appointment: " & appt.Subject
            Debug.Print "Time: " & Format(appt.Start, "hh:mm AMPM") & " to " & Format(appt.End, "hh:mm AMPM")
            Debug.Print "Busy Status: " & appt.BusyStatus & " (0=Free, 1=Tentative, 2=Busy, 3=OOF)"
            
            ' Skip if it's the same meeting
            Dim isSameMeeting As Boolean
            isSameMeeting = False
            
            On Error Resume Next
            If Not meetingRequest.GetAssociatedAppointment(False) Is Nothing Then
                isSameMeeting = (meetingRequest.GetAssociatedAppointment(False).EntryID = appt.EntryID)
            End If
            On Error GoTo ErrorHandler
            
            If isSameMeeting Then
                Debug.Print "Same meeting - skipping"
            Else
                ' Check for time overlap and busy status
                If (appt.End > startTime) And (appt.Start < endTime) Then
                    ' Consider Busy, OOF, and Tentative as conflicts
                    If appt.BusyStatus = olBusy Or appt.BusyStatus = olOutOfOffice Or appt.BusyStatus = olTentative Then
                        Debug.Print "CONFLICT FOUND with " & appt.Subject & "!"
                        Debug.Print "Conflict details: Meeting " & Format(startTime, "hh:mm AMPM") & "-" & Format(endTime, "hh:mm AMPM") & _
                                  " overlaps with " & appt.Subject & " (" & Format(appt.Start, "hh:mm AMPM") & "-" & Format(appt.End, "hh:mm AMPM") & _
                                  ", Status: " & Choose(appt.BusyStatus + 1, "Free", "Tentative", "Busy", "Out of Office") & ")"
                        ConflictExists = True
                        Exit For
                    Else
                        Debug.Print "Overlapping but free (Status: " & appt.BusyStatus & ")"
                    End If
                Else
                    Debug.Print "No overlap"
                End If
            End If
        End If
    Next oAppt
    
    Debug.Print vbNewLine & "Final result: " & ConflictExists
    Debug.Print "----------------------------------------"
    Exit Function

ErrorHandler:
    Debug.Print "Error in ConflictExists: " & Err.Description & " (Line: " & Erl & ")"
    ConflictExists = False
End Function

Function IsOutsideWorkingHours(startTime As Date, endTime As Date) As Boolean
    ' Remove the constants from here since they're now at module level
    
    ' Convert times to local time if they aren't already
    Dim localStartTime As Date
    Dim localEndTime As Date
    localStartTime = startTime
    localEndTime = endTime
    
    ' Check if meeting is on weekend
    If Weekday(localStartTime) = vbSaturday Or Weekday(localStartTime) = vbSunday Then
        Debug.Print "Meeting is on weekend"
        IsOutsideWorkingHours = True
        Exit Function
    End If
    
    ' Check if meeting starts before work hours or ends after work hours
    If Hour(localStartTime) < WORK_START_HOUR Or Hour(localEndTime) > WORK_END_HOUR Then
        Debug.Print "Meeting is outside work hours"
        IsOutsideWorkingHours = True
        Exit Function
    End If
    
    IsOutsideWorkingHours = False
End Function

Private Sub SetConflictType(conflictType As String)
    mConflictType = conflictType
End Sub

Private Function GetConflictType() As String
    GetConflictType = mConflictType
End Function

Sub SendConflictNotification(meetingReq As Outlook.MeetingItem)
    Dim mail As Outlook.MailItem
    Dim proposedAppt As Outlook.AppointmentItem
    Dim meetingSubject As String
    Dim messageBody As String
    Dim workHoursText As String
    
    ' Format working hours text
    workHoursText = FormatWorkHours(WORK_START_HOUR, WORK_END_HOUR)
    
    ' Get the AppointmentItem object from the meeting request
    Set proposedAppt = meetingReq.GetAssociatedAppointment(True)
    
    If Not proposedAppt Is Nothing Then
        meetingSubject = proposedAppt.Subject
    Else
        meetingSubject = meetingReq.Subject
    End If
    
    Set mail = Application.CreateItem(olMailItem)
    mail.To = meetingReq.SenderEmailAddress
    mail.Subject = "[Auto-Notification] Meeting Invitation Conflict"
    
    ' Check the type of conflict using the module-level variable
    If GetConflictType() = "OutsideHours" Then
        messageBody = _
            "Your requested meeting, """ & meetingSubject & """, is scheduled outside of working hours (" & workHoursText & ", Monday-Friday) for this calendar (caleb.banzhaf1@t-mobile.com)." & vbCrLf & vbCrLf & _
            "If this event is non-work related and the time is intentional, you may disregard this message. Otherwise, please consider scheduling during business hours." & vbCrLf & vbCrLf & _
            "This is an automated notification from Outlook." & vbCrLf & _
            "No further action is required on your part until a new meeting request is sent."
    Else
        messageBody = _
            "Your requested meeting, """ & meetingSubject & """, conflicts with an existing appointment on this calendar (caleb.banzhaf1@t-mobile.com)." & vbCrLf & vbCrLf & _
            "Please use Outlook's Scheduling Assistant to view available time slots for all meeting participants and propose a different time to avoid overlap." & vbCrLf & vbCrLf & _
            "This is an automated notification from Outlook." & vbCrLf & _
            "No further action is required on your part until a new meeting request is sent."
    End If
    
    mail.Body = messageBody
    mail.Send
End Sub

Function FormatWorkHours(startHour As Integer, endHour As Integer) As String
    Dim startTime As Date
    Dim endTime As Date
    
    ' Create dummy dates just to format the times
    startTime = DateSerial(2000, 1, 1) + TimeSerial(startHour, 0, 0)
    endTime = DateSerial(2000, 1, 1) + TimeSerial(endHour, 0, 0)
    
    ' Format as "8:00 AM - 5:00 PM"
    FormatWorkHours = Format(startTime, "h:mm AM/PM") & " - " & Format(endTime, "h:mm AM/PM")
End Function
