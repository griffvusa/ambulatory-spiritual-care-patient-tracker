Attribute VB_Name = "modPatients"
Option Explicit

' ── Status label mapping (internal snake_case <-> display label) ──────────────
Public Function StatusLabel(internalValue As String) As String
    Select Case LCase(internalValue)
        Case "ready_to_schedule": StatusLabel = "Ready to Schedule"
        Case "scheduled": StatusLabel = "Scheduled"
        Case "active": StatusLabel = "Active"
        Case "active_ongoing": StatusLabel = "Active-Ongoing"
        Case "incomplete": StatusLabel = "Incomplete"
        Case "completed": StatusLabel = "Completed"
        Case "dropped": StatusLabel = "Dropped"
        Case "on_hold": StatusLabel = "On Hold"
        Case "deleted": StatusLabel = "Deleted"
        Case Else: StatusLabel = internalValue
    End Select
End Function

Public Function StatusInternal(label As String) As String
    Select Case label
        Case "Ready to Schedule": StatusInternal = "ready_to_schedule"
        Case "Scheduled": StatusInternal = "scheduled"
        Case "Active": StatusInternal = "active"
        Case "Active-Ongoing": StatusInternal = "active_ongoing"
        Case "Incomplete": StatusInternal = "incomplete"
        Case "Completed": StatusInternal = "completed"
        Case "Dropped": StatusInternal = "dropped"
        Case "On Hold": StatusInternal = "on_hold"
        Case "Deleted": StatusInternal = "deleted"
        Case Else: StatusInternal = ""
    End Select
End Function

' All status labels a user may pick from the general Change Status dropdown --
' "deleted" is deliberately excluded, only reachable via the Delete action.
Public Function AllStatusLabelsExceptDeleted() As Collection
    Dim result As New Collection
    result.Add "Ready to Schedule"
    result.Add "Scheduled"
    result.Add "Active"
    result.Add "Active-Ongoing"
    result.Add "Incomplete"
    result.Add "Completed"
    result.Add "Dropped"
    result.Add "On Hold"
    Set AllStatusLabelsExceptDeleted = result
End Function

' ── Reads ───────────────────────────────────────────────────────────────────
Public Function GetPatient(patientId As Long) As Object
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long: r = modUtils.FindById(ws, patientId)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    If r = 0 Then
        Set GetPatient = d
        Exit Function
    End If

    d("id") = modUtils.GetVal(ws, r, "id")
    d("mrn") = modUtils.GetVal(ws, r, "mrn")
    d("last_name") = modUtils.GetVal(ws, r, "last_name")
    d("first_name") = modUtils.GetVal(ws, r, "first_name")
    d("middle_name") = modUtils.GetVal(ws, r, "middle_name")
    d("phone") = modUtils.GetVal(ws, r, "phone")
    d("date_of_birth") = modUtils.GetVal(ws, r, "date_of_birth")
    d("date_of_referral") = modUtils.GetVal(ws, r, "date_of_referral")
    d("referral_from") = modUtils.GetVal(ws, r, "referral_from")
    d("referral_reason") = modUtils.GetVal(ws, r, "referral_reason")
    d("pcp_emailed") = modUtils.GetVal(ws, r, "pcp_emailed")
    d("referral_source_id") = modUtils.GetVal(ws, r, "referral_source_id")
    d("religion_id") = modUtils.GetVal(ws, r, "religion_id")
    d("language_id") = modUtils.GetVal(ws, r, "language_id")
    d("referral_source") = modUtils.LovValueForId(modUtils.SafeCLng(d("referral_source_id")))
    d("religion") = modUtils.LovValueForId(modUtils.SafeCLng(d("religion_id")))
    d("language") = modUtils.LovValueForId(modUtils.SafeCLng(d("language_id")))
    d("current_status") = modUtils.GetVal(ws, r, "current_status")
    d("followup_date") = modUtils.GetVal(ws, r, "followup_date")
    d("is_active") = modUtils.GetVal(ws, r, "is_active")
    d("sdat_begin_score") = modUtils.GetVal(ws, r, "sdat_begin_score")
    d("sdat_begin_date") = modUtils.GetVal(ws, r, "sdat_begin_date")
    d("sdat_end_score") = modUtils.GetVal(ws, r, "sdat_end_score")
    d("sdat_end_date") = modUtils.GetVal(ws, r, "sdat_end_date")
    d("sdat_pct_improvement") = modUtils.GetVal(ws, r, "sdat_pct_improvement")
    d("notes") = modUtils.GetVal(ws, r, "notes")
    Set GetPatient = d
End Function

Public Function GetStatusHistory(patientId As Long) As Collection
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_status_history")
    Dim last As Long: last = modUtils.LastDataRow(ws)
    Dim cPid As Long: cPid = modUtils.ColIndex(ws, "patient_id")
    Dim cStatus As Long: cStatus = modUtils.ColIndex(ws, "status")
    Dim cChangedBy As Long: cChangedBy = modUtils.ColIndex(ws, "changed_by")
    Dim cChangedAt As Long: cChangedAt = modUtils.ColIndex(ws, "changed_at")

    Dim rows As New Collection
    Dim i As Long
    For i = 2 To last
        If ws.Cells(i, cPid).Value = patientId Then
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            d("id") = ws.Cells(i, 1).Value
            d("status") = StatusLabel(CStr(ws.Cells(i, cStatus).Value))
            d("changed_by_name") = modAdmin.GetUserName(CLng(ws.Cells(i, cChangedBy).Value))
            d("changed_at") = ws.Cells(i, cChangedAt).Value
            rows.Add d
        End If
    Next i

    Dim result As New Collection
    Dim n As Long: n = rows.Count
    If n = 0 Then
        Set GetStatusHistory = result
        Exit Function
    End If

    ' DESC by changed_at -- ISO timestamps sort correctly as plain strings.
    Dim arr() As Object: ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = rows(i): Next i
    Dim j As Long
    For i = 1 To n - 1
        For j = 1 To n - i
            If CStr(arr(j)("changed_at")) < CStr(arr(j + 1)("changed_at")) Then
                Dim tmp As Object: Set tmp = arr(j): Set arr(j) = arr(j + 1): Set arr(j + 1) = tmp
            End If
        Next j
    Next i
    For i = 1 To n: result.Add arr(i): Next i
    Set GetStatusHistory = result
End Function

' Mirrors src/main/ipc/patients.js's listPatients filter logic exactly:
' statusFilter="deleted" -> only current_status=deleted rows; a specific status
' -> that status AND is_active=1; "" (All) -> is_active=1 (excludes deleted).
Public Function ListForWorkQueue(statusFilter As String, referralSourceId As Long, searchText As String) As Collection
    Dim result As New Collection
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim last As Long: last = modUtils.LastDataRow(ws)
    Dim cId As Long: cId = modUtils.ColIndex(ws, "id")
    Dim cMrn As Long: cMrn = modUtils.ColIndex(ws, "mrn")
    Dim cLast As Long: cLast = modUtils.ColIndex(ws, "last_name")
    Dim cFirst As Long: cFirst = modUtils.ColIndex(ws, "first_name")
    Dim cReferralDate As Long: cReferralDate = modUtils.ColIndex(ws, "date_of_referral")
    Dim cReferralSrc As Long: cReferralSrc = modUtils.ColIndex(ws, "referral_source_id")
    Dim cLanguage As Long: cLanguage = modUtils.ColIndex(ws, "language_id")
    Dim cStatus As Long: cStatus = modUtils.ColIndex(ws, "current_status")
    Dim cActive As Long: cActive = modUtils.ColIndex(ws, "is_active")
    Dim cFollowup As Long: cFollowup = modUtils.ColIndex(ws, "followup_date")

    Dim search As String: search = LCase(Trim(searchText))
    Dim i As Long
    For i = 2 To last
        Dim status As String: status = CStr(ws.Cells(i, cStatus).Value)
        Dim isActive As Boolean: isActive = (ws.Cells(i, cActive).Value = 1)
        Dim keep As Boolean: keep = True

        If statusFilter = "deleted" Then
            If status <> "deleted" Then keep = False
        ElseIf statusFilter <> "" Then
            If status <> statusFilter Or Not isActive Then keep = False
        Else
            If Not isActive Then keep = False
        End If

        If keep And referralSourceId > 0 Then
            If ws.Cells(i, cReferralSrc).Value <> referralSourceId Then keep = False
        End If

        If keep And search <> "" Then
            Dim ln As String: ln = LCase(CStr(ws.Cells(i, cLast).Value))
            Dim fn As String: fn = LCase(CStr(ws.Cells(i, cFirst).Value))
            Dim mrn As String: mrn = LCase(CStr(ws.Cells(i, cMrn).Value))
            If InStr(ln, search) = 0 And InStr(fn, search) = 0 And InStr(mrn, search) = 0 Then
                keep = False
            End If
        End If

        If keep Then
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            Dim pid As Long: pid = CLng(ws.Cells(i, cId).Value)
            d("id") = pid
            d("last_name") = ws.Cells(i, cLast).Value
            d("first_name") = ws.Cells(i, cFirst).Value
            d("mrn") = ws.Cells(i, cMrn).Value
            d("date_of_referral") = ws.Cells(i, cReferralDate).Value
            d("referral_source") = modUtils.LovValueForId(CLng(ws.Cells(i, cReferralSrc).Value))
            d("language") = modUtils.LovValueForId(CLng(ws.Cells(i, cLanguage).Value))
            d("status_label") = StatusLabel(status)
            d("followup_date") = ws.Cells(i, cFollowup).Value & ""
            d("date_of_referral") = ws.Cells(i, cReferralDate).Value & ""
            If FollowupApplies(status) Then d("followup_reminder") = FollowupReminder(CStr(d("followup_date")))
            d("next_appointment") = NextScheduledAppointmentText(pid)
            result.Add d
        End If
    Next i

    Set ListForWorkQueue = result
End Function

Private Function NextScheduledAppointmentText(patientId As Long) As String
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_appointments")
    Dim last As Long: last = modUtils.LastDataRow(ws)
    Dim cPid As Long: cPid = modUtils.ColIndex(ws, "patient_id")
    Dim cApptDate As Long: cApptDate = modUtils.ColIndex(ws, "date")
    Dim cTime As Long: cTime = modUtils.ColIndex(ws, "time")
    Dim cStatus As Long: cStatus = modUtils.ColIndex(ws, "status")
    Dim today As String: today = modUtils.DateISO(Date)

    Dim bestDate As String: bestDate = ""
    Dim bestTime As String: bestTime = ""
    Dim i As Long
    For i = 2 To last
        If ws.Cells(i, cPid).Value = patientId And CStr(ws.Cells(i, cStatus).Value) = "scheduled" Then
            Dim d As String: d = CStr(ws.Cells(i, cApptDate).Value)
            If d >= today Then
                If bestDate = "" Or d < bestDate Then
                    bestDate = d
                    bestTime = CStr(ws.Cells(i, cTime).Value)
                End If
            End If
        End If
    Next i

    If bestDate = "" Then
        NextScheduledAppointmentText = ""
    Else
        NextScheduledAppointmentText = bestDate & " " & bestTime
    End If
End Function

' A follow-up date only generates reminders while the patient is still in
' care: not for completed or dropped patients (deleted ones are inactive).
Private Function FollowupApplies(status As String) As Boolean
    FollowupApplies = (status <> "completed" And status <> "dropped" And status <> "deleted")
End Function

' "OVERDUE" / "DUE TODAY" / "" for a follow-up date (ISO text compares correctly).
Public Function FollowupReminder(followupDate As String) As String
    If followupDate = "" Then Exit Function
    Dim today As String: today = modUtils.DateISO(Date)
    If followupDate < today Then
        FollowupReminder = "OVERDUE"
    ElseIf followupDate = today Then
        FollowupReminder = "DUE TODAY"
    End If
End Function

' Active patients (not completed/dropped) whose follow-up date is today or earlier.
Public Function FollowupsDueCount() As Long
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim last As Long: last = modUtils.LastDataRow(ws)
    Dim cStatus As Long: cStatus = modUtils.ColIndex(ws, "current_status")
    Dim cActive As Long: cActive = modUtils.ColIndex(ws, "is_active")
    Dim cFollowup As Long: cFollowup = modUtils.ColIndex(ws, "followup_date")
    Dim i As Long, n As Long
    For i = 2 To last
        If ws.Cells(i, cActive).Value = 1 And FollowupApplies(CStr(ws.Cells(i, cStatus).Value)) Then
            If FollowupReminder(CStr(ws.Cells(i, cFollowup).Value)) <> "" Then n = n + 1
        End If
    Next i
    FollowupsDueCount = n
End Function

' Login-time pop-up listing active patients (not completed/dropped) whose
' follow-up date is today or earlier (overdue first). Silent when nothing is due.
Public Sub ShowFollowupReminders()
    Const MAX_LINES As Long = 15
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim last As Long: last = modUtils.LastDataRow(ws)
    Dim cStatus As Long: cStatus = modUtils.ColIndex(ws, "current_status")
    Dim cActive As Long: cActive = modUtils.ColIndex(ws, "is_active")
    Dim cFollowup As Long: cFollowup = modUtils.ColIndex(ws, "followup_date")
    Dim cLast As Long: cLast = modUtils.ColIndex(ws, "last_name")
    Dim cFirst As Long: cFirst = modUtils.ColIndex(ws, "first_name")
    Dim cMrn As Long: cMrn = modUtils.ColIndex(ws, "mrn")

    Dim overdue As String, dueToday As String
    Dim nOverdue As Long, nToday As Long, shown As Long
    Dim i As Long
    For i = 2 To last
        If ws.Cells(i, cActive).Value = 1 And FollowupApplies(CStr(ws.Cells(i, cStatus).Value)) Then
            Dim fu As String: fu = CStr(ws.Cells(i, cFollowup).Value)
            Dim reminder As String: reminder = FollowupReminder(fu)
            If reminder <> "" Then
                Dim entry As String
                entry = "   " & ws.Cells(i, cLast).Value & ", " & ws.Cells(i, cFirst).Value & _
                       "  (MRN " & ws.Cells(i, cMrn).Value & ")  " & modUtils.FmtDate(fu) & vbCrLf
                If reminder = "OVERDUE" Then
                    nOverdue = nOverdue + 1
                    If nOverdue <= MAX_LINES Then overdue = overdue & entry
                Else
                    nToday = nToday + 1
                    If nToday <= MAX_LINES Then dueToday = dueToday & entry
                End If
            End If
        End If
    Next i

    If nOverdue + nToday = 0 Then Exit Sub

    Dim msg As String
    If nOverdue > 0 Then
        msg = msg & "OVERDUE (" & nOverdue & "):" & vbCrLf & overdue
        If nOverdue > MAX_LINES Then msg = msg & "   ...and " & (nOverdue - MAX_LINES) & " more" & vbCrLf
        msg = msg & vbCrLf
    End If
    If nToday > 0 Then
        msg = msg & "DUE TODAY (" & nToday & "):" & vbCrLf & dueToday
        If nToday > MAX_LINES Then msg = msg & "   ...and " & (nToday - MAX_LINES) & " more" & vbCrLf
    End If
    MsgBox msg, vbExclamation, "Follow-up Reminders"
End Sub

Public Function CountsByStatus() As Object
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")
    Dim statuses As Variant: statuses = Array("ready_to_schedule", "scheduled", "active", "active_ongoing", "incomplete", _
                     "completed", "dropped", "on_hold")
    Dim s As Variant
    For Each s In statuses
        counts(CStr(s)) = 0
    Next s
    counts("total_active") = 0

    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim last As Long: last = modUtils.LastDataRow(ws)
    Dim cStatus As Long: cStatus = modUtils.ColIndex(ws, "current_status")
    Dim cActive As Long: cActive = modUtils.ColIndex(ws, "is_active")
    Dim i As Long
    For i = 2 To last
        If ws.Cells(i, cActive).Value = 1 Then
            Dim st As String: st = CStr(ws.Cells(i, cStatus).Value)
            If counts.Exists(st) Then counts(st) = counts(st) + 1
            counts("total_active") = counts("total_active") + 1
        End If
    Next i
    Set CountsByStatus = counts
End Function

' ── Writes ──────────────────────────────────────────────────────────────────
Public Sub Save(patientId As Long, mrn As String, lastName As String, firstName As String, _
                 middleName As String, phone As String, referralDate As String, _
                 referralSourceId As Long, religionId As Long, languageId As Long, _
                 Optional dateOfBirth As String = "", Optional referralFrom As String = "", _
                 Optional referralReason As String = "", Optional pcpEmailed As String = "", _
                 Optional initialStatusDate As String = "")
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long
    Dim newId As Long

    If patientId = 0 Then
        r = modUtils.LastDataRow(ws) + 1
        newId = modUtils.NextId(ws)
        modUtils.SetVal ws, r, "id", newId
        modUtils.SetVal ws, r, "current_status", "ready_to_schedule"
        modUtils.SetVal ws, r, "is_active", 1
        modUtils.SetVal ws, r, "created_at", modUtils.NowISO()
    Else
        r = modUtils.FindById(ws, patientId)
        If r = 0 Then Exit Sub
        newId = patientId
    End If

    modUtils.SetVal ws, r, "mrn", mrn
    modUtils.SetVal ws, r, "last_name", lastName
    modUtils.SetVal ws, r, "first_name", firstName
    modUtils.SetVal ws, r, "middle_name", middleName
    modUtils.SetVal ws, r, "phone", phone
    modUtils.SetVal ws, r, "date_of_birth", dateOfBirth
    modUtils.SetVal ws, r, "referral_from", referralFrom
    modUtils.SetVal ws, r, "referral_reason", referralReason
    modUtils.SetVal ws, r, "pcp_emailed", pcpEmailed
    modUtils.SetVal ws, r, "date_of_referral", referralDate
    modUtils.SetVal ws, r, "referral_source_id", referralSourceId
    modUtils.SetVal ws, r, "religion_id", religionId
    modUtils.SetVal ws, r, "language_id", languageId

    If patientId = 0 Then
        ' initialStatusDate (ISO, "" = now) is when the patient entered "ready to schedule"
        AppendStatusHistory newId, "ready_to_schedule", initialStatusDate
    End If
    modUtils.AutoSave
End Sub

' effectiveDate ("YYYY-MM-DD", "" = now) back-dates the history entry: the
' date is the one given, the time-of-day is the current time so same-day
' entries keep their entry order.
Private Sub AppendStatusHistory(patientId As Long, status As String, Optional effectiveDate As String = "")
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_status_history")
    Dim r As Long: r = modUtils.LastDataRow(ws) + 1
    modUtils.SetVal ws, r, "id", modUtils.NextId(ws)
    modUtils.SetVal ws, r, "patient_id", patientId
    modUtils.SetVal ws, r, "status", status
    modUtils.SetVal ws, r, "changed_by", modAuth.gUserId
    If effectiveDate = "" Then
        modUtils.SetVal ws, r, "changed_at", modUtils.NowISO()
    Else
        modUtils.SetVal ws, r, "changed_at", effectiveDate & " " & Format(Time, "hh:nn:ss")
    End If
End Sub

' Corrects the date of an existing status-history entry. Only the date part of
' changed_at is replaced (the time of day is kept) so same-day ordering holds.
' Returns False if the entry doesn't exist or newDate isn't a valid YYYY-MM-DD.
Public Function UpdateStatusHistoryDate(historyId As Long, newDate As String) As Boolean
    If Not modUtils.IsValidIsoDate(newDate) Or newDate = "" Then Exit Function
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_status_history")
    Dim r As Long: r = modUtils.FindById(ws, historyId)
    If r = 0 Then Exit Function
    Dim existing As String: existing = CStr(modUtils.GetVal(ws, r, "changed_at"))
    Dim timePart As String
    If Len(existing) > 10 Then timePart = Mid(existing, 11) Else timePart = " 00:00:00"
    modUtils.SetVal ws, r, "changed_at", newDate & timePart
    modUtils.AutoSave
    UpdateStatusHistoryDate = True
End Function

' Sets or clears (isoDate = "") a patient's follow-up date, independent of status.
Public Sub SaveFollowupDate(patientId As Long, isoDate As String)
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long: r = modUtils.FindById(ws, patientId)
    If r = 0 Then Exit Sub
    modUtils.SetVal ws, r, "followup_date", isoDate
    modUtils.AutoSave
End Sub

' effectiveDate ("YYYY-MM-DD", "" = now) back-dates the status history entry.
Public Sub ChangeStatus(patientId As Long, newStatus As String, Optional effectiveDate As String = "")
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long: r = modUtils.FindById(ws, patientId)
    If r = 0 Then Exit Sub
    modUtils.SetVal ws, r, "current_status", newStatus
    If newStatus = "deleted" Then
        modUtils.SetVal ws, r, "is_active", 0
    End If
    AppendStatusHistory patientId, newStatus, effectiveDate
    modUtils.AutoSave
End Sub

Public Sub Delete(patientId As Long)
    ChangeStatus patientId, "deleted"
End Sub

Public Sub Restore(patientId As Long)
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long: r = modUtils.FindById(ws, patientId)
    If r = 0 Then Exit Sub
    modUtils.SetVal ws, r, "is_active", 1
    ChangeStatus patientId, "on_hold"
End Sub

' Mirrors patients.js's computeSdatPct exactly: null if either score is missing
' or beginScore=0; else Math.round((begin-end)*1000/begin)/10.
Public Sub SaveSdat(patientId As Long, beginScore As Variant, beginDate As String, _
                     endScore As Variant, endDate As String)
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long: r = modUtils.FindById(ws, patientId)
    If r = 0 Then Exit Sub
    modUtils.SetVal ws, r, "sdat_begin_score", beginScore
    modUtils.SetVal ws, r, "sdat_begin_date", beginDate
    modUtils.SetVal ws, r, "sdat_end_score", endScore
    modUtils.SetVal ws, r, "sdat_end_date", endDate
    modUtils.SetVal ws, r, "sdat_pct_improvement", ComputeSdatPct(beginScore, endScore)
    modUtils.AutoSave
End Sub

Private Function ComputeSdatPct(beginScore As Variant, endScore As Variant) As Variant
    If Trim(beginScore & "") = "" Or Trim(endScore & "") = "" Then
        ComputeSdatPct = Null
        Exit Function
    End If
    Dim b As Double: b = CDbl(beginScore)
    Dim e As Double: e = CDbl(endScore)
    If b = 0 Then
        ComputeSdatPct = Null
        Exit Function
    End If
    ComputeSdatPct = modUtils.RoundHalfUp((b - e) * 1000 / b, 0) / 10
End Function

Public Sub SaveNotes(patientId As Long, notes As String)
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("_data_patients")
    Dim r As Long: r = modUtils.FindById(ws, patientId)
    If r = 0 Then Exit Sub
    If Len(notes) > 256 Then notes = Left(notes, 256)
    modUtils.SetVal ws, r, "notes", notes
    modUtils.AutoSave
End Sub

' ── WorkQueue sheet UI entry points ────────────────────────────────────────
Public Sub UI_OpenAddPatient()
    AddEditPatientForm.PatientId = 0
    AddEditPatientForm.Show
    RefreshWorkQueue
End Sub

Public Sub UI_ViewSelectedPatient()
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("WorkQueue")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("tblWorkQueue")
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "No patients to view.", vbExclamation
        Exit Sub
    End If
    If Intersect(Application.ActiveCell, tbl.DataBodyRange) Is Nothing Then
        MsgBox "Select a patient row first.", vbExclamation
        Exit Sub
    End If
    Dim rowOffset As Long: rowOffset = Application.ActiveCell.Row - tbl.DataBodyRange.Row + 1
    Dim pid As Long: pid = tbl.DataBodyRange(rowOffset, 1).Value
    PatientDetailForm.PatientId = pid
    PatientDetailForm.Show
    RefreshWorkQueue
End Sub

Public Sub UI_ExportWorkQueue()
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("WorkQueue")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("tblWorkQueue")
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "No data to export.", vbExclamation
        Exit Sub
    End If
    Dim lastRow As Long: lastRow = 4 + tbl.ListRows.Count
    modExport.ExportRangeToXlsx ws.Range(ws.Cells(4, 2), ws.Cells(lastRow, 11)), "WorkQueue.xlsx"
End Sub

' ISO "yyyy-mm-dd" -> a real Date (shown as mm/dd/yyyy by the column's number
' format, and sorted chronologically); "" if blank/invalid.
Private Function IsoToDateValue(iso As String) As Variant
    If Not modUtils.IsValidIsoDate(iso) Or iso = "" Then
        IsoToDateValue = ""
    Else
        IsoToDateValue = DateSerial(CInt(Left(iso, 4)), CInt(Mid(iso, 6, 2)), CInt(Mid(iso, 9, 2)))
    End If
End Function

' "yyyy-mm-dd hh:mm" -> Date+time value (or just the date if the time isn't HH:MM).
Private Function NextApptValue(text As String) As Variant
    If Len(text) < 10 Then
        NextApptValue = ""
        Exit Function
    End If
    Dim v As Variant: v = IsoToDateValue(Left(text, 10))
    If VarType(v) <> vbDate Then
        NextApptValue = ""
        Exit Function
    End If
    Dim t As String: t = Trim(Mid(text, 11))
    If t <> "" And modUtils.IsValidTime(t) Then v = v + TimeValue(t)
    NextApptValue = v
End Function

Public Sub RefreshWorkQueue()
    Dim ws As Worksheet: Set ws = modUtils.DataSheet("WorkQueue")

    Application.EnableEvents = False

    ws.Range("P2:P50").ClearContents
    Dim sources As Collection: Set sources = modUtils.ActiveLovValues("referral_source")
    Dim i As Long
    For i = 1 To sources.Count
        ws.Cells(1 + i, 16).Value = sources(i) ' column P = 16
    Next i

    Dim statusLabel As String: statusLabel = CStr(ws.Range("C1").Value)
    Dim statusFilter As String
    If statusLabel = "" Or statusLabel = "All" Then
        statusFilter = ""
    Else
        statusFilter = StatusInternal(statusLabel)
    End If

    Dim srcLabel As String: srcLabel = CStr(ws.Range("E1").Value)
    Dim referralSourceId As Long
    If srcLabel = "" Or srcLabel = "All" Then
        referralSourceId = 0
    Else
        referralSourceId = modUtils.LovIdForValue("referral_source", srcLabel)
    End If

    Dim searchText As String: searchText = CStr(ws.Range("G1").Value)
    Dim results As Collection: Set results = ListForWorkQueue(statusFilter, referralSourceId, searchText)

    Dim tbl As ListObject: Set tbl = ws.ListObjects("tblWorkQueue")
    modUtils.UnprotectForRefresh ws
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If

    If results.Count > 0 Then
        Dim r As Long: r = 1
        Dim d As Object
        For Each d In results
            tbl.ListRows.Add
            tbl.DataBodyRange(r, 1).Value = d("id")
            tbl.DataBodyRange(r, 2).Value = d("last_name")
            tbl.DataBodyRange(r, 3).Value = d("first_name")
            tbl.DataBodyRange(r, 4).Value = d("mrn")
            tbl.DataBodyRange(r, 5).Value = IsoToDateValue(CStr(d("date_of_referral")))
            tbl.DataBodyRange(r, 6).Value = d("referral_source")
            tbl.DataBodyRange(r, 7).Value = d("language")
            tbl.DataBodyRange(r, 8).Value = NextApptValue(CStr(d("next_appointment")))
            tbl.DataBodyRange(r, 9).Value = d("status_label")
            tbl.DataBodyRange(r, 10).Value = IsoToDateValue(CStr(d("followup_date")))
            tbl.DataBodyRange(r, 11).Value = d("followup_reminder")
            r = r + 1
        Next d
    End If
    modUtils.ReprotectAfterRefresh ws

    Dim counts As Object: Set counts = CountsByStatus()
    ws.Range("B2").Value = "Total Active: " & counts("total_active") & _
        "   Ready to Schedule: " & counts("ready_to_schedule") & _
        "   Scheduled: " & counts("scheduled") & _
        "   Active: " & counts("active") & _
        "   Active-Ongoing: " & counts("active_ongoing") & _
        "   Incomplete: " & counts("incomplete") & _
        "   Follow-ups due: " & FollowupsDueCount() & _
        "   Completed: " & counts("completed") & _
        "   Dropped: " & counts("dropped") & _
        "   On Hold: " & counts("on_hold")

    Application.EnableEvents = True
End Sub
