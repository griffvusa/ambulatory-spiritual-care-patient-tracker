VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PatientDetailForm 
   Caption         =   "Patient Detail"
   ClientHeight    =   12228
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12576
   OleObjectBlob   =   "PatientDetailForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PatientDetailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public patientId As Long

Private mLblDobCaption As MSForms.Label   ' birth date row is added at runtime (see EnsureBirthDateRow)
Private mLblDob As MSForms.Label
Private mLblFrom As MSForms.Label         ' referral rows are also added at runtime (see EnsureReferralRows)
Private mLblReason As MSForms.Label
Private mLblPcp As MSForms.Label
Private mLblFollowup As MSForms.Label      ' follow-up date entry is added at runtime (see EnsureFollowupControls)
Private mTxtFollowup As MSForms.TextBox
Private WithEvents mBtnPickFollowup As MSForms.CommandButton
Private WithEvents mBtnPickStatusDate As MSForms.CommandButton

' Values as last loaded, so any Apply/Save button can tell what is pending.
Private mSnapStatusLabel As String
Private mSnapFollowup As String
Private mSnapSdatBeginScore As String
Private mSnapSdatBeginDate As String
Private mSnapSdatEndScore As String
Private mSnapSdatEndDate As String
Private mSnapNotes As String
Private WithEvents mBtnSaveFollowup As MSForms.CommandButton
Private mLblStatusDate As MSForms.Label    ' status (effective) date entry, also added at runtime
Private mTxtStatusDate As MSForms.TextBox
Private WithEvents mBtnEditHistory As MSForms.CommandButton   ' "Edit Date" for a status-history entry, added at runtime

Private Sub UserForm_Activate()
    LoadPatient
End Sub

' Adds a "Born" row under Phone once per form instance, in code so the form's
' binary layout doesn't have to change. Left-panel and bottom-panel controls
' slide down one row; the appointments panel on the right stays put.
Private Sub EnsureBirthDateRow()
    If Not mLblDob Is Nothing Then Exit Sub
    Const ROW_H As Single = 20
    Dim edge As Single: edge = lblPhone.Top + lblPhone.Height
    Dim bottomPanelTop As Single: bottomPanelTop = txtSdatBeginScore.Top - 25
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If c.Top >= edge Then
            If Not (c.Left >= lstAppointments.Left - 5 And c.Top < bottomPanelTop) Then c.Top = c.Top + ROW_H
        End If
    Next c
    Me.Height = Me.Height + ROW_H

    Set mLblDobCaption = Me.Controls.Add("Forms.Label.1", "lblDobCaption", True)
    With mLblDobCaption
        .Caption = "Born:"
        .Left = lblReferralDateCaption.Left: .Top = lblPhone.Top + ROW_H
        .Width = lblReferralDateCaption.Width: .Height = lblPhone.Height
        .Font.Name = lblReferralDateCaption.Font.Name: .Font.Size = lblReferralDateCaption.Font.Size
    End With
    Set mLblDob = Me.Controls.Add("Forms.Label.1", "lblDob", True)
    With mLblDob
        .Left = lblPhone.Left: .Top = lblPhone.Top + ROW_H
        .Width = 180: .Height = lblPhone.Height
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
End Sub

' Adds Referral From / Reason / PCP Emailed rows under Language once per form
' instance, in code so the form's binary layout doesn't have to change.
Private Sub EnsureReferralRows()
    If Not mLblFrom Is Nothing Then Exit Sub
    Const ROW_H As Single = 20
    Const REASON_H As Single = 34
    Dim total As Single: total = ROW_H + REASON_H + ROW_H
    Dim edge As Single: edge = lblLanguage.Top + lblLanguage.Height
    Dim bottomPanelTop As Single: bottomPanelTop = txtSdatBeginScore.Top - 25
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If c.Top >= edge Then
            If Not (c.Left >= lstAppointments.Left - 5 And c.Top < bottomPanelTop) Then c.Top = c.Top + total
        End If
    Next c
    Me.Height = Me.Height + total

    Dim y As Single: y = lblLanguage.Top + ROW_H
    AddCaption "lblFromCaption", "From:", y
    Set mLblFrom = AddValue("lblFrom", y, lblPhone.Height)
    y = y + ROW_H
    AddCaption "lblReasonCaption", "Reason:", y
    Set mLblReason = AddValue("lblReason", y, REASON_H - 4)
    mLblReason.WordWrap = True
    y = y + REASON_H
    AddCaption "lblPcpCaption", "PCP emailed:", y
    Set mLblPcp = AddValue("lblPcp", y, lblPhone.Height)
End Sub

Private Sub AddCaption(ctlName As String, caption As String, topPos As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .Caption = caption
        .Left = lblReferralDateCaption.Left: .Top = topPos
        .Width = lblReferralDateCaption.Width: .Height = lblPhone.Height
        .Font.Name = lblReferralDateCaption.Font.Name: .Font.Size = lblReferralDateCaption.Font.Size
    End With
End Sub

Private Function AddValue(ctlName As String, topPos As Single, h As Single) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .Left = lblLanguage.Left: .Top = topPos
        .Width = 200: .Height = h
        .Font.Name = lblLanguage.Font.Name: .Font.Size = lblLanguage.Font.Size
    End With
    Set AddValue = lbl
End Function

Private Sub LoadPatient()
    Dim p As Object: Set p = modPatients.GetPatient(patientId)

    lblPatientName.Caption = Trim(p("first_name") & " " & p("middle_name") & " " & p("last_name"))
    lblMrn.Caption = p("mrn") & ""
    lblPhone.Caption = p("phone") & ""
    EnsureBirthDateRow
    Dim dob As String: dob = p("date_of_birth") & ""
    If dob = "" Then
        mLblDob.Caption = ""
    ElseIf modUtils.AgeFromIsoDate(dob) <> "" Then
        mLblDob.Caption = modUtils.FmtDate(dob) & "  (Age " & modUtils.AgeFromIsoDate(dob) & ")"
    Else
        mLblDob.Caption = modUtils.FmtDate(dob)
    End If
    lblReferralDate.Caption = modUtils.FmtDate(p("date_of_referral") & "")
    lblReferralSource.Caption = p("referral_source") & ""
    lblReligion.Caption = p("religion") & ""
    lblLanguage.Caption = p("language") & ""
    EnsureReferralRows
    EnsureFollowupControls
    mLblFrom.Caption = p("referral_from") & ""
    mLblReason.Caption = p("referral_reason") & ""
    mLblPcp.Caption = p("pcp_emailed") & ""

    Dim currentStatus As String: currentStatus = p("current_status") & ""
    lblStatusBadge.Caption = modPatients.statusLabel(currentStatus)
    mTxtFollowup.Text = modUtils.FmtDate(p("followup_date") & "")
    lblStatusBadge.BackColor = StatusBadgeColor(currentStatus)

    ' ColumnWidths was fixed at 7 columns at design time (date/time/type/
    ' consultant/status/last-appt/notes); a form's own code can safely widen
    ' its own control's runtime ColumnCount to smuggle in a leading 0pt ID
    ' column, which is not the same restricted operation as external
    ' automation touching a Designer's own layout properties.
    lstAppointments.Clear
    lstAppointments.ColumnCount = 8
    lstAppointments.ColumnWidths = "0 pt;55 pt;40 pt;55 pt;70 pt;55 pt;25 pt;90 pt"
    Dim ap As Object
    For Each ap In modAppointments.GetForPatient(patientId)
        lstAppointments.AddItem ap("id")
        Dim apptIdx As Long: apptIdx = lstAppointments.ListCount - 1
        lstAppointments.List(apptIdx, 1) = modUtils.FmtDate(ap("date") & "")
        lstAppointments.List(apptIdx, 2) = ap("time")
        lstAppointments.List(apptIdx, 3) = ap("type")
        lstAppointments.List(apptIdx, 4) = ap("consultant_name")
        lstAppointments.List(apptIdx, 5) = ap("status")
        lstAppointments.List(apptIdx, 6) = IIf(ap("is_last_appointment") = 1, "Yes", "")
        lstAppointments.List(apptIdx, 7) = ap("notes")
    Next ap

    cboChangeStatus.Clear
    Dim lbl As Variant
    For Each lbl In modPatients.AllStatusLabelsExceptDeleted()
        cboChangeStatus.AddItem CStr(lbl)
    Next lbl
    If currentStatus <> "deleted" Then
        cboChangeStatus.value = modPatients.statusLabel(currentStatus)
    End If

    lstStatusHistory.Clear
    lstStatusHistory.ColumnCount = 4                                  ' 4th column: hidden history id
    lstStatusHistory.ColumnWidths = "80 pt;90 pt;80 pt;0 pt"
    Dim h As Object
    For Each h In modPatients.GetStatusHistory(patientId)
        lstStatusHistory.AddItem h("status")
        lstStatusHistory.List(lstStatusHistory.ListCount - 1, 1) = h("changed_by_name")
        lstStatusHistory.List(lstStatusHistory.ListCount - 1, 2) = modUtils.FmtDate(h("changed_at") & "")
        lstStatusHistory.List(lstStatusHistory.ListCount - 1, 3) = h("id")
    Next h

    txtSdatBeginScore.Text = p("sdat_begin_score") & ""
    txtSdatBeginDate.Text = modUtils.FmtDate(p("sdat_begin_date") & "")
    txtSdatEndScore.Text = p("sdat_end_score") & ""
    txtSdatEndDate.Text = modUtils.FmtDate(p("sdat_end_date") & "")
    If Trim(p("sdat_pct_improvement") & "") = "" Then
        lblSdatPctImprovement.Caption = ""
    Else
        lblSdatPctImprovement.Caption = "Improvement: " & p("sdat_pct_improvement") & "%"
    End If
    txtNotes.Text = p("notes") & ""

    Dim isDeleted As Boolean: isDeleted = (currentStatus = "deleted")
    btnDeletePatient.Visible = Not isDeleted
    btnRestorePatient.Visible = isDeleted

    mSnapStatusLabel = IIf(currentStatus <> "deleted", modPatients.statusLabel(currentStatus), "")
    mSnapFollowup = Trim(mTxtFollowup.Text)
    mSnapSdatBeginScore = Trim(txtSdatBeginScore.Text)
    mSnapSdatBeginDate = Trim(txtSdatBeginDate.Text)
    mSnapSdatEndScore = Trim(txtSdatEndScore.Text)
    mSnapSdatEndDate = Trim(txtSdatEndDate.Text)
    mSnapNotes = txtNotes.Text

    ' The appointments list snaps to whole rows once it's filled and can end up
    ' ~7pt taller than designed, running under the Add/Edit Appointment
    ' buttons; keep the buttons just below wherever it actually ends.
    Dim apptBtnTop As Single: apptBtnTop = lstAppointments.Top + lstAppointments.Height + 4
    btnAddAppointment.Top = apptBtnTop
    btnEditAppointment.Top = apptBtnTop
End Sub

' BackColor uses the same BGR-packed OLE_COLOR format as VBA's RGB() function,
' matching the Status Workflow colors from the design spec.
Private Function StatusBadgeColor(currentStatus As String) As Long
    Select Case LCase(currentStatus)
        Case "ready_to_schedule": StatusBadgeColor = RGB(&HC6, &HEF, &HCE)
        Case "scheduled": StatusBadgeColor = RGB(&HFF, &HEB, &H9C)
        Case "active": StatusBadgeColor = RGB(&HB4, &HE5, &HA2)
        Case "active_ongoing": StatusBadgeColor = RGB(&H8F, &HD9, &HC9)
        Case "incomplete": StatusBadgeColor = RGB(&HF4, &HB1, &H83)
        Case "completed": StatusBadgeColor = RGB(&H9D, &HC3, &HE6)
        Case "dropped": StatusBadgeColor = RGB(&HFF, &HC7, &HCE)
        Case "on_hold": StatusBadgeColor = RGB(&HD9, &HD9, &HD9)
        Case "deleted": StatusBadgeColor = RGB(&HBF, &HBF, &HBF)
        Case Else: StatusBadgeColor = RGB(255, 255, 255)
    End Select
End Function

Private Sub btnApplyStatus_Click()
    SaveAllPending True
End Sub

' One save for the whole screen. Apply and every Save button on this form call
' this, so an entry in any box -- status, follow-up date, SDAT or notes -- is
' saved whichever of them is pressed, whether you filled in one field or
' several. Nobody has to press Apply AND Save. Everything is validated first;
' if anything is invalid nothing is saved. A status is recorded when a
' different status is picked, or when the status date box has been changed
' (to back-date the status). fromApply only controls whether the user is told
' when nothing at all was changed.
Private Sub SaveAllPending(fromApply As Boolean)
    Dim today As String: today = modUtils.DateISO(Date)

    ' ── status change ──
    Dim statusDateText As String: statusDateText = Trim(mTxtStatusDate.Text)
    Dim statusDateEdited As Boolean
    statusDateEdited = (statusDateText <> "" And statusDateText <> Format(Date, "mm/dd/yyyy"))
    Dim doStatus As Boolean
    doStatus = (Trim(cboChangeStatus.value) <> "") And _
               (cboChangeStatus.value <> mSnapStatusLabel Or statusDateEdited)
    Dim newStatus As String, statusDate As String
    If doStatus Then
        newStatus = modPatients.StatusInternal(cboChangeStatus.value)
        statusDate = modUtils.ParseDate(mTxtStatusDate.Text)   ' ISO or ""
        If Trim(mTxtStatusDate.Text) <> "" Then
            If statusDate = "" Then
                MsgBox "Enter the status date as MM/DD/YYYY, or leave it blank to use today.", _
                    vbExclamation, "Status date"
                mTxtStatusDate.SetFocus
                Exit Sub
            End If
            If statusDate > today Then
                MsgBox "The status date can't be in the future.", vbExclamation, "Status date"
                mTxtStatusDate.SetFocus
                Exit Sub
            End If
            If statusDate = today Then statusDate = ""   ' same as the default
        End If
    End If

    ' ── follow-up date ──
    Dim doFollowup As Boolean: doFollowup = (Trim(mTxtFollowup.Text) <> mSnapFollowup)
    Dim followupIso As String
    If doFollowup Then
        followupIso = modUtils.ParseDate(mTxtFollowup.Text)
        If Trim(mTxtFollowup.Text) <> "" And followupIso = "" Then
            MsgBox "Enter the follow-up date as MM/DD/YYYY.", vbExclamation, "Follow-up date"
            mTxtFollowup.SetFocus
            Exit Sub
        End If
    End If

    ' ── SDAT ──
    Dim doSdat As Boolean
    doSdat = (Trim(txtSdatBeginScore.Text) <> mSnapSdatBeginScore) Or _
             (Trim(txtSdatBeginDate.Text) <> mSnapSdatBeginDate) Or _
             (Trim(txtSdatEndScore.Text) <> mSnapSdatEndScore) Or _
             (Trim(txtSdatEndDate.Text) <> mSnapSdatEndDate)
    Dim beginScore As Variant, endScore As Variant
    If doSdat Then
        beginScore = ParsedScore(txtSdatBeginScore.Text)
        endScore = ParsedScore(txtSdatEndScore.Text)
        If beginScore = "invalid" Or endScore = "invalid" Then
            MsgBox "SDAT scores must be a whole number between 0 and 40, or blank.", vbExclamation
            Exit Sub
        End If
        If Not modUtils.IsValidDateInput(txtSdatBeginDate.Text) Or Not modUtils.IsValidDateInput(txtSdatEndDate.Text) Then
            MsgBox "SDAT dates must be MM/DD/YYYY, or blank.", vbExclamation
            Exit Sub
        End If
    End If

    ' ── notes ──
    Dim doNotes As Boolean: doNotes = (txtNotes.Text <> mSnapNotes)
    If doNotes And Len(txtNotes.Text) > 256 Then
        MsgBox "Notes must be 256 characters or fewer.", vbExclamation
        Exit Sub
    End If

    ' ── everything valid: save what changed ──
    If Not (doStatus Or doFollowup Or doSdat Or doNotes) Then
        If fromApply Then MsgBox "Nothing has been changed.", vbInformation, "Patient"
        Exit Sub
    End If
    If doStatus Then
        modPatients.ChangeStatus patientId, newStatus, statusDate
        mTxtStatusDate.Text = Format(Date, "mm/dd/yyyy")   ' back to the default for the next change
    End If
    If doFollowup Then modPatients.SaveFollowupDate patientId, followupIso
    If doSdat Then
        modPatients.SaveSdat patientId, beginScore, modUtils.ParseDate(txtSdatBeginDate.Text), _
            endScore, modUtils.ParseDate(txtSdatEndDate.Text)
    End If
    If doNotes Then modPatients.SaveNotes patientId, txtNotes.Text
    LoadPatient
End Sub

' Status-change date entry and the patient's follow-up date (with a calendar
' button and its own Save), in two rows under the status picker. Built in code
' so the form's binary layout doesn't change: everything below the picker on
' the left and bottom panels slides down two rows.
Private Sub EnsureFollowupControls()
    If Not mTxtFollowup Is Nothing Then Exit Sub
    Const ROW_H As Single = 22
    Const SHIFT As Single = 2 * ROW_H
    Dim edge As Single: edge = cboChangeStatus.Top + cboChangeStatus.Height
    Dim bottomPanelTop As Single: bottomPanelTop = txtSdatBeginScore.Top - 25
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If c.Top >= edge Then
            If Not (c.Left >= lstAppointments.Left - 5 And c.Top < bottomPanelTop) Then c.Top = c.Top + SHIFT
        End If
    Next c
    Me.Height = Me.Height + SHIFT

    Dim y As Single: y = edge + 4
    Set mLblStatusDate = AddStatusLabel("lblStatusDate", "Status date:", y + 2)
    Set mTxtStatusDate = AddStatusText("txtStatusDate", y)
    mTxtStatusDate.Text = Format(Date, "mm/dd/yyyy")   ' defaults to today; type or pick another date to back-date
    mTxtStatusDate.ControlTipText = "Defaults to today. Type MM/DD/YYYY or use the calendar button to back-date this status change."
    Set mBtnPickStatusDate = Me.Controls.Add("Forms.CommandButton.1", "btnPickStatusDate", True)
    With mBtnPickStatusDate
        .Caption = "..."
        .Left = mTxtStatusDate.Left + mTxtStatusDate.Width + 2: .Top = y
        .Width = 22: .Height = 18
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With

    y = y + ROW_H
    Set mLblFollowup = AddStatusLabel("lblFollowupDate", "Follow-up:", y + 2)
    Set mTxtFollowup = AddStatusText("txtFollowupDate", y)
    mTxtFollowup.ControlTipText = "Follow-up date (MM/DD/YYYY). Clear it and Save to remove the follow-up."

    Set mBtnPickFollowup = Me.Controls.Add("Forms.CommandButton.1", "btnPickFollowup", True)
    With mBtnPickFollowup
        .Caption = "..."
        .Left = mTxtFollowup.Left + mTxtFollowup.Width + 2: .Top = y
        .Width = 22: .Height = 18
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
    Set mBtnSaveFollowup = Me.Controls.Add("Forms.CommandButton.1", "btnSaveFollowup", True)
    With mBtnSaveFollowup
        .Caption = "Save"
        .Left = mBtnPickFollowup.Left + mBtnPickFollowup.Width + 4: .Top = y
        .Width = 40: .Height = 18
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With

    Set mBtnEditHistory = Me.Controls.Add("Forms.CommandButton.1", "btnEditHistoryDate", True)
    With mBtnEditHistory
        .Caption = "Edit Date..."
        .Left = lstStatusHistory.Left + lstStatusHistory.Width - 80
        .Top = lblStatusHistoryHeader.Top - 3
        .Width = 80: .Height = 18
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
End Sub

Private Sub mBtnEditHistory_Click()
    EditSelectedHistoryDate
End Sub

Private Sub mBtnPickStatusDate_Click()
    CalendarPickerForm.InitialDate = modUtils.ParseDate(mTxtStatusDate.Text)
    CalendarPickerForm.SelectedDate = ""
    CalendarPickerForm.Show
    If CalendarPickerForm.SelectedDate <> "" Then
        mTxtStatusDate.Text = modUtils.FmtDate(CalendarPickerForm.SelectedDate)
    End If
    Unload CalendarPickerForm
End Sub

Private Sub mBtnPickFollowup_Click()
    CalendarPickerForm.InitialDate = modUtils.ParseDate(mTxtFollowup.Text)
    CalendarPickerForm.SelectedDate = ""
    CalendarPickerForm.Show
    If CalendarPickerForm.SelectedDate <> "" Then
        mTxtFollowup.Text = modUtils.FmtDate(CalendarPickerForm.SelectedDate)
    End If
    Unload CalendarPickerForm
End Sub

' Saves the follow-up date (a blank box clears it) along with anything else pending.
Private Sub mBtnSaveFollowup_Click()
    SaveAllPending False
End Sub

Private Sub lstStatusHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EditSelectedHistoryDate
End Sub

' Lets the user correct the date of the selected status-history entry.
Private Sub EditSelectedHistoryDate()
    If lstStatusHistory.ListIndex < 0 Then
        MsgBox "Select a status history entry first.", vbInformation, "Edit Status Date"
        Exit Sub
    End If
    Dim idx As Long: idx = lstStatusHistory.ListIndex
    Dim current As String: current = Left(CStr(lstStatusHistory.List(idx, 2)), 10)
    Dim answer As String
    answer = InputBox("Date for the " & Chr(34) & lstStatusHistory.List(idx, 0) & Chr(34) & " status (MM/DD/YYYY):", _
                      "Edit Status Date", current)
    answer = Trim(answer)
    If answer = "" Or answer = current Then Exit Sub   ' cancelled or unchanged

    Dim answerIso As String: answerIso = modUtils.ParseDate(answer)
    If answerIso = "" Then
        MsgBox "That isn't a valid date. Use MM/DD/YYYY.", vbExclamation, "Edit Status Date"
        Exit Sub
    End If
    If answerIso > modUtils.DateISO(Date) Then
        MsgBox "The status date can't be in the future.", vbExclamation, "Edit Status Date"
        Exit Sub
    End If
    If modPatients.UpdateStatusHistoryDate(CLng(lstStatusHistory.List(idx, 3)), answerIso) Then
        LoadPatient
    Else
        MsgBox "Couldn't update that entry.", vbExclamation, "Edit Status Date"
    End If
End Sub

Private Function AddStatusLabel(ctlName As String, caption As String, topPos As Single) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .Caption = caption
        .Left = cboChangeStatus.Left: .Top = topPos
        .Width = 58: .Height = 16
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
    Set AddStatusLabel = lbl
End Function

Private Function AddStatusText(ctlName As String, topPos As Single) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    Set tb = Me.Controls.Add("Forms.TextBox.1", ctlName, True)
    With tb
        .Left = cboChangeStatus.Left + 60: .Top = topPos
        .Width = 72: .Height = 18
        .MaxLength = 10
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
    Set AddStatusText = tb
End Function

Private Sub btnEditPatient_Click()
    AddEditPatientForm.patientId = patientId
    AddEditPatientForm.Show
    LoadPatient
End Sub

Private Sub btnAddAppointment_Click()
    AddEditAppointmentForm.patientId = patientId
    AddEditAppointmentForm.appointmentId = 0
    AddEditAppointmentForm.Show
    LoadPatient
End Sub

Private Sub btnEditAppointment_Click()
    If lstAppointments.ListIndex = -1 Then
        MsgBox "Select an appointment first.", vbExclamation
        Exit Sub
    End If
    AddEditAppointmentForm.appointmentId = CLng(lstAppointments.Column(0, lstAppointments.ListIndex))
    AddEditAppointmentForm.patientId = patientId
    AddEditAppointmentForm.Show
    LoadPatient
End Sub

Private Sub btnPickSdatBeginDate_Click()
    CalendarPickerForm.InitialDate = modUtils.ParseDate(txtSdatBeginDate.Text)
    CalendarPickerForm.SelectedDate = ""
    CalendarPickerForm.Show
    If CalendarPickerForm.SelectedDate <> "" Then
        txtSdatBeginDate.Text = modUtils.FmtDate(CalendarPickerForm.SelectedDate)
    End If
    Unload CalendarPickerForm
End Sub

Private Sub btnPickSdatEndDate_Click()
    CalendarPickerForm.InitialDate = modUtils.ParseDate(txtSdatEndDate.Text)
    CalendarPickerForm.SelectedDate = ""
    CalendarPickerForm.Show
    If CalendarPickerForm.SelectedDate <> "" Then
        txtSdatEndDate.Text = modUtils.FmtDate(CalendarPickerForm.SelectedDate)
    End If
    Unload CalendarPickerForm
End Sub

Private Sub btnSaveSdat_Click()
    SaveAllPending False
End Sub

' Returns a Long 0-40, Empty for a blank string, or the string "invalid".
Private Function ParsedScore(s As String) As Variant
    If Trim(s) = "" Then
        ParsedScore = Empty
        Exit Function
    End If
    If Not IsNumeric(s) Then
        ParsedScore = "invalid"
        Exit Function
    End If
    ' IsNumeric accepts values too large for CLng (e.g. "99999999999" or
    ' scientific notation) -- CLng(s) below would raise an unhandled Overflow
    ' error rather than the intended "invalid" result, so guard it explicitly
    ' instead of relying on VBA's Or (which does not short-circuit and would
    ' evaluate CLng(s) unconditionally).
    On Error Resume Next
    Dim longVal As Long
    longVal = CLng(s)
    Dim overflowed As Boolean: overflowed = (Err.Number <> 0)
    On Error GoTo 0
    If overflowed Or longVal <> CDbl(s) Or longVal < 0 Or longVal > 40 Then
        ParsedScore = "invalid"
    Else
        ParsedScore = longVal
    End If
End Function

Private Sub btnSaveNotes_Click()
    SaveAllPending False
End Sub

Private Sub btnDeletePatient_Click()
    If MsgBox("Delete this patient? This can be undone from the Deleted filter.", _
              vbQuestion + vbYesNo, "Confirm Delete") <> vbYes Then Exit Sub
    modPatients.Delete patientId
    LoadPatient
End Sub

Private Sub btnRestorePatient_Click()
    modPatients.Restore patientId
    LoadPatient
End Sub


