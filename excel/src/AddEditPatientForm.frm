VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddEditPatientForm 
   Caption         =   "Add / Edit Patient"
   ClientHeight    =   5832
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6588
   OleObjectBlob   =   "AddEditPatientForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddEditPatientForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PatientId As Long   ' 0 = new patient

Private mPhoneUpdating As Boolean
Private WithEvents mTxtDob As MSForms.TextBox   ' birth date row is added at runtime (see AddBirthDateRow)
Private mLblDob As MSForms.Label
Private mLblAge As MSForms.Label
Private mTxtFrom As MSForms.TextBox     ' referral rows are also added at runtime (see AddReferralRows)
Private mTxtReason As MSForms.TextBox
Private mTxtPcp As MSForms.TextBox
Private mLoaded As Boolean   ' Activate re-fires after the calendar picker closes; load once

Private Sub UserForm_Activate()
    If mLoaded Then Exit Sub
    mLoaded = True
    AddBirthDateRow
    AddReferralRows
    lblReferralDateHint.Caption = "(MM/DD/YYYY)"
    modUtils.PopulateLovCombo cboReferralSource, "referral_source", True
    modUtils.PopulateLovCombo cboReligion, "religion", True
    modUtils.PopulateLovCombo cboLanguage, "language", True, "English,Spanish"

    If PatientId > 0 Then
        Dim p As Object: Set p = modPatients.GetPatient(PatientId)
        txtMrn.Text = p("mrn") & ""
        txtLastName.Text = p("last_name") & ""
        txtFirstName.Text = p("first_name") & ""
        txtMiddleName.Text = p("middle_name") & ""
        txtPhone.Text = p("phone") & ""
        mTxtDob.Text = modUtils.FmtDate(p("date_of_birth") & "")
        mTxtFrom.Text = p("referral_from") & ""
        mTxtReason.Text = p("referral_reason") & ""
        mTxtPcp.Text = p("pcp_emailed") & ""
        txtReferralDate.Text = modUtils.FmtDate(p("date_of_referral") & "")
        cboReferralSource.value = p("referral_source") & ""
        cboReligion.value = p("religion") & ""
        cboLanguage.value = p("language") & ""
        Me.Caption = "Edit Patient"
    Else
        Me.Caption = "Add Patient"
        ' Defaults to today; still editable by typing or via the calendar button.
        txtReferralDate.Text = Format(Date, "mm/dd/yyyy")
    End If
End Sub

' Inserts a "Birth Date" row directly under Phone and a live age readout beside it.
' Built in code (like CalendarPickerForm's day grid) so the form's binary layout
' doesn't have to change: every control below the Phone row slides down one row.
Private Sub AddBirthDateRow()
    Const ROW_H As Single = 26
    Dim edge As Single: edge = txtPhone.Top + txtPhone.Height
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If c.Top >= edge Then c.Top = c.Top + ROW_H
    Next c
    Me.Height = Me.Height + ROW_H

    Set mLblDob = Me.Controls.Add("Forms.Label.1", "lblDateOfBirth", True)
    With mLblDob
        .Caption = "Birth Date:"
        .Left = lblPhone.Left: .Top = lblPhone.Top + ROW_H
        .Width = lblPhone.Width: .Height = lblPhone.Height
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With

    Set mTxtDob = Me.Controls.Add("Forms.TextBox.1", "txtDateOfBirth", True)
    With mTxtDob
        .Left = txtPhone.Left: .Top = txtPhone.Top + ROW_H
        .Width = 100: .Height = txtPhone.Height
        .MaxLength = 10
        .Font.Name = txtPhone.Font.Name: .Font.Size = txtPhone.Font.Size
        .TabIndex = txtPhone.TabIndex + 1
    End With

    Set mLblAge = Me.Controls.Add("Forms.Label.1", "lblAge", True)
    With mLblAge
        .Left = mTxtDob.Left + mTxtDob.Width + 8: .Top = mTxtDob.Top + 3
        .Width = txtPhone.Left + txtPhone.Width - .Left: .Height = lblPhone.Height
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
    UpdateAgeLabel
End Sub

' Adds Referral From, Referral Reason and "PCP emailed?" under Language; the
' Save/Cancel buttons and error label slide down to make room.
Private Sub AddReferralRows()
    Const GAP As Single = 6
    Const FROM_H As Single = 20
    Const REASON_H As Single = 44
    Const PCP_H As Single = 20
    Dim total As Single: total = FROM_H + REASON_H + PCP_H + 3 * GAP

    Dim edge As Single: edge = cboLanguage.Top + cboLanguage.Height
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If c.Top >= edge Then c.Top = c.Top + total
    Next c
    Me.Height = Me.Height + total

    Dim y As Single: y = edge + GAP
    AddRowLabel "lblReferralFrom", "Referral From:", y + 2
    Set mTxtFrom = AddRowTextBox("txtReferralFrom", y, FROM_H)
    mTxtFrom.MaxLength = 100

    y = y + FROM_H + GAP
    AddRowLabel "lblReferralReason", "Referral Reason:", y + 2
    Set mTxtReason = AddRowTextBox("txtReferralReason", y, REASON_H)
    mTxtReason.MultiLine = True
    mTxtReason.ScrollBars = fmScrollBarsVertical
    mTxtReason.MaxLength = 500

    y = y + REASON_H + GAP
    AddRowLabel "lblPcpEmailed", "PCP Emailed?", y + 2
    Set mTxtPcp = AddRowTextBox("txtPcpEmailed", y, PCP_H)
    mTxtPcp.MaxLength = 100
End Sub

Private Sub AddRowLabel(ctlName As String, caption As String, topPos As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .Caption = caption
        .Left = lblPhone.Left: .Top = topPos
        .Width = lblPhone.Width: .Height = lblPhone.Height
        .Font.Name = lblPhone.Font.Name: .Font.Size = lblPhone.Font.Size
    End With
End Sub

Private Function AddRowTextBox(ctlName As String, topPos As Single, h As Single) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    Set tb = Me.Controls.Add("Forms.TextBox.1", ctlName, True)
    With tb
        .Left = txtPhone.Left: .Top = topPos
        .Width = txtPhone.Width: .Height = h
        .Font.Name = txtPhone.Font.Name: .Font.Size = txtPhone.Font.Size
    End With
    Set AddRowTextBox = tb
End Function

Private Sub UpdateAgeLabel()
    Dim s As String: s = Trim(mTxtDob.Text)
    Dim iso As String: iso = modUtils.ParseDate(s)
    If s = "" Then
        mLblAge.Caption = "(MM/DD/YYYY)"
    ElseIf iso <> "" And modUtils.AgeFromIsoDate(iso) <> "" Then
        mLblAge.Caption = "Age " & modUtils.AgeFromIsoDate(iso)
    Else
        mLblAge.Caption = "Invalid date"
    End If
End Sub

Private Sub mTxtDob_Change()
    UpdateAgeLabel
End Sub

' Blocks any non-digit keystroke so the phone field can only ever contain 0-9.
Private Sub txtPhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        ' digit or backspace -- allow
    Else
        KeyAscii = 0
    End If
End Sub

' KeyPress can't catch pasted text -- strip any non-digits that get through that way.
Private Sub txtPhone_Change()
    If mPhoneUpdating Then Exit Sub
    Dim cleaned As String, i As Long, ch As String
    For i = 1 To Len(txtPhone.Text)
        ch = Mid(txtPhone.Text, i, 1)
        If ch >= "0" And ch <= "9" Then cleaned = cleaned & ch
    Next i
    If cleaned <> txtPhone.Text Then
        mPhoneUpdating = True
        Dim pos As Long: pos = txtPhone.SelStart
        txtPhone.Text = cleaned
        txtPhone.SelStart = IIf(pos > Len(cleaned), Len(cleaned), pos)
        mPhoneUpdating = False
    End If
End Sub

Private Sub btnPickDate_Click()
    CalendarPickerForm.InitialDate = modUtils.ParseDate(txtReferralDate.Text)
    CalendarPickerForm.SelectedDate = ""
    CalendarPickerForm.Show
    If CalendarPickerForm.SelectedDate <> "" Then
        txtReferralDate.Text = modUtils.FmtDate(CalendarPickerForm.SelectedDate)
    End If
    Unload CalendarPickerForm
End Sub

Private Sub btnSave_Click()
    If Trim(txtLastName.Text) = "" Or Trim(txtFirstName.Text) = "" Then
        ShowError "First and last name are required."
        Exit Sub
    End If
    If Trim(txtReferralDate.Text) = "" Then
        ShowError "Referral date is required."
        Exit Sub
    End If
    Dim referralIso As String: referralIso = modUtils.ParseDate(txtReferralDate.Text)
    If referralIso = "" Then
        ShowError "Referral date must be MM/DD/YYYY."
        Exit Sub
    End If

    Dim dob As String: dob = modUtils.ParseDate(mTxtDob.Text)   ' ISO, or "" if blank/invalid
    If Trim(mTxtDob.Text) <> "" And (dob = "" Or modUtils.AgeFromIsoDate(dob) = "") Then
        ShowError "Birth date must be a past date as MM/DD/YYYY."
        Exit Sub
    End If

    ' There's no status field on this form: a new patient's initial status date is
    ' the referral date (today if that's in the future). Correct it afterwards
    ' from the patient's status history (Edit Date) if needed.
    Dim statusIso As String: statusIso = referralIso
    If statusIso > modUtils.DateISO(Date) Then statusIso = modUtils.DateISO(Date)

    Dim referralSourceId As Long: referralSourceId = modUtils.LovIdForValue("referral_source", cboReferralSource.value)
    Dim religionId As Long: religionId = modUtils.LovIdForValue("religion", cboReligion.value)
    Dim languageId As Long: languageId = modUtils.LovIdForValue("language", cboLanguage.value)

    modPatients.Save PatientId, Trim(txtMrn.Text), Trim(txtLastName.Text), Trim(txtFirstName.Text), _
        Trim(txtMiddleName.Text), Trim(txtPhone.Text), referralIso, _
        referralSourceId, religionId, languageId, dob, _
        Trim(mTxtFrom.Text), Trim(mTxtReason.Text), Trim(mTxtPcp.Text), statusIso
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub ShowError(msg As String)
    lblError.Caption = msg
    lblError.Visible = True
End Sub


