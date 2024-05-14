Sub Update_Current()
'make variables - AAAAAAAA
Dim index As Integer
Dim description As String
Dim account As String
Dim basis As Single
Dim useful_life As Single
Dim service_date As Date

Dim lookup_month As String
Dim lookup_parse As String
Dim enter_date As Date

Dim detail_row As Integer
Dim je_lookup As Integer

Dim account_check As String
Dim basis_check As Single
Dim useful_life_check As Single
Dim service_date_check As Date

Dim overwrite_check As Boolean
Dim overwrite As Integer

Dim monthly_dep As Single
Dim start_column As String
Dim end_column As String

Dim net_value As Single
Dim adjustment_loop As Integer
Dim adjustment_column As String
Dim adjustment_amount As Single
Dim check_amount As Single

'data is validated at the sheet level, so not done here

'check index exists since blank is allowed for functional reasons
index = Range("C4").Value
If index = 0 Then
    MsgBox ("Please select an index number to change values.")
    Exit Sub
End If

'put all form data into variables
description = Range("C6").Value
account = Range("J6").Value
basis = Range("C8").Value
service_date = Range("H8").Value
useful_life = Range("K8").Value
lookup_month = Range("J4").Value

'parse lookup_month into enter_date
lookup_parse = lookup_month & "-01"
enter_date = CDate(lookup_parse)

'select proper row to edit in Details sheet
detail_row = ThisWorkbook.Worksheets(9).Range("J6").Value

'see if popup for adjustment needed
'first check JE hist against insert month
je_lookup = ThisWorkbook.Worksheets(9).Range("J4")
If (Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(7).Range("C" & je_lookup & ":L14")) <> 0) Then
    'go through data change checks and set bool
    overwrite_check = False
    
    account_check = ThisWorkbook.Worksheets(6).Range("B" & detail_row).Value
    basis_check = ThisWorkbook.Worksheets(6).Range("F" & detail_row).Value
    useful_life_check = ThisWorkbook.Worksheets(6).Range("G" & detail_row).Value
    service_date_check = ThisWorkbook.Worksheets(6).Range("E" & detail_row).Value
    
    If account_check <> account Then
        overwrite_check = True
    End If
    If basis_check <> basis Then
        overwrite_check = True
    End If
    If useful_life_check <> useful_life Then
        overwrite_check = True
    End If
    If service_date_check <> service_date Then
        overwrite_check = True
    End If

    'warn if needed
    If overwrite_check Then
        overwrite = MsgBox("Adjusting this item will affect the journal entry in an already calculated period. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Journal Entry Overwrite")
        If overwrite = vbNo Then
            Exit Sub
        End If
    End If
End If

'calc monthly dep
monthly_dep = basis / (useful_life * 12)
monthly_dep = WorksheetFunction.Round(monthly_dep, 2)

'put into Details sheet, including month alloc - AAAAAAAA
ThisWorkbook.Worksheets(6).Range("A" & detail_row).Value = index
ThisWorkbook.Worksheets(6).Range("B" & detail_row).Value = account
ThisWorkbook.Worksheets(6).Range("C" & detail_row).Value = account 'lookup class from account
ThisWorkbook.Worksheets(6).Range("D" & detail_row).Value = description
ThisWorkbook.Worksheets(6).Range("E" & detail_row).Value = service_date
ThisWorkbook.Worksheets(6).Range("F" & detail_row).Value = basis
ThisWorkbook.Worksheets(6).Range("G" & detail_row).Value = useful_life
ThisWorkbook.Worksheets(6).Range("H" & detail_row).Value = "=G" & detail_row & " * 12"
ThisWorkbook.Worksheets(6).Range("I" & detail_row).Value = 0
ThisWorkbook.Worksheets(6).Range("J" & detail_row).Value = "=F" & detail_row & " - I" & detail_row
ThisWorkbook.Worksheets(6).Range("K" & detail_row).Value = monthly_dep
ThisWorkbook.Worksheets(6).Range("L" & detail_row).Value = "=SUM(P" & detail_row & ":AA" & detail_row & ")"
ThisWorkbook.Worksheets(6).Range("M" & detail_row).Value = "=I" & detail_row & " + L" & detail_row
ThisWorkbook.Worksheets(6).Range("N" & detail_row).Value = "=ROUND(F" & detail_row & " + M" & detail_row & ",2)"

'enter month alloc
start_column = ThisWorkbook.Worksheets(9).Range("J3").Value
end_column = ThisWorkbook.Worksheets(9).Range("B14").Value
ThisWorkbook.Worksheets(6).Range(start_column & detail_row & ":" & end_column & detail_row).Value = monthly_dep

'fix extra depreciations past net of zero - checks at $2 positive because of cumulative rounding errors
net_value = ThisWorkbook.Worksheets(6).Range("N" & detail_row).Value
If net_value < 2 Then
    'remove excess and edit last to get to true zero
    For adjustment_loop = 1 To 12
        adjustment_column = ThisWorkbook.Worksheets(9).Range("B" & (15 - adjustment_loop)).Value
        'add column to check, see if needs to be eliminated entirely
        check_amount = ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_row).Value
        If check_amount <= ((net_value * -1) + 2) Then
            'eliminate
            ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_row) = 0
            net_value = net_value + check_amount
        End If
        If (net_value <= 2) And (net_value >= -2) Then
        'adjust if less than 2 +/-
            adjustment_amount = net_value + ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_row).Value
            ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_row) = adjustment_amount
        End If
    Next adjustment_loop
End If

'clear index and set all back to formulas
Range("C4").Value = ""
Range("C6").Value = "=VLOOKUP($C$4,Detail!$A$1:$N$87,4)"
Range("J6").Value = "=VLOOKUP($C$4,Detail!$A$1:$N$87,2)"
Range("C8").Value = "=VLOOKUP($C$4,Detail!$A$1:$N$87,6)"
Range("H8").Value = "=VLOOKUP($C$4,Detail!$A$1:$N$87,5)"
Range("K8").Value = "=VLOOKUP($C$4,Detail!$A$1:$N$87,7)"

End Sub
Sub Remove_Current()
'make variables
Dim index As Integer

Dim lookup_month As String
Dim lookup_parse As String
Dim enter_date As Date

Dim detail_row As Integer
Dim je_lookup As Integer

Dim overwrite As Integer

Dim monthly_dep As Single
Dim start_column As String
Dim end_column As String


'data is validated at the sheet level, so not done here

'check index exists since blank is allowed for functional reasons
index = Range("C4").Value
If index = 0 Then
    MsgBox ("Please select an index number to remove values.")
    Exit Sub
End If

'select proper row to remove in Details sheet
detail_row = ThisWorkbook.Worksheets(9).Range("K6").Value

'put form data into variables
lookup_month = Range("J4").Value

'parse lookup_month into enter_date
lookup_parse = lookup_month & "-01"
enter_date = CDate(lookup_parse)

'check JE hist against insert month and warn if needed
je_lookup = ThisWorkbook.Worksheets(9).Range("K4")
If Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(7).Range("C" & je_lookup & ":L14")) <> 0 Then
    overwrite = MsgBox("Removing this item will affect the journal entry in an already calculated period. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Journal Entry Overwrite")
    If overwrite = vbNo Then
        Exit Sub
    End If
End If

'delete month alloc
start_column = ThisWorkbook.Worksheets(9).Range("K3").Value
end_column = ThisWorkbook.Worksheets(9).Range("B14").Value
ThisWorkbook.Worksheets(6).Range(start_column & detail_row & ":" & end_column & detail_row).Value = ""

'set total net to zero
ThisWorkbook.Worksheets(6).Range("M" & detail_row).Value = ThisWorkbook.Worksheets(6).Range("J" & detail_row).Value

'clear form to show complete
Range("C4").Value = ""

End Sub
Sub Add_New()
'make variables
Dim index As Integer
Dim description As String
Dim account As String
Dim basis As Single
Dim useful_life As Single

Dim lookup_month As String
Dim lookup_parse As String
Dim enter_date As Date

Dim detail_max As Integer
Dim je_lookup As Integer

Dim overwrite As Integer

Dim monthly_dep As Single
Dim start_column As String
Dim end_column As String

Dim net_value As Single
Dim adjustment_loop As Integer
Dim adjustment_column As String
Dim adjustment_amount As Single
Dim check_amount As Single


'set index as the next integer
detail_max = ThisWorkbook.Worksheets(6).Range("A" & ThisWorkbook.Worksheets(6).Rows.Count).End(xlUp).Row
index = WorksheetFunction.Max(ThisWorkbook.Worksheets(6).Range("A2:A" & detail_max)) + 1

'data is validated at the sheet level, so not done here

'put all form data into variables
description = Range("C6").Value
account = Range("J6").Value
basis = Range("C8").Value
useful_life = Range("K8").Value
lookup_month = Range("J4").Value

'parse lookup_month into enter_date
lookup_parse = lookup_month & "-01"
enter_date = CDate(lookup_parse)

'check JE hist against insert month and warn if needed
je_lookup = ThisWorkbook.Worksheets(9).Range("I4")
If Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(7).Range("C" & je_lookup & ":L14")) <> 0 Then
    overwrite = MsgBox("Adding this item will affect the journal entry in an already calculated period. Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Journal Entry Overwrite")
    If overwrite = vbNo Then
        Exit Sub
    End If
End If

'insert row in Details to preserve formulas
detail_max = detail_max + 1
ThisWorkbook.Worksheets(6).Range("A" & detail_max).EntireRow.Insert

'calc monthly dep
monthly_dep = basis / (useful_life * 12)
monthly_dep = WorksheetFunction.Round(monthly_dep, 2)

'put into Details sheet, including month alloc
ThisWorkbook.Worksheets(6).Range("A" & detail_max).Value = index
ThisWorkbook.Worksheets(6).Range("B" & detail_max).Value = account
ThisWorkbook.Worksheets(6).Range("C" & detail_max).Value = "=VLOOKUP(" & Chr(34) & account & Chr(34) & ",B2:C" & detail_max - 1 & ",2,FALSE)" 'lookup class from account
ThisWorkbook.Worksheets(6).Range("C" & detail_max).Value = ThisWorkbook.Worksheets(6).Range("C" & detail_max).Value 'removes formula and changes to set value
ThisWorkbook.Worksheets(6).Range("D" & detail_max).Value = description
ThisWorkbook.Worksheets(6).Range("E" & detail_max).Value = enter_date
ThisWorkbook.Worksheets(6).Range("F" & detail_max).Value = basis
ThisWorkbook.Worksheets(6).Range("G" & detail_max).Value = useful_life
ThisWorkbook.Worksheets(6).Range("H" & detail_max).Value = "=G" & detail_max & " * 12"
ThisWorkbook.Worksheets(6).Range("I" & detail_max).Value = 0
ThisWorkbook.Worksheets(6).Range("J" & detail_max).Value = "=F" & detail_max & " - I" & detail_max
ThisWorkbook.Worksheets(6).Range("K" & detail_max).Value = monthly_dep
ThisWorkbook.Worksheets(6).Range("L" & detail_max).Value = "=SUM(P" & detail_max & ":AA" & detail_max & ")"
ThisWorkbook.Worksheets(6).Range("M" & detail_max).Value = "=I" & detail_max & " + L" & detail_max
ThisWorkbook.Worksheets(6).Range("N" & detail_max).Value = "=ROUND(F" & detail_max & " - M" & detail_max & ",2)"

'enter month alloc
start_column = ThisWorkbook.Worksheets(9).Range("I3").Value
end_column = ThisWorkbook.Worksheets(9).Range("B14").Value
ThisWorkbook.Worksheets(6).Range(start_column & detail_max & ":" & end_column & detail_max).Value = monthly_dep

'fix extra depreciations past net of zero - checks at $2 positive because of cumulative rounding errors
net_value = ThisWorkbook.Worksheets(6).Range("N" & detail_max).Value
If net_value < 2 Then
    'remove excess and edit last to get to true zero
    For adjustment_loop = 1 To 12
        adjustment_column = ThisWorkbook.Worksheets(9).Range("B" & (15 - adjustment_loop)).Value
        'add column to check, see if needs to be eliminated entirely
        check_amount = ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_max).Value
        If check_amount <= ((net_value * -1) + 2) Then
            'eliminate
            ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_max) = 0
            net_value = net_value + check_amount
        End If
        If (net_value <= 2) And (net_value >= -2) Then
        'adjust if less than 2 +/-
            adjustment_amount = net_value + ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_max).Value
            ThisWorkbook.Worksheets(6).Range(adjustment_column & detail_max) = adjustment_amount
        End If
    Next adjustment_loop
End If

'clear form to show complete
Range("C6").Value = ""
Range("J6").Value = ""
Range("C8").Value = ""
Range("K8").Value = ""

End Sub

Sub Lookup_Balance()
'make variable for each category and for month
Dim act_15100 As Single
Dim act_15199 As Single
Dim act_15200 As Single
Dim act_15299 As Single
Dim act_15300 As Single
Dim act_15399 As Single
Dim act_15400 As Single
Dim act_15499 As Single
Dim act_15500 As Single
Dim act_15599 As Single
Dim act_15600 As Single
Dim act_15699 As Single
Dim act_15700 As Single
Dim act_15710 As Single
Dim act_15719 As Single
Dim act_15799 As Single
Dim act_15800 As Single
Dim act_15900 As Single

Dim lookup_month As String
Dim today As Date

Dim lookup_parse As String
Dim lookup_date As Date

Dim detail_max As Integer
Dim detail_row As Integer

Dim detail_date As Date
Dim detail_basis As Single
Dim detail_dep As Single
Dim start_column As String
Dim end_column As String
Dim detail_act As String

Dim bs_row As Integer


'get current day and lookup month
lookup_month = Range("D3").Value
today = Date

'parse lookup_month into lookup_date for using in next
lookup_parse = lookup_month & "-01"
lookup_date = CDate(lookup_parse)
lookup_date = DateAdd("m", 1, lookup_date)

'set all accounts to start at zero
act_15100 = 0
act_15199 = 0
act_15200 = 0
act_15299 = 0
act_15300 = 0
act_15399 = 0
act_15400 = 0
act_15499 = 0
act_15500 = 0
act_15599 = 0
act_15600 = 0
act_15699 = 0
act_15700 = 0
act_15710 = 0
act_15719 = 0
act_15799 = 0
act_15800 = 0
act_15900 = 0


'go through detail sheet and add to appropriate categories
detail_max = ThisWorkbook.Worksheets(6).Range("A" & Rows.Count).End(xlUp).Row
start_column = ThisWorkbook.Worksheets(9).Range("G2").Value
end_column = ThisWorkbook.Worksheets(9).Range("G3").Value
For detail_row = 2 To detail_max
    detail_basis = 0
    detail_dep = 0
    'check against placed in service date
    detail_date = ThisWorkbook.Worksheets(6).Range("E" & detail_row).Value
    If detail_date < lookup_date Then
        'get data about basis
        detail_basis = detail_basis + ThisWorkbook.Worksheets(6).Range("F" & detail_row).Value
        detail_basis = WorksheetFunction.Round(detail_basis, 2)
        'get data about accrual
        'as of BoY
        detail_dep = detail_dep + ThisWorkbook.Worksheets(6).Range("I" & detail_row).Value
        'as of curr month - use months accts sheet formula for column lookup
        detail_dep = detail_dep + Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(6).Range(start_column & detail_row & ":" & end_column & detail_row))
        detail_dep = WorksheetFunction.Round(detail_dep, 2)
        'put into proper account
        detail_act = ThisWorkbook.Worksheets(6).Range("B" & detail_row).Value
        
        If detail_act = "15100/15199" Then
            act_15100 = act_15100 + detail_basis
            act_15199 = act_15199 - detail_dep
            act_15100 = WorksheetFunction.Round(act_15100, 2)
            act_15199 = WorksheetFunction.Round(act_15199, 2)
        End If
        If detail_act = "15200/15299" Then
            act_15200 = act_15200 + detail_basis
            act_15299 = act_15299 - detail_dep
            act_15200 = WorksheetFunction.Round(act_15200, 2)
            act_15299 = WorksheetFunction.Round(act_15299, 2)
        End If
        If detail_act = "15300/15399" Then
            act_15300 = act_15300 + detail_basis
            act_15399 = act_15399 - detail_dep
            act_15300 = WorksheetFunction.Round(act_15300, 2)
            act_15399 = WorksheetFunction.Round(act_15399, 2)
        End If
        If detail_act = "15400/15499" Then
            act_15400 = act_15400 + detail_basis
            act_15499 = act_15499 - detail_dep
            act_15400 = WorksheetFunction.Round(act_15400, 2)
            act_15499 = WorksheetFunction.Round(act_15499, 2)
        End If
        If detail_act = "15500/15599" Then
            act_15500 = act_15500 + detail_basis
            act_15599 = act_15599 - detail_dep
            act_15500 = WorksheetFunction.Round(act_15500, 2)
            act_15599 = WorksheetFunction.Round(act_15599, 2)
        End If
        If detail_act = "15600/15699" Then
            act_15600 = act_15600 + detail_basis
            act_15699 = act_15699 - detail_dep
            act_15600 = WorksheetFunction.Round(act_15600, 2)
            act_15699 = WorksheetFunction.Round(act_15699, 2)
        End If
        If detail_act = "15700/15799" Then
            act_15700 = act_15700 + detail_basis
            act_15799 = act_15799 - detail_dep
            act_15700 = WorksheetFunction.Round(act_15700, 2)
            act_15799 = WorksheetFunction.Round(act_15799, 2)
        End If
        If detail_act = "15710/15719" Then
            act_15710 = act_15710 + detail_basis
            act_15719 = act_15719 - detail_dep
            act_15710 = WorksheetFunction.Round(act_15710, 2)
            act_15719 = WorksheetFunction.Round(act_15719, 2)
        End If
        If detail_act = "15800" Then
            act_15800 = act_15800 + detail_basis
            act_15800 = WorksheetFunction.Round(act_15800, 2)
        End If
        If detail_act = "15900" Then
            act_15900 = act_15900 + detail_basis
            act_15900 = WorksheetFunction.Round(act_15900, 2)
        End If
        
    End If
Next detail_row

'round all accounts for cleanliness
act_15100 = WorksheetFunction.Round(act_15100, 2)
act_15199 = WorksheetFunction.Round(act_15199, 2)
act_15200 = WorksheetFunction.Round(act_15200, 2)
act_15299 = WorksheetFunction.Round(act_15299, 2)
act_15300 = WorksheetFunction.Round(act_15300, 2)
act_15399 = WorksheetFunction.Round(act_15399, 2)
act_15400 = WorksheetFunction.Round(act_15400, 2)
act_15499 = WorksheetFunction.Round(act_15499, 2)
act_15500 = WorksheetFunction.Round(act_15500, 2)
act_15599 = WorksheetFunction.Round(act_15599, 2)
act_15600 = WorksheetFunction.Round(act_15600, 2)
act_15699 = WorksheetFunction.Round(act_15699, 2)
act_15700 = WorksheetFunction.Round(act_15700, 2)
act_15710 = WorksheetFunction.Round(act_15710, 2)
act_15719 = WorksheetFunction.Round(act_15719, 2)
act_15799 = WorksheetFunction.Round(act_15799, 2)
act_15800 = WorksheetFunction.Round(act_15800, 2)
act_15900 = WorksheetFunction.Round(act_15900, 2)

'add to detail balance sheet - formula auto-updates from this on main sheet
bs_row = ThisWorkbook.Worksheets(9).Range("G4").Value
ThisWorkbook.Worksheets(8).Range("B" & bs_row).Value = today
ThisWorkbook.Worksheets(8).Range("C" & bs_row).Value = act_15100
ThisWorkbook.Worksheets(8).Range("D" & bs_row).Value = act_15199
ThisWorkbook.Worksheets(8).Range("E" & bs_row).Value = act_15200
ThisWorkbook.Worksheets(8).Range("F" & bs_row).Value = act_15299
ThisWorkbook.Worksheets(8).Range("G" & bs_row).Value = act_15300
ThisWorkbook.Worksheets(8).Range("H" & bs_row).Value = act_15399
ThisWorkbook.Worksheets(8).Range("I" & bs_row).Value = act_15400
ThisWorkbook.Worksheets(8).Range("J" & bs_row).Value = act_15499
ThisWorkbook.Worksheets(8).Range("K" & bs_row).Value = act_15500
ThisWorkbook.Worksheets(8).Range("L" & bs_row).Value = act_15599
ThisWorkbook.Worksheets(8).Range("M" & bs_row).Value = act_15600
ThisWorkbook.Worksheets(8).Range("N" & bs_row).Value = act_15699
ThisWorkbook.Worksheets(8).Range("O" & bs_row).Value = act_15700
ThisWorkbook.Worksheets(8).Range("P" & bs_row).Value = act_15710
ThisWorkbook.Worksheets(8).Range("Q" & bs_row).Value = act_15719
ThisWorkbook.Worksheets(8).Range("R" & bs_row).Value = act_15799
ThisWorkbook.Worksheets(8).Range("S" & bs_row).Value = act_15800
ThisWorkbook.Worksheets(8).Range("T" & bs_row).Value = act_15900

End Sub
Sub New_Year_Sheet()
'make variables
'construct new name
'update all dates
'delete the old data
'add beginning balance
'plot new months
'remove monthly dep amounts as needed
'popup saying complete

End Sub
Sub Calc_JE()
'make variable for each category and for month
Dim act_80200 As Single
Dim act_15299 As Single
Dim act_80300 As Single
Dim act_15399 As Single
Dim act_80400 As Single
Dim act_15499 As Single
Dim act_80500 As Single
Dim act_15599 As Single
Dim act_80600 As Single
Dim act_15699 As Single

Dim lookup_month As String
Dim today As Date

Dim lookup_parse As String
Dim lookup_date As Date

Dim detail_max As Integer
Dim detail_row As Integer

Dim detail_date As Date
Dim detail_dep As Single
Dim start_column As String
Dim end_column As String
Dim detail_act As String

Dim je_lookup As Integer

Dim act_80200_check As Single
Dim act_15299_check As Single
Dim act_80300_check As Single
Dim act_15399_check As Single
Dim act_80400_check As Single
Dim act_15499_check As Single
Dim act_80500_check As Single
Dim act_15599_check As Single
Dim act_80600_check As Single
Dim act_15699_check As Single

Dim popup_needed As Boolean
Dim overwrite As Integer


'get current day and lookup month
lookup_month = Range("D3").Value
today = Date

'parse lookup_month into lookup_date for using in next
lookup_parse = lookup_month & "-01"
lookup_date = CDate(lookup_parse)
lookup_date = DateAdd("m", 1, lookup_date)

'go through detail sheet and calculate
detail_max = ThisWorkbook.Worksheets(6).Range("A" & Rows.Count).End(xlUp).Row
month_column = ThisWorkbook.Worksheets(9).Range("H3").Value
For detail_row = 2 To detail_max
    detail_dep = 0
    'check against placed in service date
    detail_date = ThisWorkbook.Worksheets(6).Range("E" & detail_row).Value
    If detail_date < lookup_date Then
        'get data about accrual
        'as of curr month - use months accts sheet formula for column lookup
        detail_dep = detail_dep + ThisWorkbook.Worksheets(6).Range(month_column & detail_row).Value
        detail_dep = WorksheetFunction.Round(detail_dep, 2)
        'put into proper account
        detail_act = ThisWorkbook.Worksheets(6).Range("B" & detail_row).Value
        
        If detail_act = "15200/15299" Then
            act_15299 = act_15299 + detail_dep
        End If
        If detail_act = "15300/15399" Then
            act_15399 = act_15399 + detail_dep
        End If
        If detail_act = "15400/15499" Then
            act_15499 = act_15499 + detail_dep
        End If
        If detail_act = "15500/15599" Then
            act_15599 = act_15599 + detail_dep
        End If
        If detail_act = "15600/15699" Then
            act_15699 = act_15699 + detail_dep
        End If
        
    End If
Next detail_row

'round all accounts for cleanliness
act_15299 = WorksheetFunction.Round(act_15299, 2)
act_15399 = WorksheetFunction.Round(act_15399, 2)
act_15499 = WorksheetFunction.Round(act_15499, 2)
act_15599 = WorksheetFunction.Round(act_15599, 2)
act_15699 = WorksheetFunction.Round(act_15699, 2)

'match expense accounts to balance sheet accounts
act_80200 = act_15299
act_80300 = act_15399
act_80400 = act_15499
act_80500 = act_15599
act_80600 = act_15699

'check if exists already
je_lookup = ThisWorkbook.Worksheets(9).Range("H4")
If ThisWorkbook.Worksheets(7).Range("B" & je_lookup).Value <> 0 Then
    'gather for check
    act_80200_check = ThisWorkbook.Worksheets(7).Range("C" & je_lookup).Value
    act_15299_check = ThisWorkbook.Worksheets(7).Range("D" & je_lookup).Value
    act_80300_check = ThisWorkbook.Worksheets(7).Range("E" & je_lookup).Value
    act_15399_check = ThisWorkbook.Worksheets(7).Range("F" & je_lookup).Value
    act_80400_check = ThisWorkbook.Worksheets(7).Range("G" & je_lookup).Value
    act_15499_check = ThisWorkbook.Worksheets(7).Range("H" & je_lookup).Value
    act_80500_check = ThisWorkbook.Worksheets(7).Range("I" & je_lookup).Value
    act_15599_check = ThisWorkbook.Worksheets(7).Range("J" & je_lookup).Value
    act_80600_check = ThisWorkbook.Worksheets(7).Range("K" & je_lookup).Value
    act_15699_check = ThisWorkbook.Worksheets(7).Range("L" & je_lookup).Value
    
    'check if different - toggle popup
    popup_needed = False
    If act_80200_check <> act_80200 Then
        popup_needed = True
    End If
    If act_15299_check <> act_15299 Then
        popup_needed = True
    End If
    If act_80300_check <> act_80300 Then
        popup_needed = True
    End If
    If act_15399_check <> act_15399 Then
        popup_needed = True
    End If
    If act_80400_check <> act_80400 Then
        popup_needed = True
    End If
    If act_15499_check <> act_15499 Then
        popup_needed = True
    End If
    If act_80500_check <> act_80500 Then
        popup_needed = True
    End If
    If act_15599_check <> act_15599 Then
        popup_needed = True
    End If
    If act_80600_check <> act_80600 Then
        popup_needed = True
    End If
    If act_15699_check <> act_15699 Then
        popup_needed = True
    End If
    
    'popup warning
    If popup_needed Then
        overwrite = MsgBox("There is already a journal entry for this month. Do you want to overwrite with new values?", vbQuestion + vbYesNo + vbDefaultButton2, "Journal Entry Overwrite")
        If overwrite = vbNo Then
            Exit Sub
        End If
    End If
End If
 
'add to detail balance sheet - formula auto-updates from this on main sheet
bs_row = detail_basis + ThisWorkbook.Worksheets(9).Range("H4").Value
ThisWorkbook.Worksheets(7).Range("B" & je_lookup).Value = today
ThisWorkbook.Worksheets(7).Range("C" & je_lookup).Value = act_80200
ThisWorkbook.Worksheets(7).Range("D" & je_lookup).Value = act_15299
ThisWorkbook.Worksheets(7).Range("E" & je_lookup).Value = act_80300
ThisWorkbook.Worksheets(7).Range("F" & je_lookup).Value = act_15399
ThisWorkbook.Worksheets(7).Range("G" & je_lookup).Value = act_80400
ThisWorkbook.Worksheets(7).Range("H" & je_lookup).Value = act_15499
ThisWorkbook.Worksheets(7).Range("I" & je_lookup).Value = act_80500
ThisWorkbook.Worksheets(7).Range("J" & je_lookup).Value = act_15599
ThisWorkbook.Worksheets(7).Range("K" & je_lookup).Value = act_80600
ThisWorkbook.Worksheets(7).Range("L" & je_lookup).Value = act_15699

End Sub
