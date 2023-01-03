Sub Chemformula()
For Each cl In Selection.Cells
ntr = InStr(1, UCase(Range(cl.Address)), "H2O2") + _
    InStr(1, UCase(Range(cl.Address)), "H2SO4") + _
    InStr(1, UCase(Range(cl.Address)), "NAOH") + _
    InStr(1, UCase(Range(cl.Address)), "CLO2") + _
    InStr(1, UCase(Range(cl.Address)), "O2")
If ntr > 0 Then
    If InStr(1, UCase(Range(cl.Address)), "H2O2") > 0 Then
        var2 = InStr(1, Range(cl.Address), "h2o2", 1)
        cl.Characters(Start:=var2, Length:=1).Caption = UCase(cl.Characters(Start:=var2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 1, Length:=1).Font.Subscript = True
        cl.Characters(Start:=var2 + 2, Length:=1).Caption = UCase(cl.Characters(Start:=var2 + 2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 3, Length:=1).Font.Subscript = True
    End If
    
    If InStr(1, UCase(Range(cl.Address)), "H2SO4") > 0 Then
        var2 = InStr(1, Range(cl.Address), "h2so4", 1)
        cl.Characters(Start:=var2, Length:=1).Caption = UCase(cl.Characters(Start:=var2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 1, Length:=1).Font.Subscript = True
        cl.Characters(Start:=var2 + 2, Length:=1).Caption = UCase(cl.Characters(Start:=var2 + 2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 3, Length:=1).Caption = UCase(cl.Characters(Start:=var2 + 3, Length:=1).Caption)
        cl.Characters(Start:=var2 + 4, Length:=1).Font.Subscript = True
    End If
    
    If InStr(1, UCase(Range(cl.Address)), "NAOH") > 0 Then
        var2 = InStr(1, Range(cl.Address), "naoh", 1)
        cl.Characters(Start:=var2, Length:=1).Caption = UCase(cl.Characters(Start:=var2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 1, Length:=1).Caption = LCase(cl.Characters(Start:=var2 + 1, Length:=1).Caption)
        cl.Characters(Start:=var2 + 2, Length:=2).Caption = UCase(cl.Characters(Start:=var2 + 2, Length:=2).Caption)
    End If
    
    If InStr(1, UCase(Range(cl.Address)), "CLO2") > 0 Then
        var2 = InStr(1, Range(cl.Address), "clo2", 1)
        cl.Characters(Start:=var2, Length:=1).Caption = UCase(cl.Characters(Start:=var2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 1, Length:=1).Caption = LCase(cl.Characters(Start:=var2 + 1, Length:=1).Caption)
        cl.Characters(Start:=var2 + 2, Length:=1).Caption = UCase(cl.Characters(Start:=var2 + 2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 3, Length:=1).Font.Subscript = True
    End If
    
    If InStr(1, UCase(Range(cl.Address)), "O2") > 0 Then
        var2 = InStr(1, Range(cl.Address), "o2", 1)
        cl.Characters(Start:=var2, Length:=1).Caption = UCase(cl.Characters(Start:=var2, Length:=1).Caption)
        cl.Characters(Start:=var2 + 1, Length:=1).Font.Subscript = True
    End If
End If
Next
End Sub

