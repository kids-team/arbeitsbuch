Sub generateYear()
    If Sheets.Count > 2 Then
        MsgBox "Bitte löschen sie zunächst alle Kalenderwochen bis auf KW 1"
        Exit Sub
    End If
    If Sheets.Count = 1 Then
        MsgBox "KW 1 nicht gefunden. Möchten Sie eine Datei auswählen, die KW 1 enthält?"
        Exit Sub
    End If
    Dim x As Integer
    For x = 1 To 52
        ActiveWorkbook.Sheets(x + 1).Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        ActiveWorkbook.ActiveSheet.Name = "KW " & x + 1
        ActiveWorkbook.ActiveSheet.Range("P1").Value = x + 1
        ActiveWorkbook.ActiveSheet.Range("V27").Value = "='KW " & x & "'!V32"
    Next
    
    ActiveWorkbook.Save

End Sub


Sub analyzeData()

    Dim max As Integer
    max = Sheets.Count - 1
    For i = 1 To max
        ActiveWorkbook.Sheets(1).Range("V" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("V35").Value
        ActiveWorkbook.Sheets(1).Range("W" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("C37").Value
        ActiveWorkbook.Sheets(1).Range("X" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("F37").Value
        ActiveWorkbook.Sheets(1).Range("Y" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("I37").Value
        ActiveWorkbook.Sheets(1).Range("Z" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("L37").Value
        ActiveWorkbook.Sheets(1).Range("AA" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("O37").Value
        ActiveWorkbook.Sheets(1).Range("AB" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("R37").Value
        ActiveWorkbook.Sheets(1).Range("AC" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("U37").Value
        ActiveWorkbook.Sheets(1).Range("AE" & 5 + i).Value = ActiveWorkbook.Sheets(i + 1).Range("V32").Value
    Next
    

End Sub
