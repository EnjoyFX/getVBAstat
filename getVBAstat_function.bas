Function GetVBAstat() As String

    Dim wb As Workbook, n%, count%, module As CodeModule, name$, lines&, totalLines&, counter&, added$, the_type, _
        totalMdl&, mdlLines&, totalFrm&, frmLines&, totalCls&, clsLines&, txt$, the_line$

    On Error GoTo errHandler

    the_line = String(40, "-")

    txt = "VBA Project Statistic" & vbNewLine & vbNewLine & "By module/form/class name:" & vbNewLine & the_line & vbNewLine & _
          "Name " & vbTab & " Number of lines" & vbNewLine & the_line & vbNewLine

    Set wb = ThisWorkbook

    count = wb.VBProject.VBComponents.count

    totalLines = 0: counter = 0
    totalMdl = 0: mdlLines = 0
    totalCls = 0: clsLines = 0
    totalFrm = 0: frmLines = 0

    added = String(20, " ")

    For n = 1 To count

        the_type = wb.VBProject.VBComponents.Item(n).Type    ' 1 - unit /2 - class /3 - form

        Set module = wb.VBProject.VBComponents.Item(n).CodeModule

        If the_type < 4 Then

            With module

                name = .name: name = name + added: name = Left(name, 20)
                lines = .CountOfLines

                totalLines = totalLines + lines
                counter = counter + 1


                If the_type = 1 Then
                    totalMdl = totalMdl + 1
                    mdlLines = mdlLines + lines
                ElseIf the_type = 2 Then
                    totalCls = totalCls + 1
                    clsLines = clsLines + lines
                ElseIf the_type = 3 Then
                    totalFrm = totalFrm + 1
                    frmLines = frmLines + lines
                End If


            End With

            txt = txt & name & lines & vbNewLine

        End If


    Next


    txt = txt & vbNewLine & "Total by type:" & vbNewLine
    txt = txt & the_line & vbNewLine
    txt = txt & "Modules: " & totalMdl & vbNewLine
    txt = txt & "Modules' number of lines: " & mdlLines & vbNewLine

    txt = txt & the_line & vbNewLine
    txt = txt & "Classes: " & totalCls & vbNewLine
    txt = txt & "Classes' number of lines: " & clsLines & vbNewLine

    txt = txt & the_line & vbNewLine
    txt = txt & "Forms:   " & totalFrm & vbNewLine
    txt = txt & "Forms'   number of lines: " & frmLines & vbNewLine

    txt = txt & the_line & vbNewLine
    txt = txt & "Total modules/classes/forms: " & counter & vbNewLine
    txt = txt & "Total number of lines:    " & totalLines & vbNewLine

    Debug.Print txt

    GetVBAstat = txt

    On Error GoTo 0

    Exit Function

errHandler:

    Debug.Print Err.Number & " " & Err.Description
    Err.Clear
    Resume Next

End Function
