Sub GetVBAstat()

    Dim wb As Workbook, n%, count%, module As CodeModule, name$, lines&, totalLines&, counter&, added$, the_type, _
        totalMdl&, mdlLines&, totalFrm&, frmLines&, totalCls&, clsLines&

    On Error GoTo errHandler

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

                name = .name: name = name + added: name = left(name, 20)
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
        
            Debug.Print name, lines, the_type

        End If
        
    Next

    Debug.Print String(40, "-")
    Debug.Print "Modules: "; totalMdl
    Debug.Print "Modules' number of lines: "; mdlLines

    Debug.Print String(40, "-")
    Debug.Print "Classes: "; totalCls
    Debug.Print "Classes' number of lines: "; clsLines

    Debug.Print String(40, "-")
    Debug.Print "Forms:   "; totalFrm
    Debug.Print "Forms'   number of lines: "; frmLines

    Debug.Print String(40, "-")
    Debug.Print "Total mdls/cls/frms: "; counter
    Debug.Print "Total number of lines:    "; totalLines

    On Error GoTo 0

    Exit Sub

errHandler:

    Debug.Print err.number & " " & err.description
    err.Clear
    Resume Next

End Sub
