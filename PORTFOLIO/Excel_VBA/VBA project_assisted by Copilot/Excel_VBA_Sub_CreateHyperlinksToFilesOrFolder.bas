Sub CreateHyperlinksToFilesOrFoldersAllColumns()
    'THIS CODE IS MOSTLY VIBE CODED BY USING COPILOT. OF COURSE THE END RESULT IS SLIGHTLY POLISHED MY ME BECAUSE THE AI RESULT USUALLY INCLUDES UNNECESSARY CODE LINES AND HARD CODED LINES
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String, userinput As String
    

    Call turnoffstuff
    Set ws = shSisällys '
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' DYNAMIC ONEDRIVE ROOT AS THERE MIGHT BE USERS WHO HAS EITHER "TYÖPÖYTÄ" OR "DESKTOP"
    
    userinput = InputBox("Työpöytä or Desktop:", "Write either Työpöytä or Desktop")
    
    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/" & userinput) + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Kalibrointitodistukset\"
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Kalibrointitodistus.pdf"
            folderName = positionValue & "_Kalibrointitodistukset"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "F"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "F"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "F").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation

If userinput = "Työpöytä" Then

Call CreateHyperlinksToFilesOrFolders2
Call CreateHyperlinksToFilesOrFolders3
Call CreateHyperlinksToFilesOrFolders4

Else
Call CreateHyperlinksToFilesOrFolders2_DESKTOP
Call CreateHyperlinksToFilesOrFolders3_DESKTOP
Call CreateHyperlinksToFilesOrFolders4_DESKTOP
End If

Call turnonstuff



End Sub

'FROM THIS LINE ALL THE WAY DOWN THE CODES ARE JUST SOMETHING YOU CALL IN ORDER TO RUN THE MAIN CODE ABOVE---------------------------------------

Sub CreateHyperlinksToFilesOrFolders2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String
    
    Set ws = shSisällys
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
 
    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/Työpöytä") + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Laitetiedot ja asetukset\"
    
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Laitetiedot ja asetukset.pdf"
            folderName = positionValue & "_Laitetiedot ja asetukset"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "G"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "G"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "G").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation
End Sub

Sub CreateHyperlinksToFilesOrFolders2_DESKTOP()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String
    
    Set ws = shSisällys
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
  
    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/Desktop") + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Laitetiedot ja asetukset\"
    
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Laitetiedot ja asetukset.pdf"
            folderName = positionValue & "_Laitetiedot ja asetukset"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "G"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "G"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "G").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation
End Sub
Sub CreateHyperlinksToFilesOrFolders3()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String
    
    Set ws = shSisällys ' <-- Ensure shData is your sheet object
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    

    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/Työpöytä") + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Erityisdokumentaatiot\"
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Erityisdokumentaatio.pdf"
            folderName = positionValue & "_Erityisdokumentaatiot"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "H"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "H"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "H").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation
End Sub
Sub CreateHyperlinksToFilesOrFolders3_DESKTOP()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String
    
    Set ws = shSisällys ' <-- Ensure shData is your sheet object
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
  
    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/Desktop") + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Erityisdokumentaatiot\"
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Erityisdokumentaatio.pdf"
            folderName = positionValue & "_Erityisdokumentaatiot"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "H"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "H"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "H").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation
End Sub
Sub CreateHyperlinksToFilesOrFolders4()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String
    
    Set ws = shSisällys ' <-- Ensure shData is your sheet object
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    

    
    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/Työpöytä") + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Laitekohtaiset todistukset\"
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Laitekohtainen todistus.pdf"
            folderName = positionValue & "_Laitekohtaiset todistukset"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "I"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "I"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "I").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation
End Sub
Sub CreateHyperlinksToFilesOrFolders4_DESKTOP()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim positionValue As String
    Dim fileName As String, folderName As String
    Dim localRoot As String, tailPath As String, urlpath As String
    Dim pdfPath As String, folderPath As String
    
    Set ws = shSisällys ' <-- Ensure shData is your sheet object
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    
    urlpath = ThisWorkbook.path
    tailPath = Right(urlpath, Len(urlpath) - InStr(urlpath, "/Desktop") + 1)
    localRoot = Environ("OneDrive") & tailPath & "\Laitekohtaiset todistukset\"
    
    Dim i As Long
    For i = 2 To lastRow
        positionValue = Trim(ws.Cells(i, "A").Value)
        If positionValue <> "" Then
            ' Build file and folder names
            fileName = positionValue & "_Laitekohtainen todistus.pdf"
            folderName = positionValue & "_Laitekohtaiset todistukset"
            
            ' Full paths
            pdfPath = localRoot & fileName
            folderPath = localRoot & folderName & "\"
            'Debug.Print pdfPath
            'Debug.Print folderPath
            ' Check existence
            If Dir(pdfPath) <> "" Then
                ' PDF exists ? hyperlink to PDF
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, "I"), _
                    Address:=pdfPath, _
                    TextToDisplay:=fileName

            ElseIf Len(Dir(folderPath, vbDirectory)) > 0 Then
            ' Folder exists ? hyperlink to folder
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, "I"), _
            Address:=folderPath, _
            TextToDisplay:=folderName

            Else
                ' Neither exists ? clear cell
                ws.Cells(i, "I").Value = ""
            End If
        End If
    
    Next i
    
    'MsgBox "Hyperlinks created successfully!", vbInformation
End Sub
Sub turnoffstuff()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False


End Sub

Sub turnonstuff()

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub
