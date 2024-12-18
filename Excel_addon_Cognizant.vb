// Custom UI

/*
Sub LoadCustRibbon()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "Excel.officeUI"


ribbonXML = "<mso:customUI      xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "      <mso:tab id='additionalTols' label='Additional Options' insertBeforeQ='mso:TabFormat'>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='reportGroup' label='Elements From Build Samples' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport1' label='MC to Same Sheet' imageMso='TableSelect'  onAction='PERSONAL.XLSB!Elements_Same_Sheet.getAllMCToSameSheet'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport2' label='MD to Same Sheet' imageMso='TableSelect'  onAction='PERSONAL.XLSB!Elements_Same_Sheet.getAllMDToSameSheet'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport3' label='Elements to Same Sheet' imageMso='TableSelect'  onAction='PERSONAL.XLSB!Elements_Same_Sheet.getAllElementsToSameSheet'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport4' label='MC to Seperate Sheet' imageMso='ShapesDuplicate'  onAction='PERSONAL.XLSB!Elements_Different_Sheet.getMDToExcelDifferentSheet'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport5' label='MD to Seperate Sheet' imageMso='ShapesDuplicate'  onAction='PERSONAL.XLSB!Elements_Different_Sheet.getMCToExcelDifferentSheet'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport6' label='Elements to Seperate Sheet' imageMso='ShapesDuplicate'  onAction='PERSONAL.XLSB!Elements_Different_Sheet.getElementsToExcelDifferentSheet'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='reportGroup2' label='QI Spec SHeet' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='runReport7' label='MC Block Names from QI Spec Sheet' imageMso='SlidesPerPage9Slides'  onAction='PERSONAL.XLSB!MCFromQI.getMCNameQISpec'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "</mso:customUI>"

ribbonXML = Replace(ribbonXML, """", "")

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub

Sub ClearCustRibbon()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "Excel.officeUI"

ribbonXML = "<mso:customUI           xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
"<mso:ribbon></mso:ribbon></mso:customUI>"

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub




*/



// Elements_Different_Sheet

/*
Sub getMDToExcelDifferentSheet()
    
    On Error GoTo error_handle

    Dim fso As Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsFile As Scripting.File
    
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wDocTextBox As String
    Dim nNumber As Integer
    
    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    Dim wordDocData As String
       

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    

    Set fso = CreateObject("Scripting.filesystemobject")
    Set fsFolder = fso.GetFolder(diaFolder.SelectedItems(1))
    
    
    
    startCell = 2
    
    For Each fsFile In fsFolder.Files
        
        Set wApp = CreateObject("word.Application")
        Set wDoc = wApp.Documents.Open(fsFile.path)
    
        Debug.Print fsFile.path
        
        ' Body Content
        wordDocData = wDoc.Content.Text
        
        
        ' Text Boxes
        With wDoc
            For nNumber = .Shapes.Count To 1 Step -1
                If .Shapes(nNumber).Type = msoTextBox Then
                    wDocTextBox = wDocTextBox & .Shapes(nNumber).TextFrame.TextRange.Text
                End If
            Next
        End With
        wordDocData = wordDocData & wDocTextBox
        wDocTextBox = ""
        
        
        ' Header
        With wDoc.Sections(1).Footers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        ' Footer
        With wDoc.Sections(1).Headers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        
        
        wDoc.Close
        wApp.Quit
        
        new_sheet_name = Left(fsFile.Name, (InStrRev(fsFile.Name, ".", -1, vbTextCompare) - 1))
        
        ' Create Sheet
        
        Sheets.Add After:=Sheets(Sheets.Count)
        
        new_sheet_name = Left(new_sheet_name, 20)
        
        Sheets(ActiveSheet.Name).Name = new_sheet_name
        
        
        mc_list = getMD(wordDocData)
        
        For i = 0 To UBound(mc_list) - 1
            Sheets(new_sheet_name).Cells(2 + i, 1) = mc_list(i, 0)
        Next i
        
    
        ' Trimming
        count_rows_final = WorksheetFunction.CountA(Columns(1))
        
        Range("B2").Select
        ActiveCell.FormulaR1C1 = _
            "=TRIM(SUBSTITUTE(SUBSTITUTE(RC[-1],""<"",""""),"">"",""""))"
        Range("B2").Select
        Selection.AutoFill Destination:=Range("B2:B" & count_rows_final + 1)
        
        
        
        
        ' Unique Values
        
        Range("C2").Select
        ActiveCell.Formula2R1C1 = "=UNIQUE(C[-1])"
        
        ' Resize
        Columns("B:C").Select
        Columns("B:C").EntireColumn.AutoFit
        
        Range("D3").Select
    
        
    Next fsFile
    
Done:
    Exit Sub
    
    
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"
    


End Sub


Sub getMCToExcelDifferentSheet()
    
    On Error GoTo error_handle

    Dim fso As Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsFile As Scripting.File
    
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wDocTextBox As String
    Dim nNumber As Integer
    
    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    Dim wordDocData As String
       

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    

    Set fso = CreateObject("Scripting.filesystemobject")
    Set fsFolder = fso.GetFolder(diaFolder.SelectedItems(1))
    
    
    
    startCell = 2
    
    For Each fsFile In fsFolder.Files
        
        Set wApp = CreateObject("word.Application")
        Set wDoc = wApp.Documents.Open(fsFile.path)
    
        Debug.Print fsFile.path
        
        ' Body Content
        wordDocData = wDoc.Content.Text
        
        
        ' Text Boxes
        With wDoc
            For nNumber = .Shapes.Count To 1 Step -1
                If .Shapes(nNumber).Type = msoTextBox Then
                    wDocTextBox = wDocTextBox & .Shapes(nNumber).TextFrame.TextRange.Text
                End If
            Next
        End With
        wordDocData = wordDocData & wDocTextBox
        wDocTextBox = ""
        
        
        ' Header
        With wDoc.Sections(1).Footers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        ' Footer
        With wDoc.Sections(1).Headers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        
        
        wDoc.Close
        wApp.Quit
        
        new_sheet_name = Left(fsFile.Name, (InStrRev(fsFile.Name, ".", -1, vbTextCompare) - 1))
        
        ' Create Sheet
        
        Sheets.Add After:=Sheets(Sheets.Count)
        
        new_sheet_name = Left(new_sheet_name, 20)
        
        Sheets(ActiveSheet.Name).Name = new_sheet_name
        
        
        mc_list = getMC(wordDocData)
        
        For i = 0 To UBound(mc_list) - 1
            Sheets(new_sheet_name).Cells(2 + i, 1) = mc_list(i, 0)
        Next i
        
    
        ' Trimming
        count_rows_final = WorksheetFunction.CountA(Columns(1))
        
        Range("B2").Select
        ActiveCell.FormulaR1C1 = _
            "=TRIM(SUBSTITUTE(SUBSTITUTE(RC[-1],""<"",""""),"">"",""""))"
        Range("B2").Select
        Selection.AutoFill Destination:=Range("B2:B" & count_rows_final + 1)
        
        
        
        
        ' Unique Values
        
        Range("C2").Select
        ActiveCell.Formula2R1C1 = "=UNIQUE(C[-1])"
        
        ' Resize
        Columns("B:C").Select
        Columns("B:C").EntireColumn.AutoFit
        
        Range("D3").Select
    
        
    Next fsFile
    
Done:
    Exit Sub
    
    
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"
    


End Sub

Sub getElementsToExcelDifferentSheet()
    
    On Error GoTo error_handle

    Dim fso As Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsFile As Scripting.File
    
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wDocTextBox As String
    Dim nNumber As Integer
    
    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    Dim wordDocData As String
       

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    
    ' Regex
    elRegex = InputBox("What is the Regex", "Enter The Regular Expression", "(<MC.*?>)")

    Set fso = CreateObject("Scripting.filesystemobject")
    Set fsFolder = fso.GetFolder(diaFolder.SelectedItems(1))
    
    
    
    startCell = 2
    
    For Each fsFile In fsFolder.Files
        
        Set wApp = CreateObject("word.Application")
        Set wDoc = wApp.Documents.Open(fsFile.path)
    
        Debug.Print fsFile.path
        
        ' Body Content
        wordDocData = wDoc.Content.Text
        
        
        ' Text Boxes
        With wDoc
            For nNumber = .Shapes.Count To 1 Step -1
                If .Shapes(nNumber).Type = msoTextBox Then
                    wDocTextBox = wDocTextBox & .Shapes(nNumber).TextFrame.TextRange.Text
                End If
            Next
        End With
        wordDocData = wordDocData & wDocTextBox
        wDocTextBox = ""
        
        
        ' Header
        With wDoc.Sections(1).Footers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        ' Footer
        With wDoc.Sections(1).Headers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        
        
        wDoc.Close
        wApp.Quit
        
        new_sheet_name = Left(fsFile.Name, (InStrRev(fsFile.Name, ".", -1, vbTextCompare) - 1))
        
        ' Create Sheet
        
        Sheets.Add After:=Sheets(Sheets.Count)
        
        new_sheet_name = Left(new_sheet_name, 20)
        
        Sheets(ActiveSheet.Name).Name = new_sheet_name
        
        
        el_list = getElement(wordDocData, elRegex)
        
        For i = 0 To UBound(el_list) - 1
            Sheets(new_sheet_name).Cells(2 + i, 1) = el_list(i, 0)
        Next i
        
    
        ' Trimming
        count_rows_final = WorksheetFunction.CountA(Columns(1))
        
        Range("B2").Select
        ActiveCell.FormulaR1C1 = _
            "=TRIM(SUBSTITUTE(SUBSTITUTE(RC[-1],""<"",""""),"">"",""""))"
        Range("B2").Select
        Selection.AutoFill Destination:=Range("B2:B" & count_rows_final + 1)
        
        
        
        
        ' Unique Values
        
        Range("C2").Select
        ActiveCell.Formula2R1C1 = "=UNIQUE(C[-1])"
        
        ' Resize
        Columns("B:C").Select
        Columns("B:C").EntireColumn.AutoFit
        
        Range("D3").Select
    
        
    Next fsFile
    
Done:
    Exit Sub
    
    
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"
    


End Sub
Sub PrintArray(Data As Variant, C1 As Range)
    'Debug.Print (Data.Type)
    C1.Resize(UBound(Data, 1), UBound(Data, 2)) = Data

End Sub


*/


// Elements_Same_Sheet

/*

Sub getAllMDToSameSheet()
    
    On Error GoTo error_handle
    
    Dim fso As Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsFile As Scripting.File
    
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wDocTextBox As String
    Dim nNumber As Integer

    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    Dim wordDocData As String
       

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    

    Set fso = CreateObject("Scripting.filesystemobject")
    Set fsFolder = fso.GetFolder(diaFolder.SelectedItems(1))
    
        
    new_sheet_name = "All MD"
    
    ' Create Sheet
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
    Sheets(ActiveSheet.Name).Name = new_sheet_name
    
    
    startCell = 2
    
    For Each fsFile In fsFolder.Files
        
        Set wApp = CreateObject("word.Application")
        Set wDoc = wApp.Documents.Open(fsFile.path)
    
        Debug.Print fsFile.path
        
        ' Body
        wordDocData = wDoc.Content.Text
        
        
        ' Text Boxes
        With wDoc
            For nNumber = .Shapes.Count To 1 Step -1
                If .Shapes(nNumber).Type = msoTextBox Then
                    wDocTextBox = wDocTextBox & .Shapes(nNumber).TextFrame.TextRange.Text
                End If
            Next
        End With
        wordDocData = wordDocData & wDocTextBox
        wDocTextBox = ""
        
        
        ' Header
        With wDoc.Sections(1).Footers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        ' Footer
        With wDoc.Sections(1).Headers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With

        wDoc.Close
        wApp.Quit
        
        md_list = getMD(wordDocData)
        
        count_rows = WorksheetFunction.CountA(Columns(1))
        
        Sheets(new_sheet_name).Cells(count_rows + startCell, 1) = fsFile.Name
        
        For i = 0 To UBound(md_list) - 1
            Sheets(new_sheet_name).Cells(count_rows + startCell + i + 1, 1) = md_list(i, 0)
        Next i
        
    Next fsFile

    
    
    ' Trimming
    count_rows_final = WorksheetFunction.CountA(Columns(1))
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=TRIM(SUBSTITUTE(SUBSTITUTE(RC[-1],""<"",""""),"">"",""""))"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & count_rows_final + 1)
    
    
    
    
    ' Unique Values
    
    Range("C2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(C[-1])"
    
    ' Resize
    Columns("B:C").Select
    Columns("B:C").EntireColumn.AutoFit
    
    Range("D3").Select
    
Done:
    Exit Sub
    
    
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"
    

End Sub


Sub getAllMCToSameSheet()
    
    On Error GoTo error_handle
    
    Dim fso As Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsFile As Scripting.File
    
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wDocTextBox As String
    Dim nNumber As Integer

    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    Dim wordDocData As String
       

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    

    Set fso = CreateObject("Scripting.filesystemobject")
    Set fsFolder = fso.GetFolder(diaFolder.SelectedItems(1))
    
        
    new_sheet_name = "All MC"
    
    ' Create Sheet
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
    Sheets(ActiveSheet.Name).Name = new_sheet_name
    
    
    startCell = 2
    
    For Each fsFile In fsFolder.Files
        
        Set wApp = CreateObject("word.Application")
        Set wDoc = wApp.Documents.Open(fsFile.path)
    
        Debug.Print fsFile.path
        
        ' Body
        wordDocData = wDoc.Content.Text
        
        
        ' Text Boxes
        With wDoc
            For nNumber = .Shapes.Count To 1 Step -1
                If .Shapes(nNumber).Type = msoTextBox Then
                    wDocTextBox = wDocTextBox & .Shapes(nNumber).TextFrame.TextRange.Text
                End If
            Next
        End With
        wordDocData = wordDocData & wDocTextBox
        wDocTextBox = ""
        
        
        ' Header
        With wDoc.Sections(1).Footers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        ' Footer
        With wDoc.Sections(1).Headers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With

        wDoc.Close
        wApp.Quit
        
        md_list = getMC(wordDocData)
        
        count_rows = WorksheetFunction.CountA(Columns(1))
        
        Sheets(new_sheet_name).Cells(count_rows + startCell, 1) = fsFile.Name
        
        For i = 0 To UBound(md_list) - 1
            Sheets(new_sheet_name).Cells(count_rows + startCell + i + 1, 1) = md_list(i, 0)
        Next i
        
    Next fsFile

    
    
    ' Trimming
    count_rows_final = WorksheetFunction.CountA(Columns(1))
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=TRIM(SUBSTITUTE(SUBSTITUTE(RC[-1],""<"",""""),"">"",""""))"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & count_rows_final + 1)
    
    
    
    
    ' Unique Values
    
    Range("C2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(C[-1])"
    
    ' Resize
    Columns("B:C").Select
    Columns("B:C").EntireColumn.AutoFit
    
    Range("D3").Select
    
Done:
    Exit Sub
    
    
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"
    

End Sub


Sub getAllElementsToSameSheet()
    
    On Error GoTo error_handle
    
    Dim fso As Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsFile As Scripting.File
    
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wDocTextBox As String
    Dim nNumber As Integer

    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    Dim wordDocData As String
       

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show
    
    
    ' Regex
    elRegex = InputBox("What is the Regex", "Enter The Regular Expression", "(<MC.*?>)")
    

    Set fso = CreateObject("Scripting.filesystemobject")
    Set fsFolder = fso.GetFolder(diaFolder.SelectedItems(1))
    
    ' Sheet Name
    new_sheet_name = "All Elements"
    
    ' Create Sheet
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ' Rename Sheet to particular name(new_sheet_name)
    Sheets(ActiveSheet.Name).Name = new_sheet_name
    
    
    startCell = 2
    
    For Each fsFile In fsFolder.Files
        
        Set wApp = CreateObject("word.Application")
        Set wDoc = wApp.Documents.Open(fsFile.path)
    
        Debug.Print fsFile.path
        
        ' Body
        wordDocData = wDoc.Content.Text
        
        
        ' Text Boxes
        With wDoc
            For nNumber = .Shapes.Count To 1 Step -1
                If .Shapes(nNumber).Type = msoTextBox Then
                    wDocTextBox = wDocTextBox & .Shapes(nNumber).TextFrame.TextRange.Text
                End If
            Next
        End With
        wordDocData = wordDocData & wDocTextBox
        wDocTextBox = ""
        
        
        ' Header
        With wDoc.Sections(1).Footers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        ' Footer
        With wDoc.Sections(1).Headers(wdHeaderFooterPrimary)
            If .Range.Text <> vbCr Then
                wordDocData = wordDocData & .Range.Text
            End If
        End With
        
        
        ' Close this Doc
        wDoc.Close
        wApp.Quit
        
        ' Get Elements list
        element_list = getElement(wordDocData, elRegex)
        
        ' Last unfilled row in Sheet
        count_rows = WorksheetFunction.CountA(Columns(1))
        
        ' Adding Doc Name to SHeet
        Sheets(new_sheet_name).Cells(count_rows + startCell, 1) = fsFile.Name
        
        ' Iterate and add Elements
        For i = 0 To UBound(element_list) - 1
            Sheets(new_sheet_name).Cells(count_rows + startCell + i + 1, 1) = element_list(i, 0)
        Next i
        
    Next fsFile

    
    
    ' Trimming
    count_rows_final = WorksheetFunction.CountA(Columns(1))
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=TRIM(SUBSTITUTE(SUBSTITUTE(RC[-1],""<"",""""),"">"",""""))"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & count_rows_final + 1)
    
    
    
    
    ' Unique Values
    
    Range("C2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(C[-1])"
    
    ' Resize
    Columns("B:C").Select
    Columns("B:C").EntireColumn.AutoFit
    
    Range("D3").Select
    
Done:
    Exit Sub
    
    ' Error Handling
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"
    

End Sub

*/


// Helper_Functions

/*

Function Eval(Ref As String)
    Application.Volatile
    Eval = Evaluate(Ref)
End Function




Function getMC(mc As String)


    Dim RE As Object
    Dim Match As Object
    Dim MatchCount As Object

    Dim returnArr As Variant


    Set RE = CreateObject("vbscript.regexp")
    RE.Pattern = "(<MC.*?>)"
    RE.Global = True
    RE.ignorecase = True

    Set Match = RE.Execute(mc)
    ReDim returnArr(Match.Count, 0)
    
    For i = 0 To Match.Count - 1
    
        returnArr(i, 0) = (Match.Item(i).submatches.Item(0))
    Next
    getMC = returnArr

End Function



Function getMD(mc As String)


    Dim RE As Object
    Dim Match As Object
    Dim MatchCount As Object

    Dim returnArr As Variant


    Set RE = CreateObject("vbscript.regexp")
    RE.Pattern = "(<MD.*?>)"
    RE.Global = True
    RE.ignorecase = True

    Set Match = RE.Execute(mc)
    ReDim returnArr(Match.Count, 0)
    
    For i = 0 To Match.Count - 1
        
        returnArr(i, 0) = (Match.Item(i).submatches.Item(0))
    Next

    getMD = returnArr

End Function



Function getElement(ByVal mc As String, ByVal elRegex As String)


    Dim RE As Object
    Dim Match As Object
    Dim MatchCount As Object

    Dim returnArr As Variant
    
    
    Set RE = CreateObject("vbscript.regexp")
    RE.Pattern = elRegex
    RE.Global = True
    RE.ignorecase = True

    Set Match = RE.Execute(mc)
    ReDim returnArr(Match.Count, 0)
    
    For i = 0 To Match.Count - 1
        
        returnArr(i, 0) = (Match.Item(i).submatches.Item(0))
    Next
    Debug.Print returnArr(0, 0)
    'Debug.Print returnArr(1, 0)

    getElement = returnArr

End Function



*/



// MCFromQI

/*

Sub getMCNameQISpec()

    On Error GoTo error_handle
    
    
    Dim thisWb As Workbook
    Dim qiWorkbook As Workbook
    
    Set thisWb = ActiveWorkbook
    
    
    ' Get QI Sheet Path from User
    
    qiSheetPath = InputBox("What is the location of QI Spec Sheet", "QI Spec Sheet", "C:\path\to\excel\sheet\QI SPec SHeet.xlsx")
       
    
    ' Add New Sheet
    Sheets.Add After:=Sheets(Sheets.Count)
    new_sheet_name = Sheets(ActiveSheet.Name).Name
    
    
    ' QI Workbook
    Set qiWorkbook = Workbooks.Open(qiSheetPath)
    
    
    ' Array to hold Sheet Values
    Dim sheets_to_get(5)
    
    ' Adding sheet names from QI Specs
    sheets_to_get(0) = "0-199"
    sheets_to_get(1) = "200-399"
    sheets_to_get(2) = "400-599"
    sheets_to_get(3) = "600-799"
    sheets_to_get(4) = "800-999"
    
    
    ' For each Sheet in sheet_names array ->
    For i = LBound(sheets_to_get) To UBound(sheets_to_get) - 1
        
        ' Get Last Row for the sheet in QI SPec sheet
        lastRowQiSheet = qiWorkbook.Sheets(sheets_to_get(i)).Cells(qiWorkbook.Sheets(sheets_to_get(i)).Rows.Count, "A").End(xlUp).Row
        ' Debug.Print lastRowQiSheet
        
        ' Copy Data from that sheet
        ' This only gets first 3 columns in QI Spec
        ' To Get More Change .Range("A2:C" & lastRowQiSheet) to .Range("A2:D" & lastRowQiSheet)
        ' The above change copies 4 columns
        qiWorkbook.Sheets(sheets_to_get(i)).Range("A2:D" & lastRowQiSheet).Copy
        
        ' Get last unfilled row in current sheet
        lastRowThisSheet = thisWb.Sheets(new_sheet_name).Cells(thisWb.Sheets(new_sheet_name).Rows.Count, "A").End(xlUp).Row + 1
        
        ' Pasting data to the new sheet
        thisWb.Sheets(new_sheet_name).Range("A" & lastRowThisSheet).PasteSpecial Paste:=xlPasteValues
    
    Next i
    
    ' Close the QI Work book
    qiWorkbook.Close (False)
    
Done:
    Exit Sub
    
    ' Error Handling
error_handle:
        MsgBox "Some error occured. Check if the path specified is correct"


End Sub

*/


// BAT Files

/*

echo off

@echo off
for /f "delims=" %%a in ('wmic OS Get localdatetime  ^| find "."') do set dt=%%a

set stamp2=%dt:~0,4%_%dt:~4,2%_%dt:~6,2%_%dt:~8,2%_%dt:~10,2%_%dt:~12,2%

echo Creating Folder : "%stamp2%"
cd "Output"
mkdir %stamp2%


cd "../Input Word Files"


for %%X in (*.doc) do cscript.exe //Nologo ../SaveAsPdf.js "%%X" "%stamp2%"


pause

*/


// SaveAsPdf.js

/*

var obj = new ActiveXObject("Scripting.FileSystemObject");

// Get Word Doc Name
var docPath = WScript.Arguments(0);
// Get OutputFolder Name
var outputFolder = WScript.Arguments(1);
docPath = obj.GetAbsolutePathName(docPath);

// Setting varables for Output PDF Folder
var pdfPath = docPath.replace(/\.doc[^.]*$/, ".pdf");
var pdfPath = pdfPath.replace("Input Word Files", "Output\\" + outputFolder);

var objWord = null;
// Creation
try {
  objWord = new ActiveXObject("Word.Application");
  objWord.Visible = false;

  var objDoc = objWord.Documents.Open(docPath);
  var format = 17;
  objDoc.SaveAs(pdfPath, format);
  objDoc.Close();
  WScript.Echo("------------------------------------");
  WScript.Echo("Saving\n" + docPath + "\n=======CONVERTED===============>>>>>>\n" + pdfPath);
  WScript.Echo("------------------------------------");
} finally {
  if (objWord != null) {
    objWord.Quit();
  }
}


*/
