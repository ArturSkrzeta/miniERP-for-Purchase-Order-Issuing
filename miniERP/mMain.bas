Attribute VB_Name = "mMain"
Option Explicit

Sub CreatePurchaseOrder()
    
    Call TurnOFFExcleFeatures
    ChDir ThisWorkbook.Path & "\New PRs"
    
    ' - - - Intro - - -
    Dim scriptWb As Workbook:               Set scriptWb = Excel.Application.ThisWorkbook
    Dim posTrackerWs As Worksheet:          Set posTrackerWs = posTracker
    ' - - - Loading PR - - -
    Dim filePath As String:                 filePath = Excel.Application.GetOpenFilename("Excel Files (*.xls*), *xls*", , "Browse for Purchase Request")
    If CStr(filePath) = "False" Then GoTo NoFileTakenError
    Dim prWb As Workbook: Set prWb = Workbooks.Open(filePath)
    Dim prWs As Worksheet: Set prWs = prWb.Sheets("Form")
    If prWs.Range("E12").Value = "" Then GoTo NoCompanyCodeOnForm
    ' - - - Generate PO Number - - -
    Dim companyCode As String:              companyCode = prWs.Range("E12").Value
    Dim newPoNumber As String: newPoNumber = GeneratePoNumber(companyCode)
    If CheckIfPoExist(newPoNumber & ".pdf") = True Then GoTo PoAlreadyExistError
    ' - - - Getting PR data - - -
    
    Dim i As Integer
    posTrackerWs.Activate
    Dim poIssueDate As Date:                poIssueDate = Date
    Dim prDetails(1 To 11) As String
    
    prDetails(1) = newPoNumber          ' ! PO Number
    prDetails(2) = poIssueDate          ' ! PO issued
    prDetails(3) = companyCode          ' ! Company
    prDetails(4) = prWs.Range("A23")    ' ! Requestor
    prDetails(5) = prWs.Range("D23")    ' ! Cost Center
    prDetails(6) = prWs.Range("I34")    ' ! Total
    prDetails(7) = prWs.Range("J34")    ' ! Currency
    prDetails(8) = prWs.Range("C4")     ' ! Supplier Full Name
    prDetails(9) = prWs.Range("C5")     ' ! Vendor SAP #
    prDetails(10) = prWs.Range("C9")    ' ! Vendor contact email
    prDetails(11) = prWs.Range("G12")   ' ! Delivery

    Dim insertedRowNumber As String: insertedRowNumber = posTrackerWs.Range("B100000").End(xlUp).Offset(1, 0).row
    
    ' - - - Tracking Data - - -
    
    posTrackerWs.Range("B100000").End(xlUp).Offset(1, 0).Value = prDetails(1)
    
    For i = 1 To UBound(prDetails, 1) - 1
        posTrackerWs.Range("B100000").End(xlUp).Offset(0, i) = prDetails(1 + i)
    Next i
    
    ' - - - Formatting table - - -
    
    Call RemoveTextWrapping(insertedRowNumber)
    posTrackerWs.Range("C" & insertedRowNumber).NumberFormat = "mm/dd/yyyy"
    
    ' - - - PO Number and PO Issued Date to PO pdf - - -
    
    prWs.Range("G23").Value = newPoNumber
    prWs.Range("F23").Value = poIssueDate
    
    ' - - - Printing PR to PO.pdf - - - -
    
    Call GenerateFolder("PO PDF Issued")
    Dim pdfPath As String:                  pdfPath = ThisWorkbook.Path & "\PO PDF Issued\" & newPoNumber & ".pdf"
    Call DeleteButtons(prWb)
    
    If CheckIfPoExist(newPoNumber & ".pdf") = False Then
        On Error Resume Next
        prWb.Sheets("Form").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        
        If Err <> 0 Then MsgBox "The PO has not been printed as there is such name already printed", vbExclamation
        On Error GoTo 0
    Else
        GoTo PoAlreadyExistError
    End If

    
    ' - - - Archiving Excel PR - - - -
    
    Call GenerateFolder("Archive PRs")
    Dim archivingFolderPath As String:      archivingFolderPath = ThisWorkbook.Path & "\Archive PRs"
    Dim wbPathForKilling As String: wbPathForKilling = prWb.Path & "\" & prWb.Name
    
    prWb.SaveAs archivingFolderPath & "\" & Left(prWb.Name, Len(prWb.Name) - 5) & " for PO " & newPoNumber & ".xlsm"
    prWb.Close savechanges:=True
    
    Kill wbPathForKilling
    
    ' - - - Closing Script
    
    scriptWb.Save

    Call TurnOnExcleFeatures

    Exit Sub

' - - -  Error Hanlders - - - -

NoFileTakenError:
    MsgBox "There was no Purchase Request file chosen. Macro stops working.", vbExclamation
    Call TurnOnExcleFeatures
    Exit Sub
    
PoAlreadyExistError:
    MsgBox "Watch Out! There is such PO with the name in the PO folder already.", vbExclamation
    prWb.Close savechanges:=False
    Call TurnOnExcleFeatures
    Exit Sub
    
NoCompanyCodeOnForm:
    MsgBox "There was no Company Code on Purchase Request Form.", vbExclamation
    prWb.Close savechanges:=False
    Call TurnOnExcleFeatures

End Sub

Private Sub DeleteButtons(prWb As Workbook)
    On Error Resume Next
    Dim sh As Shape
    For Each sh In prWb.Sheets("Form").Shapes
        If sh.Name = "btnClean" Or sh.Name = "btnSave" Then sh.Delete
    Next sh
    On Error GoTo 0
End Sub

Private Sub RemoveTextWrapping(rowNumber As String)
    
    Dim colCount As Integer:    colCount = posTracker.Range(Range("B6"), Range("B6").End(xlToRight)).Columns.Count
    Dim dataRng As Range:       Set dataRng = posTracker.Range(Range("B" & rowNumber), Range("B" & rowNumber).Offset(0, colCount - 1))
    
    With posTracker.Rows(rowNumber & ":" & rowNumber)
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    dataRng.Style = "Style 1"
    
End Sub

Private Sub GenerateFolder(folderName As String)

    Dim FSO As Object:              Set FSO = CreateObject("Scripting.Filesystemobject")
    Dim archivePath As String:      archivePath = ThisWorkbook.Path & "\" & folderName
    
    If Not FSO.FolderExists(archivePath) Then
        FSO.CreateFolder archivePath
    End If

End Sub

Private Function CheckIfPoExist(currentPoName As String) As Boolean

    Dim FSO As Object:          Set FSO = CreateObject("Scripting.Filesystemobject")
    Dim folder As Object:       Set folder = FSO.getfolder(ThisWorkbook.Path & "\PO PDF Issued")
    Dim file As Object
    
    For Each file In folder.Files
        If file.Name = currentPoName Then
            CheckIfPoExist = True
            Exit Function
        End If
    Next file
    
    CheckIfPoExist = False

End Function






