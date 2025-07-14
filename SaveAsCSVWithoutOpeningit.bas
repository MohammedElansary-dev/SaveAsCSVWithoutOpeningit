' === AutoSaveSheetAsUTF8CSVWithTimestamp ===
' Description:
' Saves the active sheet as a UTF-8 CSV file automatically,
' using the same name as the workbook PLUS a timestamp (e.g., "MyFile_20240625_1730.csv"),
' and saving to the same folder as the .xlsm file.
' Deletes the first row (headers or unwanted row) from the copied sheet before saving.

Sub AutoSaveSheetAsUTF8CSVWithTimestamp(ByRef control As Office.IRibbonControl)

    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet

    Dim wbPath As String
    Dim wbNameNoExt As String
    Dim csvPath As String
    Dim timestamp As String

    ' Get path of current workbook
    wbPath = ThisWorkbook.Path & Application.PathSeparator

    ' Get workbook name without extension
    wbNameNoExt = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' Create timestamp in format yyyy-mm-dd HH-mm
    timestamp = Format(Now, "yyyy-mm-dd HH-mm")

    ' Build full CSV file path
    csvPath = wbPath & wbNameNoExt & "_Export_" & timestamp & ".csv"

    ' Copy active sheet to new temporary workbook
    sourceSheet.Copy

    ' === Delete first row ===
    With ActiveWorkbook.Sheets(1)
        .Rows(1).Delete
    End With

    ' Save the new workbook as UTF-8 CSV (62 = xlCSVUTF8)
    ActiveWorkbook.SaveAs Filename:=csvPath, FileFormat:=62

    ' Close temporary workbook
    ActiveWorkbook.Close SaveChanges:=False

    ' Notify user
    MsgBox "CSV file saved successfully (first row removed):" & vbCrLf & csvPath, vbInformation

End Sub
