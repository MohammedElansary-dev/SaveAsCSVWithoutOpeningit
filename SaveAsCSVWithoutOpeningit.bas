' === AutoSaveSheetAsUTF8CSVWithTimestamp ===
' Description:
' Saves the active sheet as a UTF-8 CSV file automatically,
' using the same name as the workbook PLUS a timestamp (e.g., "MyFile_20240625_1730.csv"),
' and saving to the same folder as the .xlsm file.

Sub AutoSaveSheetAsUTF8CSVWithTimestamp()

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

    ' Create timestamp in format yyyymmdd_HHmm (e.g., 20250625_1730)
    timestamp = Format(Now, "yyyy/mm/dd HH:mm")

    ' Build full CSV file path (you can customize text like "_Export_" if needed)
    csvPath = wbPath & wbNameNoExt & "_Export_" & timestamp & ".csv"

    ' Copy active sheet to new temporary workbook
    sourceSheet.Copy

    ' Save the new workbook as UTF-8 CSV (FileFormat:=62)
    ActiveWorkbook.SaveAs Filename:=csvPath, FileFormat:=62

    ' Close the temporary workbook without saving again
    ActiveWorkbook.Close SaveChanges:=False

    ' Optional: Notify user
    MsgBox "CSV file saved successfully:" & vbCrLf & csvPath, vbInformation

End Sub
