Attribute VB_Name = "CopyPasteArr"
Sub CopyPaste(CopyRange As Range, PasteRange As Range, Add As Boolean)
    
    CopyRange.Copy
    If Add Is True Then
        PasteRange.PasteSpecial xlPasteAll, xlPasteSpecialOperationAdd
    Else
        PasteRange.PasteSpecial xlPasteAll
    End If
    
End Sub
Sub FormatFile()

Dim ShtName(4) As String, SourceFile As String, DestFile As String
Dim CopyRng(4) As String, PasteRng(4) As String
Dim CopiedData As Variant, CopySheet As Worksheet, PasteSheet As Worksheet

SourceFile = ""
DestFile = ""
ShtName(0) = ""
ShtName(1) = ""
ShtName(2) = ""
ShtName(3) = ""
ShtName(4) = ""

CopyRng(0) = ""
CopyRng(1) = ""
CopyRng(2) = ""
CopyRng(3) = ""
CopyRng(4) = ""

PasteRng(0) = ""
PasteRng(1) = ""
PasteRng(2) = ""
PasteRng(3) = ""
PasteRng(4) = ""

For i = 0 To 4
    Set CopySheet = Workbooks(SourceFile).Worksheets(ShtName(i))
    Set PasteSheet = Workbooks(DestFile).Workheets(ShtName(i))
    For j = 0 To 4
        If j = 0 Then
            Call CopyPaste(CopySheet.Range(CopyRng(j)), PasteSheet.Range(PasteRng(j)), False)
        Else
            Call CopyPaste(CopySheet.Range(CopyRng(j)), PasteSheet.Range(PasteRng(j)), True)
        End If
    Next j
Next i

End Sub
