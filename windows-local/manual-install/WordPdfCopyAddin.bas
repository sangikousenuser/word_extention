Attribute VB_Name = "WordPdfCopyAddin"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub CopyMemoryPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Any, ByVal Length As Long)
    Private Declare Sub CopyMemoryPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const CF_HDROP As Long = 15
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = GMEM_MOVEABLE Or GMEM_ZEROINIT

Public Sub CreatePdfAndCopy()
    On Error GoTo ErrorHandler

    If Documents.Count = 0 Then
        MsgBox "Word 文書を開いてから実行してください。", vbExclamation, "PDF作成"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    If Len(doc.Path) = 0 Then
        Dialogs(wdDialogFileSaveAs).Show
        If Len(doc.Path) = 0 Then
            MsgBox "文書が保存されていないため、処理を中止しました。", vbExclamation, "PDF作成"
            Exit Sub
        End If
    Else
        doc.Save
    End If

    Dim outputDir As String
    outputDir = Environ$("USERPROFILE") & "\Documents\WordPdfCopyExports"
    EnsureFolder outputDir

    Dim baseName As String
    baseName = SanitizeFileName(GetBaseName(doc.Name))

    Dim stamp As String
    stamp = Format(Now, "yyyymmdd_hhnnss")

    Dim pdfPath As String
    Dim docxPath As String
    pdfPath = outputDir & "\" & baseName & "_" & stamp & ".pdf"
    docxPath = outputDir & "\" & baseName & "_" & stamp & ".docx"

    doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False
    doc.SaveCopyAs docxPath

    CopyFilesToClipboard Array(pdfPath, docxPath)

    MsgBox "PDF と DOCX を作成してクリップボードにコピーしました。" & vbCrLf & vbCrLf & pdfPath & vbCrLf & docxPath, vbInformation, "PDF作成"
    Exit Sub

ErrorHandler:
    MsgBox "PDF作成に失敗しました。" & vbCrLf & Err.Number & ": " & Err.Description, vbCritical, "PDF作成"
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Function GetBaseName(ByVal fileName As String) As String
    Dim position As Long
    position = InStrRev(fileName, ".")
    If position > 1 Then
        GetBaseName = Left$(fileName, position - 1)
    Else
        GetBaseName = fileName
    End If
End Function

Private Function SanitizeFileName(ByVal value As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")

    Dim item As Variant
    For Each item In invalidChars
        value = Replace(value, CStr(item), "_")
    Next item

    If Len(Trim$(value)) = 0 Then
        value = "document"
    End If

    SanitizeFileName = value
End Function

Private Sub CopyFilesToClipboard(ByVal files As Variant)
    Dim fileList As String
    Dim i As Long
    For i = LBound(files) To UBound(files)
        fileList = fileList & CStr(files(i)) & vbNullChar
    Next i
    fileList = fileList & vbNullChar

    Dim headerSize As Long
    headerSize = 20

    Dim totalBytes As LongPtr
    totalBytes = headerSize + LenB(fileList)

    Dim hGlobal As LongPtr
    hGlobal = GlobalAlloc(GHND, totalBytes)
    If hGlobal = 0 Then Err.Raise vbObjectError + 7001, , "クリップボード用メモリを確保できませんでした。"

    Dim lockedMemory As LongPtr
    lockedMemory = GlobalLock(hGlobal)
    If lockedMemory = 0 Then
        GlobalFree hGlobal
        Err.Raise vbObjectError + 7002, , "クリップボード用メモリをロックできませんでした。"
    End If

    Dim pFiles As Long
    Dim fWide As Long
    pFiles = headerSize
    fWide = 1

    CopyMemoryAny lockedMemory, pFiles, 4
    CopyMemoryAny lockedMemory + 16, fWide, 4
    CopyMemoryPtr lockedMemory + headerSize, StrPtr(fileList), LenB(fileList)
    GlobalUnlock hGlobal

    If OpenClipboard(0) = 0 Then
        GlobalFree hGlobal
        Err.Raise vbObjectError + 7003, , "クリップボードを開けませんでした。"
    End If

    EmptyClipboard
    If SetClipboardData(CF_HDROP, hGlobal) = 0 Then
        CloseClipboard
        GlobalFree hGlobal
        Err.Raise vbObjectError + 7004, , "クリップボードへファイルを設定できませんでした。"
    End If

    CloseClipboard
End Sub
