$ErrorActionPreference = "Stop"

$addinName = "WordPdfCopyAddin.dotm"
$startupDir = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
$addinPath = Join-Path $startupDir $addinName
$workDir = Join-Path $env:TEMP ("WordPdfCopyAddin_" + [guid]::NewGuid().ToString("N"))
$vbaPath = Join-Path $workDir "WordPdfCopyAddin.bas"
$tempAddinPath = Join-Path $workDir $addinName

function Write-Info($message) {
    Write-Host "[Word PDF Copy Add-in] $message"
}

function Set-AccessVbom {
    $versions = @("16.0", "15.0", "14.0")
    foreach ($version in $versions) {
        $securityPath = "HKCU:\Software\Microsoft\Office\$version\Word\Security"
        if (Test-Path "HKCU:\Software\Microsoft\Office\$version\Word") {
            New-Item -Path $securityPath -Force | Out-Null
            New-ItemProperty -Path $securityPath -Name "AccessVBOM" -Value 1 -PropertyType DWord -Force | Out-Null
        }
    }
}

function Release-ComObject($object) {
    if ($null -ne $object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

function Add-ZipEntryFromString($zipPath, $entryName, $content) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $existing = $zip.GetEntry($entryName)
        if ($null -ne $existing) {
            $existing.Delete()
        }

        $entry = $zip.CreateEntry($entryName)
        $stream = $entry.Open()
        try {
            $writer = New-Object System.IO.StreamWriter($stream, [System.Text.UTF8Encoding]::new($false))
            try {
                $writer.Write($content)
            } finally {
                $writer.Dispose()
            }
        } finally {
            $stream.Dispose()
        }
    } finally {
        $zip.Dispose()
    }
}

function Upsert-RootRelationship($zipPath) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $relsEntryName = "_rels/.rels"
    $relationship = '<Relationship Id="rIdWordPdfCopyCustomUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>'
    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)

    try {
        $entry = $zip.GetEntry($relsEntryName)
        if ($null -eq $entry) {
            $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $relationship + '</Relationships>'
        } else {
            $stream = $entry.Open()
            try {
                $reader = New-Object System.IO.StreamReader($stream)
                try {
                    $rels = $reader.ReadToEnd()
                } finally {
                    $reader.Dispose()
                }
            } finally {
                $stream.Dispose()
            }

            $entry.Delete()
            if ($rels -notmatch "rIdWordPdfCopyCustomUI") {
                $rels = $rels -replace "</Relationships>", ($relationship + "</Relationships>")
            }
        }

        $newEntry = $zip.CreateEntry($relsEntryName)
        $newStream = $newEntry.Open()
        try {
            $writer = New-Object System.IO.StreamWriter($newStream, [System.Text.UTF8Encoding]::new($false))
            try {
                $writer.Write($rels)
            } finally {
                $writer.Dispose()
            }
        } finally {
            $newStream.Dispose()
        }
    } finally {
        $zip.Dispose()
    }
}

New-Item -ItemType Directory -Path $startupDir -Force | Out-Null
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host "Microsoft Word is currently running."
    Write-Host "Please save your documents, close Word, and run this installer again."
    exit 1
}

$vbaCode = @'
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
    RunCreatePdfAndCopy
End Sub

Public Sub CreatePdfAndCopyRibbon(control As IRibbonControl)
    RunCreatePdfAndCopy
End Sub

Private Sub RunCreatePdfAndCopy()
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
'@

$customUi = @'
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="WordPdfCopyTab" label="PDF作成">
        <group id="WordPdfCopyGroup" label="PDF作成">
          <button id="WordPdfCopyButton"
                  label="PDF作成"
                  size="large"
                  imageMso="FileSaveAsPdfOrXps"
                  screentip="PDF作成"
                  supertip="現在の変更を保存し、PDF と DOCX を作成してファイルとしてクリップボードにコピーします。"
                  onAction="CreatePdfAndCopyRibbon"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@

try {
    Write-Info "Preparing Word VBA access."
    Set-AccessVbom

    Write-Info "Writing VBA module."
    Set-Content -Path $vbaPath -Value $vbaCode -Encoding Default

    if (Test-Path $addinPath) {
        Remove-Item $addinPath -Force
    }

    $word = $null
    $document = $null
    try {
        Write-Info "Starting Word."
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        $word.Options.SaveNormalPrompt = $false
        $word.Options.ConfirmConversions = $false

        $document = $word.Documents.Add()
        $document.Content.Text = "Word PDF Copy Add-in"
        $document.VBProject.VBComponents.Import($vbaPath) | Out-Null

        Write-Info "Saving template add-in to temporary folder."
        $document.SaveAs($tempAddinPath, 15)
        Write-Info "Template add-in saved."
        $document.Close($false)
        $document = $null
    } finally {
        if ($null -ne $document) {
            $document.Close($false)
        }
        if ($null -ne $word) {
            $word.Quit()
        }
        Release-ComObject $document
        Release-ComObject $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    Write-Info "Adding ribbon UI."
    Add-ZipEntryFromString -zipPath $tempAddinPath -entryName "customUI/customUI.xml" -content $customUi
    Upsert-RootRelationship -zipPath $tempAddinPath

    Write-Info "Copying add-in to Word Startup folder."
    Copy-Item -Path $tempAddinPath -Destination $addinPath -Force

    Write-Info "Installed: $addinPath"
    Write-Info "Restart Microsoft Word."
} finally {
    if (Test-Path $workDir) {
        Remove-Item $workDir -Recurse -Force
    }
}
