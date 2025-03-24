# Define paths and parameters
$pdfFolder = "C:\Users\rreshma\Documents\PowershellTask"       # Folder containing PDF files
$inputText = "innovation"                # Text to search for and highlight

# Load the Microsoft Word COM object
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

# Function to check if text exists in a PDF
function Test-TextInPdf {
    param (
        [string]$pdfFile,
        [string]$searchText
    )
    try {
        $document = $wordApp.Documents.Open($pdfFile)
        $range = $document.Content
        $found = $range.Find.Execute($searchText)
        $document.Close()
        return $found
    } catch {
        Write-Host "Failed to search text in: $pdfFile" -ForegroundColor Red
        return $false
    }
}

# Function to convert PDF to DOCX
function Convert-PdfToDocx {
    param (
        [string]$pdfFile,
        [string]$docxFile
    )
    try {
        $document = $wordApp.Documents.Open($pdfFile)
        $document.SaveAs([ref]$docxFile, [ref]16) # 16 = wdFormatDocumentDefault (DOCX format)
        $document.Close()
        return $true
    } catch {
        Write-Host "Failed to convert: $pdfFile" -ForegroundColor Red
        return $false
    }
}

# Function to highlight text in DOCX
function Highlight-TextInDocx {
    param (
        [string]$docxFile
    )
    try {
        $document = $wordApp.Documents.Open($docxFile)
        $range = $document.Content
        $found = $range.Find.Execute($inputText)
        while ($found) {
            $range.HighlightColorIndex = 7 # Yellow highlight (WdColorIndex enumeration)
            $range.Collapse(0) # Collapse the range to the end of the found text
            $found = $range.Find.Execute($inputText)
        }
        $document.Save()
        $document.Close()
        return $true
    } catch {
        Write-Host "Failed to highlight text in: $docxFile" -ForegroundColor Red
        return $false
    }
}

# Function to convert DOCX back to PDF
function Convert-DocxToPdf {
    param (
        [string]$docxFile,
        [string]$pdfFile
    )
    try {
        $document = $wordApp.Documents.Open($docxFile)
        $document.SaveAs([ref]$pdfFile, [ref]17) # 17 = wdFormatPDF (PDF format)
        $document.Close()
        return $true
    } catch {
        Write-Host "Failed to convert back to PDF: $docxFile" -ForegroundColor Red
        return $false
    }
}

# Iterate through all PDF files in the folder
Get-ChildItem -Path $pdfFolder -Filter *.pdf | ForEach-Object {
    $inputPdf = $_.FullName
    $baseName = $_.BaseName
    $tempDocx = Join-Path $pdfFolder "$baseName.docx"
    $outputPdf = $inputPdf # Replace the original PDF with the highlighted version

    Write-Host "Processing: $($_.Name)"

    # Step 1: Check if the input text exists in the PDF
    if (-Not (Test-TextInPdf -pdfFile $inputPdf -searchText $inputText)) {
        Write-Host "Skipping: Input text not found in $($_.Name)" -ForegroundColor Yellow
        return
    }

    Write-Host "Input text found. Proceeding with highlighting."

    # Step 2: Convert PDF to DOCX
    if (-Not (Convert-PdfToDocx -pdfFile $inputPdf -docxFile $tempDocx)) {
        Write-Host "Skipping: Conversion to DOCX failed for $($_.Name)" -ForegroundColor Yellow
        return
    }

    # Step 3: Highlight text in DOCX
    if (-Not (Highlight-TextInDocx -docxFile $tempDocx)) {
        Write-Host "Skipping: Highlighting failed for $($_.Name)" -ForegroundColor Yellow
        Remove-Item -Path $tempDocx -Force # Clean up temporary DOCX file
        return
    }

    # Step 4: Convert DOCX back to PDF
    if (-Not (Convert-DocxToPdf -docxFile $tempDocx -pdfFile $outputPdf)) {
        Write-Host "Skipping: Conversion back to PDF failed for $($_.Name)" -ForegroundColor Yellow
        Remove-Item -Path $tempDocx -Force # Clean up temporary DOCX file
        return
    }

    # Clean up temporary DOCX file
    Remove-Item -Path $tempDocx -Force

    Write-Host "Finished processing: $($_.Name)"
}

# Quit the Word application
$wordApp.Quit()

Write-Host "All PDFs processed."