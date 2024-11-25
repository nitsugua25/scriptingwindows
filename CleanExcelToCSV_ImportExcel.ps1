# --------------------------------------------
# Script: CleanExcelToCSV_ReplaceDiacritics.ps1
# Description: Replaces diacritics with base characters in every cell of an Excel file, ensuring the column order remains the same, then exports the cleaned data to a CSV file.
# --------------------------------------------

# Function to check PowerShell version
Function Check-PowerShellVersion {
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Host "PowerShell version 5.0 or higher is required. Your version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
        exit
    }
}

# Function to ensure ImportExcel module is installed
Function Ensure-ImportExcelModule {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "The ImportExcel module is not installed. Installing it now..." -ForegroundColor Yellow
        try {
            Install-Module -Name ImportExcel -Force -Scope CurrentUser -AllowClobber
            Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to install ImportExcel module. Please install it manually." -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "ImportExcel module is already installed." -ForegroundColor Green
    }

    # Import the module
    try {
        Import-Module ImportExcel -ErrorAction Stop
    } catch {
        Write-Host "Failed to import ImportExcel module. Please ensure it's installed correctly." -ForegroundColor Red
        exit
    }
}

# Function to replace diacritics with base characters
Function Replace-Diacritics {
    param ([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    # Normalize to FormD to separate base characters from diacritics
    $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $normalized.ToCharArray()) {
        $char = [char]$c  # Ensure each character is treated as a single char
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            $null = $sb.Append($char)
        }
    }
    # Normalize back to FormC to recompose characters
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

# Main Script Execution
try {
    # Check PowerShell version
    Check-PowerShellVersion

    # Ensure ImportExcel module is available
    Ensure-ImportExcelModule

    # Prompt for input and output files
    $inputFile = Read-Host "Enter the full path to the input Excel (.xlsx) file"

    # Validate input file existence
    if (-not (Test-Path -Path $inputFile)) {
        Throw "Input file '$inputFile' does not exist."
    }

    $outputFile = Read-Host "Enter the full path for the output CSV file (with .csv extension)"

    # Ensure the output file has a .csv extension
    if (-not ($outputFile -match '\.csv$')) {
        $outputFile += '.csv'
        Write-Host "Output file extension '.csv' appended. New output file name: $outputFile" -ForegroundColor Yellow
    }

    # Import the Excel file using ImportExcel module
    Write-Host "Reading Excel file..." -ForegroundColor Cyan
    $Data = Import-Excel -Path $inputFile

    if ($Data.Count -eq 0) {
        Throw "The Excel file '$inputFile' is empty or does not contain any data."
    }

    # Get original headers
    $originalHeaders = $Data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

    if ($originalHeaders.Count -eq 0) {
        Throw "No headers found in the Excel file '$inputFile'. Ensure the first row contains headers."
    }

    # Clean headers by replacing diacritics
    Write-Host "Cleaning headers..." -ForegroundColor Cyan
    $cleanHeaders = $originalHeaders | ForEach-Object { 
        Replace-Diacritics $_
    }

    # Check for duplicate cleaned headers
    if ($cleanHeaders.Count -ne ($cleanHeaders | Select-Object -Unique).Count) {
        Throw "Duplicate headers found after cleaning. Please ensure headers are unique after replacing diacritics."
    }

    # Create a mapping from original headers to cleaned headers
    $headerMap = @{}
    for ($i = 0; $i -lt $originalHeaders.Count; $i++) {
        $headerMap[$originalHeaders[$i]] = $cleanHeaders[$i]
    }

    # Clean each cell in the dataset and create new objects with cleaned headers
    Write-Host "Processing data..." -ForegroundColor Cyan
    $cleanedData = foreach ($row in $Data) {
        # Use ordered hashtable to preserve column order
        $newObj = [ordered]@{}
        foreach ($originalHeader in $originalHeaders) {
            $cleanHeader = $headerMap[$originalHeader]
            $value = $row.$originalHeader
            if ($value -is [string]) {
                $cleanValue = Replace-Diacritics $value
            } else {
                $cleanValue = $value
            }
            $newObj[$cleanHeader] = $cleanValue
        }
        [PSCustomObject]$newObj
    }

    # Ensure column order is preserved by selecting properties in the order of cleanHeaders
    Write-Host "Exporting data with preserved column order..." -ForegroundColor Cyan
    $cleanedData | Select-Object -Property $cleanHeaders | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

    Write-Host "Process completed! Output saved to '$outputFile'" -ForegroundColor Green
} catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
