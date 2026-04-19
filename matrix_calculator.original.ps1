Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

Add-Type -AssemblyName System.Data
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Web.Extensions

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class DpiUtil {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@

[DpiUtil]::SetProcessDPIAware() | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:appRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
$script:moduleRoot = Join-Path $script:appRoot 'modules'

function S {
    param([string]$Text)
    return [System.Uri]::UnescapeDataString($Text)
}

$ui = @{
    Title = S '%E7%9F%A9%E9%98%B5%E5%B7%A5%E4%BD%9C%E5%8F%B0'
    Subtitle = S '%E5%A4%9A%E7%9F%A9%E9%98%B5%E8%BE%93%E5%85%A5%E4%B8%8E%E8%A1%A8%E8%BE%BE%E5%BC%8F%E8%AE%A1%E7%AE%97'
    MatrixCount = S '%E7%9F%A9%E9%98%B5%E6%95%B0%E9%87%8F'
    AddMatrix = S '%E5%A2%9E%E5%8A%A0'
    AddSpecial = S '%E7%89%B9%E6%AE%8A%E7%9F%A9%E9%98%B5'
    RemoveMatrix = S '%E5%87%8F%E5%B0%91'
    DataImport = S '%E6%95%B0%E6%8D%AE%E5%AF%BC%E5%85%A5%E4%B8%8E%E5%A4%84%E7%90%86'
    History = S '%E5%8E%86%E5%8F%B2%E8%AE%B0%E5%BD%95'
    Expression = S '%E8%A1%A8%E8%BE%BE%E5%BC%8F'
    Workspace = S '%E5%B7%A5%E4%BD%9C%E5%8C%BA'
    ExpressionWindow = S '%E8%A1%A8%E8%BE%BE%E5%BC%8F%E7%AA%97%E5%8F%A3'
    ResultWindow = S '%E7%BB%93%E6%9E%9C%E7%AA%97%E5%8F%A3'
    OpenExpression = S '%E8%A1%A8%E8%BE%BE%E5%BC%8F%E7%AA%97%E5%8F%A3'
    OpenResult = S '%E7%BB%93%E6%9E%9C%E7%AA%97%E5%8F%A3'
    Help = S '%E5%B8%AE%E5%8A%A9'
    DataWindow = S '%E6%95%B0%E6%8D%AE%E5%AF%BC%E5%85%A5%E4%B8%8E%E5%A4%84%E7%90%86'
    HistoryWindow = S '%E5%8E%86%E5%8F%B2%E8%AE%B0%E5%BD%95'
    HelpSearch = S '%E6%9F%A5%E8%AF%A2'
    HelpPrompt = S '%E8%BE%93%E5%85%A5%E5%87%BD%E6%95%B0%E3%80%81%E7%AE%97%E7%AC%A6%E6%88%96%E5%85%B3%E9%94%AE%E8%AF%8D'
    HelpEmpty = S '%E8%BE%93%E5%85%A5%E4%B8%8A%E6%96%B9%E5%85%B3%E9%94%AE%E8%AF%8D%E6%88%96%E7%82%B9%E5%87%BB%E5%BF%AB%E6%8D%B7%E6%9F%A5%E8%AF%A2%E6%8C%89%E9%92%AE%E3%80%82'
    HelpNotFound = S '%E6%9C%AA%E6%89%BE%E5%88%B0%E5%AF%B9%E5%BA%94%E8%AF%B4%E6%98%8E%EF%BC%8C%E8%AF%B7%E5%B0%9D%E8%AF%95%E8%BE%93%E5%85%A5%E5%87%BD%E6%95%B0%E5%90%8D%E3%80%81%E7%AE%97%E7%AC%A6%E6%88%96%E5%85%B3%E9%94%AE%E8%AF%8D%E3%80%82'
    SpecialConfig = S '%E7%89%B9%E6%AE%8A%E7%9F%A9%E9%98%B5%E8%AE%BE%E7%BD%AE'
    Apply = S '%E5%BA%94%E7%94%A8'
    Cancel = S '%E5%8F%96%E6%B6%88'
    Type = S '%E7%B1%BB%E5%9E%8B'
    Rows = S '%E8%A1%8C'
    Cols = S '%E5%88%97'
    Size = S '%E9%98%B6%E6%95%B0'
    DiagValues = S '%E5%AF%B9%E8%A7%92%E5%80%BC'
    ManualInput = S '%E6%89%8B%E5%8A%A8%E8%BE%93%E5%85%A5'
    Preset = S '%E9%A2%84%E8%AE%BE'
    ManualEditHint = S '%E5%8F%AF%E7%9B%B4%E6%8E%A5%E8%BE%93%E5%85%A5%E7%9F%A9%E9%98%B5%E3%80%81%E5%A4%9A%E7%BB%B4%E6%95%B0%E7%BB%84%E3%80%81%E7%89%B9%E6%AE%8A%E7%9F%A9%E9%98%B5%E7%AE%80%E5%86%99%E6%88%96%E5%88%86%E5%9D%97%E7%9F%A9%E9%98%B5%E3%80%82'
    BlockHint = S '%E5%88%86%E5%9D%97%E7%9F%A9%E9%98%B5%E7%A4%BA%E4%BE%8B%EF%BC%9A%5BA%20%7C%20B%3B%20B%20%7C%20A%5D'
    SpecialHint = S '%E7%89%B9%E6%AE%8A%E7%9F%A9%E9%98%B5%E7%AE%80%E5%86%99%EF%BC%9AI3%20%2F%20zeros(2%2C3)%20%2F%20ones(2%2C3)%20%2F%20diag(1%2C2%2C3)'
    Configure = S '%E8%AE%BE%E7%BD%AE'
    InputMode = S '%E8%BE%93%E5%85%A5%E6%A8%A1%E5%BC%8F'
    Identity = S '%E5%8D%95%E4%BD%8D%E7%9F%A9%E9%98%B5'
    ZeroMatrix = S '%E9%9B%B6%E7%9F%A9%E9%98%B5'
    OnesMatrix = S '%E5%85%A8%201%20%E7%9F%A9%E9%98%B5'
    DiagMatrix = S '%E5%AF%B9%E8%A7%92%E7%9F%A9%E9%98%B5'
    SpecialMatrixInvalid = S '%E7%89%B9%E6%AE%8A%E7%9F%A9%E9%98%B5%E5%8F%82%E6%95%B0%E6%97%A0%E6%95%88%E3%80%82'
    Run = S '%E8%AE%A1%E7%AE%97%E5%B9%B6%E6%98%BE%E7%A4%BA%E7%BB%93%E6%9E%9C'
    Clear = S '%E6%B8%85%E7%A9%BA'
    Sample = S '%E7%A4%BA%E4%BE%8B'
    Result = S '%E7%BB%93%E6%9E%9C'
    Ready = S '%E5%B0%B1%E7%BB%AA'
    ExprHint = S '%E8%BE%93%E5%85%A5%E5%AE%8C%E8%A1%A8%E8%BE%BE%E5%BC%8F%E5%90%8E%EF%BC%8C%E7%82%B9%E5%87%BB%E2%80%9C%E8%AE%A1%E7%AE%97%E5%B9%B6%E6%98%BE%E7%A4%BA%E7%BB%93%E6%9E%9C%E2%80%9D%E6%88%96%E7%9B%B4%E6%8E%A5%E6%8C%89%E5%9B%9E%E8%BD%A6%E3%80%82'
    Syntax = 'A*B, A.*B, A./B, A^2, EIGVALS(A), SVALS(A), LU(A), QR(A), ROWSUMS(A), COLMEANS(A), ROWMAXS(A), ' + (S '%E7%94%B2') + '*' + (S '%E4%B9%99') + ', ' + (S '%E8%BD%AC%E7%BD%AE') + '(A), RREF(A), NORM(A), I3, [A | B; B | A]'
    InputTip = '2D/ND + ' + (S '%E4%B8%AD%E6%96%87%E7%AC%A6%E5%8F%B7') + '; shortcuts: I3, ' + ((S '%E5%8D%95%E4%BD%8D%E9%98%B5') + '3') + ', zeros(2,3), diag(1,2,3), [A | B; B | A]'
    ParseEmptySuffix = S '%E4%B8%8D%E8%83%BD%E4%B8%BA%E7%A9%BA%E3%80%82'
    ParseInvalidSuffix = S '%E8%BE%93%E5%85%A5%E6%A0%BC%E5%BC%8F%E6%97%A0%E6%95%88%E3%80%82'
    ParseEmptyMatrixSuffix = S '%E5%BF%85%E9%A1%BB%E6%98%AF%E9%9D%9E%E7%A9%BA%E6%95%B0%E5%80%BC%E6%95%B0%E7%BB%84%E3%80%82'
    ParseEmptyRowSuffix = S '%E7%9A%84%E6%AF%8F%E4%B8%80%E5%B1%82%E9%83%BD%E5%BF%85%E9%A1%BB%E6%98%AF%E9%9D%9E%E7%A9%BA%E6%95%B0%E7%BB%84%E3%80%82'
    ParseRowLenSuffix = S '%E7%9A%84%E5%90%84%E5%B1%82%E9%95%BF%E5%BA%A6%E5%BF%85%E9%A1%BB%E4%B8%80%E8%87%B4%EF%BC%8C%E5%BD%A2%E6%88%90%E8%A7%84%E5%88%99%E6%95%B0%E7%BB%84%E3%80%82'
    ParseNumericSuffix = S '%E4%B8%AD%E5%8F%AA%E8%83%BD%E5%8C%85%E5%90%AB%E6%95%B0%E5%AD%97%E3%80%82'
    UnknownVar = S '%E6%9C%AA%E6%89%BE%E5%88%B0%E7%9F%A9%E9%98%B5%E5%8F%98%E9%87%8F%EF%BC%9A'
    NeedMatrix = S '%E8%AF%A5%E8%BF%90%E7%AE%97%E9%9C%80%E8%A6%81%E8%87%B3%E5%B0%91%E4%B8%A4%E7%BB%B4%E6%95%B0%E7%BB%84%EF%BC%8C%E5%B9%B6%E5%B0%86%E6%9C%80%E5%90%8E%E4%B8%A4%E7%BB%B4%E8%A7%86%E4%B8%BA%E7%9F%A9%E9%98%B5%E3%80%82'
    AddDim = S '%E9%80%90%E5%85%83%E7%B4%A0%E5%8A%A0%E6%B3%95%E8%A6%81%E6%B1%82%E4%B8%A4%E4%B8%AA%E6%95%B0%E7%BB%84%E5%BD%A2%E7%8A%B6%E5%AE%8C%E5%85%A8%E4%B8%80%E8%87%B4%E3%80%82'
    SubDim = S '%E9%80%90%E5%85%83%E7%B4%A0%E5%87%8F%E6%B3%95%E8%A6%81%E6%B1%82%E4%B8%A4%E4%B8%AA%E6%95%B0%E7%BB%84%E5%BD%A2%E7%8A%B6%E5%AE%8C%E5%85%A8%E4%B8%80%E8%87%B4%E3%80%82'
    SameShape = S '%E9%80%90%E5%85%83%E7%B4%A0%E8%BF%90%E7%AE%97%E8%A6%81%E6%B1%82%E4%B8%A4%E4%B8%AA%E6%95%B0%E7%BB%84%E5%BD%A2%E7%8A%B6%E5%AE%8C%E5%85%A8%E4%B8%80%E8%87%B4%E3%80%82'
    MulDim = S '%E4%B9%98%E6%B3%95%E8%A6%81%E6%B1%82%E5%B7%A6%E7%9F%A9%E9%98%B5%E5%88%97%E6%95%B0%E7%AD%89%E4%BA%8E%E5%8F%B3%E7%9F%A9%E9%98%B5%E8%A1%8C%E6%95%B0%E3%80%82'
    DivZero = S '%E9%99%A4%E6%95%B0%E4%B8%8D%E8%83%BD%E4%B8%BA%E9%9B%B6%E3%80%82'
    BatchDim = S '%E6%89%B9%E9%87%8F%E7%9F%A9%E9%98%B5%E8%BF%90%E7%AE%97%E8%A6%81%E6%B1%82%E5%89%8D%E7%BD%AE%E7%BB%B4%E5%BA%A6%E4%B8%80%E8%87%B4%E3%80%82'
    SquareOnly = S '%E5%8F%AA%E6%9C%89%E6%96%B9%E9%98%B5%E6%89%8D%E8%83%BD%E8%BF%9B%E8%A1%8C%E8%AF%A5%E8%BF%90%E7%AE%97%E3%80%82'
    DetSquare = S '%E5%8F%AA%E6%9C%89%E6%96%B9%E9%98%B5%E6%89%8D%E8%83%BD%E6%B1%82%E8%A1%8C%E5%88%97%E5%BC%8F%E3%80%82'
    InvSquare = S '%E5%8F%AA%E6%9C%89%E6%96%B9%E9%98%B5%E6%89%8D%E8%83%BD%E6%B1%82%E9%80%86%E7%9F%A9%E9%98%B5%E3%80%82'
    IntegerPowerOnly = S '%E7%9F%A9%E9%98%B5%E5%B9%82%E8%BF%90%E7%AE%97%E5%8F%AA%E6%94%AF%E6%8C%81%E6%95%B4%E6%95%B0%E6%AC%A1%E5%B9%82%E3%80%82'
    NotInvertible = S '%E8%AF%A5%E7%9F%A9%E9%98%B5%E4%B8%8D%E5%8F%AF%E9%80%86%E3%80%82'
    UnexpectedToken = S '%E8%A1%A8%E8%BE%BE%E5%BC%8F%E5%90%AB%E6%9C%89%E6%97%A0%E6%95%88%E7%AC%A6%E5%8F%B7%E3%80%82'
    MissingParen = S '%E7%BC%BA%E5%B0%91%E5%8F%B3%E6%8B%AC%E5%8F%B7%E3%80%82'
    UnsupportedOp = S '%E4%B8%8D%E6%94%AF%E6%8C%81%E8%AF%A5%E8%BF%90%E7%AE%97%E7%BB%84%E5%90%88%E3%80%82'
    EvalFailed = S '%E8%AE%A1%E7%AE%97%E5%A4%B1%E8%B4%A5'
    Done = S '%E8%AE%A1%E7%AE%97%E5%AE%8C%E6%88%90'
    TrivialNull = S '%E9%9B%B6%E7%A9%BA%E9%97%B4%EF%BC%9A%E5%8F%AA%E5%90%AB%E9%9B%B6%E5%90%91%E9%87%8F'
    MatrixShortcutHelp = 'Shortcuts: I3 / ' + ((S '%E5%8D%95%E4%BD%8D%E9%98%B5') + '3') + ' / zeros(2,3) / ones(2,3) / diag(1,2,3)'
    SpecialButtonText = S '%E7%89%B9%E6%AE%8A'
    LoadExpression = S '%E8%BD%BD%E5%85%A5%E8%A1%A8%E8%BE%BE%E5%BC%8F'
    HistoryEmpty = S '%E6%9A%82%E6%97%A0%E5%8E%86%E5%8F%B2%E8%AE%B0%E5%BD%95%E3%80%82'
    ImportFile = S '%E5%AF%BC%E5%85%A5%E6%96%87%E4%BB%B6'
    Browse = S '%E6%B5%8F%E8%A7%88'
    DataStatusReady = S '%E6%9C%AA%E5%AF%BC%E5%85%A5%E6%95%B0%E6%8D%AE'
    DataLoaded = S '%E6%95%B0%E6%8D%AE%E5%B7%B2%E8%BD%BD%E5%85%A5'
    DataImportFailed = S '%E5%AF%BC%E5%85%A5%E5%A4%B1%E8%B4%A5'
    DataExported = S '%E5%AF%BC%E5%87%BA%E5%AE%8C%E6%88%90'
    NoDataLoaded = S '%E8%AF%B7%E5%85%88%E5%AF%BC%E5%85%A5%E6%95%B0%E6%8D%AE%E3%80%82'
    NeedFile = S '%E8%AF%B7%E5%85%88%E9%80%89%E6%8B%A9%E6%96%87%E4%BB%B6%E3%80%82'
    FileNotSupported = S '%E5%BD%93%E5%89%8D%E4%BB%85%E6%94%AF%E6%8C%81%20CSV%20/%20TXT%20/%20XLSX%20/%20XLS%20/%20XLSM%E3%80%82'
    DataRange = S '%E6%95%B0%E6%8D%AE%E8%8C%83%E5%9B%B4'
    StartRow = S '%E8%B5%B7%E5%A7%8B%E8%A1%8C'
    StartCol = S '%E8%B5%B7%E5%A7%8B%E5%88%97'
    EndRow = S '%E7%BB%93%E6%9D%9F%E8%A1%8C'
    EndCol = S '%E7%BB%93%E6%9D%9F%E5%88%97'
    ImportAllHint = S '%E9%BB%98%E8%AE%A4%E5%85%88%E5%85%A8%E9%83%A8%E5%AF%BC%E5%85%A5%EF%BC%8C%E5%AF%BC%E5%85%A5%E5%90%8E%E5%8F%AF%E5%9C%A8%E2%80%9C%E9%81%B4%E9%80%89%E8%A1%8C%E5%88%97%E2%80%9D%E4%B8%AD%E5%86%8D%E8%AE%BE%E7%BD%AE%E8%8C%83%E5%9B%B4%E3%80%82'
    FirstRowHeader = S '%E9%A6%96%E8%A1%8C%E4%BD%9C%E4%B8%BA%E5%88%97%E5%90%8D'
    SheetName = S '%E5%B7%A5%E4%BD%9C%E8%A1%A8'
    DataOverview = S '%E6%95%B0%E6%8D%AE%E6%A6%82%E8%A7%88'
    MissingRatio = S '%E7%BC%BA%E5%A4%B1%E6%AF%94%E4%BE%8B'
    CleanMissing = S '%E7%BC%BA%E5%A4%B1%E5%80%BC%E6%B8%85%E6%B4%97'
    SelectRowsCols = S '%E9%81%B4%E9%80%89%E8%A1%8C%E5%88%97'
    EditData = S '%E4%BF%AE%E6%94%B9%E6%95%B0%E6%8D%AE'
    GroupStats = S '%E5%88%86%E7%BB%84%E7%BB%9F%E8%AE%A1'
    PlotData = S '%E6%95%B0%E6%8D%AE%E7%BB%98%E5%9B%BE'
    ExportData = S '%E5%AF%BC%E5%87%BA%E6%95%B0%E6%8D%AE/%E5%9B%BE%E7%89%87'
    PlotPlaceholder = S '%E7%BB%98%E5%9B%BE%E5%8A%9F%E8%83%BD%E9%A1%B5%E9%9D%A2%E5%B7%B2%E9%A2%84%E7%95%99%EF%BC%8C1.1.3%20%E7%89%88%E6%9C%AC%E5%86%8D%E7%BB%A7%E7%BB%AD%E5%BC%80%E5%8F%91%E3%80%82'
    ExportCsv = S '%E5%AF%BC%E5%87%BACSV'
    ExportImage = S '%E5%AF%BC%E5%87%BA%E5%9B%BE%E7%89%87'
    RemoveMissingRows = S '%E5%88%A0%E9%99%A4%E5%90%AB%E7%BC%BA%E5%A4%B1%E8%A1%8C'
    FillMissingZero = S '%E7%BC%BA%E5%A4%B1%E5%A1%AB0'
    FillMissingMean = S '%E7%BC%BA%E5%A4%B1%E5%A1%AB%E5%9D%87%E5%80%BC'
    GroupBy = S '%E5%88%86%E7%BB%84%E5%88%97'
    AggregateColumn = S '%E7%BB%9F%E8%AE%A1%E5%88%97'
    AggregateMethod = S '%E7%BB%9F%E8%AE%A1%E6%96%B9%E5%BC%8F'
}

function Normalize-CommonSymbols {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $normalized = $Text
    $replacements = @(
        @([char]0xFF08, '('),
        @([char]0xFF09, ')'),
        @([char]0xFF3B, '['),
        @([char]0xFF3D, ']'),
        @([char]0x3010, '['),
        @([char]0x3011, ']'),
        @([char]0xFF5B, '['),
        @([char]0xFF5D, ']'),
        @([char]0xFF0C, ','),
        @([char]0xFF1B, ';'),
        @([char]0xFF1A, ':'),
        @([char]0xFF0B, '+'),
        @([char]0xFF0D, '-'),
        @([char]0x2014, '-'),
        @([char]0x2013, '-'),
        @([char]0xFF0A, '*'),
        @([char]0x00D7, '*'),
        @([char]0xFF0F, '/'),
        @([char]0x00F7, '/'),
        @([char]0xFF5C, '|'),
        @([char]0x2223, '|'),
        @([char]0xFF3E, '^'),
        @([char]0xFF0E, '.'),
        @([char]0x3000, ' '),
        @([char]0x3001, ',')
    )
    foreach ($replacement in $replacements) {
        $normalized = $normalized.Replace([string]$replacement[0], [string]$replacement[1])
    }
    return $normalized
}

function Convert-MatlabMatrixToJsonLike {
    param([string]$Text)
    $trimmed = (Normalize-CommonSymbols $Text).Trim()
    if ($trimmed -notmatch '^\[.*\]$') { return $trimmed }
    if ($trimmed -match '^\s*\[\s*\[') { return $trimmed }
    $inner = $trimmed.Substring(1, $trimmed.Length - 2).Trim()
    if ([string]::IsNullOrWhiteSpace($inner)) { return $trimmed }
    $rows = @()
    foreach ($rowText in ($inner -split ';')) {
        $cleanRow = ($rowText.Trim() -replace ',', ' ')
        if ([string]::IsNullOrWhiteSpace($cleanRow)) { continue }
        $parts = @($cleanRow -split '\s+' | Where-Object { $_ -ne '' })
        $rows += ,('[' + ($parts -join ',') + ']')
    }
    if ($rows.Count -eq 0) { return $trimmed }
    return '[' + ($rows -join ',') + ']'
}

function Get-CanonicalIdentifier {
    param([string]$Name)
    $normalized = (Normalize-CommonSymbols $Name).Trim()
    $upper = $normalized.ToUpperInvariant()
    $aliasGroups = @(
        @{ Canonical = 'A'; Aliases = @('A', [string][char]0xFF21, (S '%E7%94%B2'), ((S '%E7%9F%A9%E9%98%B5') + 'A')) },
        @{ Canonical = 'B'; Aliases = @('B', [string][char]0xFF22, (S '%E4%B9%99'), ((S '%E7%9F%A9%E9%98%B5') + 'B')) },
        @{ Canonical = 'C'; Aliases = @('C', [string][char]0xFF23, (S '%E4%B8%99'), ((S '%E7%9F%A9%E9%98%B5') + 'C')) },
        @{ Canonical = 'D'; Aliases = @('D', [string][char]0xFF24, (S '%E4%B8%81'), ((S '%E7%9F%A9%E9%98%B5') + 'D')) },
        @{ Canonical = 'E'; Aliases = @('E', [string][char]0xFF25, (S '%E6%88%8A'), ((S '%E7%9F%A9%E9%98%B5') + 'E')) },
        @{ Canonical = 'F'; Aliases = @('F', [string][char]0xFF26, (S '%E5%B7%B1'), ((S '%E7%9F%A9%E9%98%B5') + 'F')) },
        @{ Canonical = 'G'; Aliases = @('G', [string][char]0xFF27, (S '%E5%BA%9A'), ((S '%E7%9F%A9%E9%98%B5') + 'G')) },
        @{ Canonical = 'H'; Aliases = @('H', [string][char]0xFF28, (S '%E8%BE%9B'), ((S '%E7%9F%A9%E9%98%B5') + 'H')) },
        @{ Canonical = 'T'; Aliases = @('T', 'TRANSPOSE', 'TRANS', (S '%E8%BD%AC%E7%BD%AE')) },
        @{ Canonical = 'REF'; Aliases = @('REF', 'ROWECHELON', (S '%E8%A1%8C%E9%98%B6%E6%A2%AF'), (S '%E9%98%B6%E6%A2%AF%E5%BD%A2')) },
        @{ Canonical = 'RREF'; Aliases = @('RREF', 'REDUCEDROWECHELON', (S '%E6%9C%80%E7%AE%80%E8%A1%8C%E9%98%B6%E6%A2%AF'), (S '%E8%A1%8C%E6%9C%80%E7%AE%80'), (S '%E6%9C%80%E7%AE%80')) },
        @{ Canonical = 'RANK'; Aliases = @('RANK', (S '%E7%A7%A9')) },
        @{ Canonical = 'TR'; Aliases = @('TR', 'TRACE', (S '%E8%BF%B9')) },
        @{ Canonical = 'NORM'; Aliases = @('NORM', (S '%E8%8C%83%E6%95%B0')) },
        @{ Canonical = 'INV'; Aliases = @('INV', 'INVERSE', (S '%E9%80%86'), (S '%E9%80%86%E7%9F%A9%E9%98%B5')) },
        @{ Canonical = 'DET'; Aliases = @('DET', 'DETERMINANT', (S '%E8%A1%8C%E5%88%97%E5%BC%8F')) },
        @{ Canonical = 'COF'; Aliases = @('COF', 'COFACTOR', (S '%E4%BD%99%E5%AD%90%E5%BC%8F'), (S '%E4%BB%A3%E6%95%B0%E4%BD%99%E5%AD%90%E5%BC%8F')) },
        @{ Canonical = 'ADJ'; Aliases = @('ADJ', 'ADJOINT', 'ADJUGATE', (S '%E4%BC%B4%E9%9A%8F'), (S '%E4%BC%B4%E9%9A%8F%E7%9F%A9%E9%98%B5')) },
        @{ Canonical = 'NULL'; Aliases = @('NULL', 'NULLSPACE', (S '%E9%9B%B6%E7%A9%BA%E9%97%B4')) },
        @{ Canonical = 'ROWSUMS'; Aliases = @('ROWSUMS', 'ROWSUM', 'SUMROWS', (S '%E8%A1%8C%E5%92%8C')) },
        @{ Canonical = 'COLSUMS'; Aliases = @('COLSUMS', 'COLSUM', 'SUMCOLS', 'SUMCOLUMNS', (S '%E5%88%97%E5%92%8C')) },
        @{ Canonical = 'ROWMEANS'; Aliases = @('ROWMEANS', 'ROWMEAN', 'MEANROWS', (S '%E8%A1%8C%E5%9D%87%E5%80%BC')) },
        @{ Canonical = 'COLMEANS'; Aliases = @('COLMEANS', 'COLMEAN', 'MEANCOLS', 'MEANCOLUMNS', (S '%E5%88%97%E5%9D%87%E5%80%BC')) },
        @{ Canonical = 'ROWMAXS'; Aliases = @('ROWMAXS', 'ROWMAX', 'MAXROWS', (S '%E8%A1%8C%E6%9C%80%E5%A4%A7%E5%80%BC')) },
        @{ Canonical = 'COLMAXS'; Aliases = @('COLMAXS', 'COLMAX', 'MAXCOLS', 'MAXCOLUMNS', (S '%E5%88%97%E6%9C%80%E5%A4%A7%E5%80%BC')) },
        @{ Canonical = 'ROWMINS'; Aliases = @('ROWMINS', 'ROWMIN', 'MINROWS', (S '%E8%A1%8C%E6%9C%80%E5%B0%8F%E5%80%BC')) },
        @{ Canonical = 'COLMINS'; Aliases = @('COLMINS', 'COLMIN', 'MINCOLS', 'MINCOLUMNS', (S '%E5%88%97%E6%9C%80%E5%B0%8F%E5%80%BC')) },
        @{ Canonical = 'EIGVALS'; Aliases = @('EIGVALS', 'EIGENVALUES', (S '%E7%89%B9%E5%BE%81%E5%80%BC')) },
        @{ Canonical = 'EIGVECS'; Aliases = @('EIGVECS', 'EIGENVECTORS', (S '%E7%89%B9%E5%BE%81%E5%90%91%E9%87%8F')) },
        @{ Canonical = 'SVALS'; Aliases = @('SVALS', 'SINGULARVALUES', 'SIGMAVALS', (S '%E5%A5%87%E5%BC%82%E5%80%BC')) },
        @{ Canonical = 'CHOL'; Aliases = @('CHOL', 'CHOLESKY', (S '%E6%A5%9A%E5%88%97%E6%96%AF%E5%9F%BA%E5%88%86%E8%A7%A3')) },
        @{ Canonical = 'LU'; Aliases = @('LU', (S '%E4%B8%89%E8%A7%92%E5%88%86%E8%A7%A3'), (S '%E4%B8%8B%E4%B8%8A%E5%88%86%E8%A7%A3')) },
        @{ Canonical = 'QR'; Aliases = @('QR', (S '%E6%AD%A3%E4%BA%A4%E5%88%86%E8%A7%A3')) }
    )
    foreach ($group in $aliasGroups) {
        foreach ($alias in $group.Aliases) {
            if ($upper -eq $alias.ToUpperInvariant()) { return $group.Canonical }
        }
    }
    return $upper
}

function Test-IsTensorCollection {
    param($Node)
    return ($null -ne $Node -and $Node -is [System.Collections.IEnumerable] -and $Node -isnot [string] -and $Node -isnot [System.Collections.IDictionary])
}

function Get-EnumerableItems {
    param($Node)
    if ($null -eq $Node) { return ,@() }
    return ,(@(foreach ($item in $Node) { ,$item }))
}


function Test-ShapeEqual {
    param($LeftShape, $RightShape)
    $left = @($LeftShape)
    $right = @($RightShape)
    if ($left.Count -ne $right.Count) { return $false }
    for ($i = 0; $i -lt $left.Count; $i++) {
        if ([int]$left[$i] -ne [int]$right[$i]) { return $false }
    }
    return $true
}

function Convert-ToTensorData {
    param($Node, [string]$Name)
    if (Test-IsTensorCollection $Node) {
        $items = Get-EnumerableItems $Node
        if ($items.Count -eq 0) { throw ($Name + $ui.ParseEmptyRowSuffix) }
        $list = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $items) {
            $list.Add((Convert-ToTensorData -Node $item -Name $Name))
        }
        return ,$list
    }
    if ($null -eq $Node) { throw ($Name + $ui.ParseNumericSuffix) }
    $number = 0.0
    $styles = [System.Globalization.NumberStyles]::Float
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $current = [System.Globalization.CultureInfo]::CurrentCulture
    if (-not [double]::TryParse($Node.ToString(), $styles, $invariant, [ref]$number) -and
        -not [double]::TryParse($Node.ToString(), $styles, $current, [ref]$number)) {
        throw ($Name + $ui.ParseNumericSuffix)
    }
    return [double]$number
}

function Get-TensorShape {
    param($Node, [string]$Name, [int]$Depth = 0)
    if (-not (Test-IsTensorCollection $Node)) { return @() }
    $items = Get-EnumerableItems $Node
    if ($items.Count -eq 0) {
        if ($Depth -eq 0) { throw ($Name + $ui.ParseEmptyMatrixSuffix) }
        throw ($Name + $ui.ParseEmptyRowSuffix)
    }
    $childShape = $null
    foreach ($item in $items) {
        $shape = @(Get-TensorShape -Node $item -Name $Name -Depth ($Depth + 1))
        if ($null -eq $childShape) {
            $childShape = $shape
        } elseif (-not (Test-ShapeEqual -LeftShape $childShape -RightShape $shape)) {
            throw ($Name + $ui.ParseRowLenSuffix)
        }
    }
    return @($items.Count) + @($childShape)
}

function New-ArrayData {
    param($Data, [int[]]$Shape)
    return @{ Data = $Data; Shape = @($Shape) }
}

function New-OnesMatrix {
    param([int]$Rows, [int]$Cols)
    $matrix = New-ZeroMatrix -Rows $Rows -Cols $Cols
    for ($r = 0; $r -lt $Rows; $r++) {
        for ($c = 0; $c -lt $Cols; $c++) { $matrix[$r][$c] = 1.0 }
    }
    return ,$matrix
}

function New-DiagonalMatrix {
    param([double[]]$Values)
    $matrix = New-ZeroMatrix -Rows $Values.Count -Cols $Values.Count
    for ($i = 0; $i -lt $Values.Count; $i++) { $matrix[$i][$i] = [double]$Values[$i] }
    return ,$matrix
}

function Try-ParseSpecialMatrixInput {
    param([string]$Text, [string]$Name)
    $normalized = (Normalize-CommonSymbols $Text).Trim()
    $compact = $normalized -replace '\s+', ''

    $identityPattern = '^(?:' + ((@('I', 'ID', 'EYE', 'IDENTITY', 'UNIT', (S '%E5%8D%95%E4%BD%8D%E9%98%B5'), (S '%E5%8D%95%E4%BD%8D%E7%9F%A9%E9%98%B5')) | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\(?(\d+)\)?$'
    $zeroPattern = '^(?:' + ((@('Z', 'ZERO', 'ZEROS', (S '%E9%9B%B6%E7%9F%A9%E9%98%B5'), (S '%E5%85%A8%E9%9B%B6%E7%9F%A9%E9%98%B5')) | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\(?(\d+)(?:,|X)(\d+)\)?$'
    $onesPattern = '^(?:' + ((@('O', 'ONE', 'ONES', (S '%E5%85%A81%E7%9F%A9%E9%98%B5'), (S '%E5%85%A8%E4%B8%80%E7%9F%A9%E9%98%B5')) | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\(?(\d+)(?:,|X)(\d+)\)?$'
    $diagPattern = '^(?:' + ((@('DIAG', (S '%E5%AF%B9%E8%A7%92'), (S '%E5%AF%B9%E8%A7%92%E9%98%B5'), (S '%E5%AF%B9%E8%A7%92%E7%9F%A9%E9%98%B5')) | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\((.+)\)$'

    if ($compact -match $identityPattern) {
        $size = [int]$matches[1]
        return New-ArrayData -Data (New-IdentityMatrix -Size $size) -Shape @($size, $size)
    }
    if ($compact -match $zeroPattern) {
        $rows = [int]$matches[1]
        $cols = [int]$matches[2]
        return New-ArrayData -Data (New-ZeroMatrix -Rows $rows -Cols $cols) -Shape @($rows, $cols)
    }
    if ($compact -match $onesPattern) {
        $rows = [int]$matches[1]
        $cols = [int]$matches[2]
        return New-ArrayData -Data (New-OnesMatrix -Rows $rows -Cols $cols) -Shape @($rows, $cols)
    }
    if ($compact -match $diagPattern) {
        $valueText = $matches[1]
        $parts = @($valueText -split '[,;]' | Where-Object { $_ -ne '' })
        if ($parts.Count -eq 0) { throw ($Name + $ui.ParseInvalidSuffix) }
        $values = New-Object 'System.Collections.Generic.List[double]'
        foreach ($part in $parts) {
            $number = 0.0
            if (-not [double]::TryParse($part, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and
                -not [double]::TryParse($part, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
                throw ($Name + $ui.ParseNumericSuffix)
            }
            $values.Add($number)
        }
        return New-ArrayData -Data (New-DiagonalMatrix -Values $values.ToArray()) -Shape @($values.Count, $values.Count)
    }

    return $null
}

function Split-TopLevelText {
    param([string]$Text, [char[]]$Separators, [switch]$WhitespaceSeparator)
    $segments = New-Object 'System.Collections.Generic.List[string]'
    $buffer = New-Object System.Text.StringBuilder
    $squareDepth = 0
    $parenDepth = 0
    foreach ($ch in $Text.ToCharArray()) {
        switch ($ch) {
            '[' { $squareDepth++; [void]$buffer.Append($ch); continue }
            ']' { $squareDepth--; [void]$buffer.Append($ch); continue }
            '(' { $parenDepth++; [void]$buffer.Append($ch); continue }
            ')' { $parenDepth--; [void]$buffer.Append($ch); continue }
        }
        $isSeparator = ($squareDepth -eq 0 -and $parenDepth -eq 0 -and ($Separators -contains $ch -or ($WhitespaceSeparator -and [char]::IsWhiteSpace($ch))))
        if ($isSeparator) {
            $token = $buffer.ToString().Trim()
            if ($token.Length -gt 0) { $segments.Add($token) }
            $null = $buffer.Clear()
            continue
        }
        [void]$buffer.Append($ch)
    }
    $finalToken = $buffer.ToString().Trim()
    if ($finalToken.Length -gt 0) { $segments.Add($finalToken) }
    return $segments.ToArray()
}

function Convert-ScalarToMatrixData {
    param([double]$Value)
    $matrix = New-ZeroMatrix -Rows 1 -Cols 1
    $matrix[0][0] = $Value
    return New-ArrayData -Data $matrix -Shape @(1, 1)
}

function Parse-JsonLikeTensorData {
    param([string]$Text, [string]$Name)
    $source = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($source)) { throw ($Name + $ui.ParseEmptyMatrixSuffix) }

    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $root = $null
    $index = 0

    while ($index -lt $source.Length) {
        $ch = $source[$index]
        if ([char]::IsWhiteSpace($ch)) {
            $index++
            continue
        }
        if ($ch -eq '[') {
            $list = New-Object 'System.Collections.Generic.List[object]'
            if ($stack.Count -gt 0) {
                ([System.Collections.Generic.List[object]]$stack.Peek()).Add($list)
            }
            $stack.Push($list)
            if ($null -eq $root) { $root = $list }
            $index++
            continue
        }
        if ($ch -eq ']') {
            if ($stack.Count -eq 0) { throw ($Name + $ui.ParseInvalidSuffix) }
            [void]$stack.Pop()
            $index++
            continue
        }
        if ($ch -eq ',') {
            $index++
            continue
        }

        $start = $index
        while ($index -lt $source.Length) {
            $current = $source[$index]
            if ($current -eq ',' -or $current -eq ']' -or [char]::IsWhiteSpace($current)) { break }
            $index++
        }
        $token = $source.Substring($start, $index - $start).Trim()
        if ([string]::IsNullOrWhiteSpace($token)) { throw ($Name + $ui.ParseNumericSuffix) }
        if ($stack.Count -eq 0) { throw ($Name + $ui.ParseInvalidSuffix) }
        $number = 0.0
        if (-not [double]::TryParse($token, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and
            -not [double]::TryParse($token, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
            throw ($Name + $ui.ParseNumericSuffix)
        }
        ([System.Collections.Generic.List[object]]$stack.Peek()).Add([double]$number)
    }

    if ($stack.Count -ne 0) { throw ($Name + $ui.ParseInvalidSuffix) }
    if ($null -eq $root -or $root.Count -eq 0) { throw ($Name + $ui.ParseEmptyMatrixSuffix) }
    return ,$root
}

function Merge-BlockMatrices {
    param($BlockRows, [string]$Name)
    if ($BlockRows.Count -eq 0) { throw ($Name + $ui.ParseEmptyMatrixSuffix) }
    $columnCount = $BlockRows[0].Count
    foreach ($row in $BlockRows) {
        if ($row.Count -ne $columnCount) { throw ($Name + $ui.ParseRowLenSuffix) }
        foreach ($cell in $row) {
            if ($cell.Shape.Count -ne 2) { throw ($Name + $ui.NeedMatrix) }
        }
    }
    $rowHeights = @()
    $colWidths = @()
    for ($r = 0; $r -lt $BlockRows.Count; $r++) {
        $height = $BlockRows[$r][0].Shape[0]
        for ($c = 1; $c -lt $columnCount; $c++) {
            if ($BlockRows[$r][$c].Shape[0] -ne $height) { throw ($Name + $ui.ParseRowLenSuffix) }
        }
        $rowHeights += $height
    }
    for ($c = 0; $c -lt $columnCount; $c++) {
        $width = $BlockRows[0][$c].Shape[1]
        for ($r = 1; $r -lt $BlockRows.Count; $r++) {
            if ($BlockRows[$r][$c].Shape[1] -ne $width) { throw ($Name + $ui.ParseRowLenSuffix) }
        }
        $colWidths += $width
    }
    $totalRows = ($rowHeights | Measure-Object -Sum).Sum
    $totalCols = ($colWidths | Measure-Object -Sum).Sum
    $result = New-ZeroMatrix -Rows $totalRows -Cols $totalCols
    $rowOffset = 0
    for ($r = 0; $r -lt $BlockRows.Count; $r++) {
        $colOffset = 0
        for ($c = 0; $c -lt $columnCount; $c++) {
            $block = $BlockRows[$r][$c].Data
            for ($i = 0; $i -lt $BlockRows[$r][$c].Shape[0]; $i++) {
                for ($j = 0; $j -lt $BlockRows[$r][$c].Shape[1]; $j++) {
                    $result[$rowOffset + $i][$colOffset + $j] = $block[$i][$j]
                }
            }
            $colOffset += $colWidths[$c]
        }
        $rowOffset += $rowHeights[$r]
    }
    return New-ArrayData -Data $result -Shape @($totalRows, $totalCols)
}

function Parse-MatrixText {
    param(
        [string]$Text,
        [string]$Name,
        [scriptblock]$ResolveMatrix = $null
    )
    $normalizedText = (Normalize-CommonSymbols $Text).Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedText)) { throw ($Name + $ui.ParseEmptySuffix) }

    $special = Try-ParseSpecialMatrixInput -Text $normalizedText -Name $Name
    if ($null -ne $special) { return $special }

    $scalarValue = 0.0
    if ([double]::TryParse($normalizedText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$scalarValue) -or
        [double]::TryParse($normalizedText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$scalarValue)) {
        return Convert-ScalarToMatrixData -Value $scalarValue
    }

    if ($null -ne $ResolveMatrix -and $normalizedText -match '^[\p{L}_][\p{L}\p{Nd}_]*$') {
        return (& $ResolveMatrix (Get-CanonicalIdentifier $normalizedText))
    }

    if ($normalizedText -match '^\[.*\]$' -and $normalizedText.Contains('|')) {
        $inner = $normalizedText.Substring(1, $normalizedText.Length - 2)
        $rowTexts = @(Split-TopLevelText -Text $inner -Separators @(';'))
        $blockRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($rowText in $rowTexts) {
            $colTexts = @(Split-TopLevelText -Text $rowText -Separators @(',', '|') -WhitespaceSeparator)
            $cells = New-Object 'System.Collections.Generic.List[object]'
            foreach ($cellText in $colTexts) {
                $cells.Add((Parse-MatrixText -Text $cellText -Name $Name -ResolveMatrix $ResolveMatrix))
            }
            $blockRows.Add($cells)
        }
        return Merge-BlockMatrices -BlockRows $blockRows -Name $Name
    }

    $trimmed = (Convert-MatlabMatrixToJsonLike $normalizedText).Trim()
    try { $data = Parse-JsonLikeTensorData -Text $trimmed -Name $Name } catch { throw $_ }
    $shape = @(Get-TensorShape -Node $data -Name $Name)
    if ($shape.Count -eq 0) { throw ($Name + $ui.ParseEmptyMatrixSuffix) }
    return New-ArrayData -Data $data -Shape $shape
}

function Parse-Matrix {
    param([string]$Text, [string]$Name, [scriptblock]$ResolveMatrix = $null)
    return Parse-MatrixText -Text $Text -Name $Name -ResolveMatrix $ResolveMatrix
}

function Clone-TensorData {
    param($Node)
    if (-not (Test-IsTensorCollection $Node)) { return [double]$Node }
    $clone = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in $Node) { $clone.Add([object](Clone-TensorData -Node $item)) }
    return ,$clone
}

function Clone-Matrix {
    param($Matrix)
    return ,(Clone-TensorData -Node $Matrix)
}

function New-ZeroMatrix {
    param([int]$Rows, [int]$Cols)
    $matrix = New-Object 'System.Collections.Generic.List[object]'
    for ($r = 0; $r -lt $Rows; $r++) {
        $row = New-Object 'System.Collections.Generic.List[double]'
        for ($c = 0; $c -lt $Cols; $c++) { $row.Add(0.0) }
        $matrix.Add($row)
    }
    return ,$matrix
}

function New-IdentityMatrix {
    param([int]$Size)
    $matrix = New-ZeroMatrix -Rows $Size -Cols $Size
    for ($i = 0; $i -lt $Size; $i++) { $matrix[$i][$i] = 1.0 }
    return ,$matrix
}

function Format-Number {
    param([double]$Value)
    if ([math]::Abs($Value) -lt 1e-10) { $Value = 0.0 }
    if ([math]::Abs($Value - [math]::Round($Value)) -lt 1e-10) { return ([math]::Round($Value)).ToString() }
    return $Value.ToString('0.######')
}

function Format-Matrix {
    param($Matrix)
    $lines = foreach ($row in $Matrix) { '[ ' + (($row | ForEach-Object { Format-Number $_ }) -join ', ') + ' ]' }
    return $lines -join [Environment]::NewLine
}

function Format-MatrixMatlab {
    param($Matrix)
    if ($null -eq $Matrix -or $Matrix.Count -eq 0) { return '[]' }
    $lines = @()
    for ($r = 0; $r -lt $Matrix.Count; $r++) {
        $rowText = ($Matrix[$r] | ForEach-Object { Format-Number $_ }) -join '  '
        if ($Matrix.Count -eq 1) {
            $lines += '[' + $rowText + ']'
        } elseif ($r -eq 0) {
            $lines += '[' + $rowText
        } elseif ($r -eq $Matrix.Count - 1) {
            $lines += ' ' + $rowText + ']'
        } else {
            $lines += ' ' + $rowText
        }
    }
    return $lines -join [Environment]::NewLine
}

function Indent-Text {
    param([string]$Text, [string]$Prefix)
    return (($Text -split "(`r`n|`n|`r)") | Where-Object { $_ -notmatch "^[`r`n]+$" } | ForEach-Object { $Prefix + $_ }) -join [Environment]::NewLine
}

function Format-TensorData {
    param($Node, [int]$Depth = 0)
    if (-not (Test-IsTensorCollection $Node)) { return Format-Number $Node }
    $items = Get-EnumerableItems $Node
    if ($items.Count -eq 0) { return '[]' }
    $allScalar = $true
    foreach ($item in $items) {
        if (Test-IsTensorCollection $item) {
            $allScalar = $false
            break
        }
    }
    if ($allScalar) {
        return '[' + (($items | ForEach-Object { Format-Number $_ }) -join ', ') + ']'
    }
    $indent = '  ' * $Depth
    $childIndent = '  ' * ($Depth + 1)
    $segments = foreach ($item in $items) {
        Indent-Text -Text (Format-TensorData -Node $item -Depth ($Depth + 1)) -Prefix $childIndent
    }
    return '[' + [Environment]::NewLine + ($segments -join (',' + [Environment]::NewLine)) + [Environment]::NewLine + $indent + ']'
}

function Format-ArrayData {
    param($ArrayData)
    if ($ArrayData.Shape.Count -eq 2) { return Format-MatrixMatlab $ArrayData.Data }
    return Format-TensorData -Node $ArrayData.Data
}

function Get-BatchShape {
    param([int[]]$Shape)
    if ($Shape.Count -le 2) { return @() }
    return @($Shape[0..($Shape.Count - 3)])
}

function Get-MatrixTailShape {
    param([int[]]$Shape)
    if ($Shape.Count -lt 2) { throw $ui.NeedMatrix }
    return @($Shape[($Shape.Count - 2)..($Shape.Count - 1)])
}

function Get-TensorNodeByIndices {
    param($Node, [int[]]$Indices)
    $current = $Node
    foreach ($index in $Indices) { $current = $current[$index] }
    return ,$current
}

function Build-BatchedTensorData {
    param([int[]]$BatchShape, [scriptblock]$LeafFactory, [int[]]$Prefix = @())
    if ($BatchShape.Count -eq 0) { return (& $LeafFactory $Prefix) }
    $result = New-Object 'System.Collections.Generic.List[object]'
    $remaining = if ($BatchShape.Count -gt 1) { @($BatchShape[1..($BatchShape.Count - 1)]) } else { @() }
    for ($i = 0; $i -lt $BatchShape[0]; $i++) {
        $result.Add((Build-BatchedTensorData -BatchShape $remaining -LeafFactory $LeafFactory -Prefix (@($Prefix + $i))))
    }
    return ,$result
}

function Invoke-BatchedMatrixUnary {
    param($ArrayData, [scriptblock]$Operation, [scriptblock]$ShapeFactory)
    $batchShape = @(Get-BatchShape $ArrayData.Shape)
    $tailShape = @(Get-MatrixTailShape $ArrayData.Shape)
    $resultShape = @(& $ShapeFactory $tailShape)
    $leafFactory = {
        param($Prefix)
        & $Operation (Get-TensorNodeByIndices -Node $ArrayData.Data -Indices $Prefix)
    }
    if ($batchShape.Count -eq 0) {
        return New-ArrayData -Data (& $leafFactory @()) -Shape $resultShape
    }
    return New-ArrayData -Data (Build-BatchedTensorData -BatchShape $batchShape -LeafFactory $leafFactory) -Shape @($batchShape + $resultShape)
}

function Invoke-BatchedMatrixBinary {
    param($LeftArray, $RightArray, [scriptblock]$Operation, [scriptblock]$ShapeFactory)
    $leftBatch = @(Get-BatchShape $LeftArray.Shape)
    $rightBatch = @(Get-BatchShape $RightArray.Shape)
    if (-not (Test-ShapeEqual -LeftShape $leftBatch -RightShape $rightBatch)) { throw $ui.BatchDim }
    $leftTail = @(Get-MatrixTailShape $LeftArray.Shape)
    $rightTail = @(Get-MatrixTailShape $RightArray.Shape)
    $resultShape = @(& $ShapeFactory $leftTail $rightTail)
    $leafFactory = {
        param($Prefix)
        & $Operation (Get-TensorNodeByIndices -Node $LeftArray.Data -Indices $Prefix) (Get-TensorNodeByIndices -Node $RightArray.Data -Indices $Prefix)
    }
    if ($leftBatch.Count -eq 0) {
        return New-ArrayData -Data (& $leafFactory @()) -Shape $resultShape
    }
    return New-ArrayData -Data (Build-BatchedTensorData -BatchShape $leftBatch -LeafFactory $leafFactory) -Shape @($leftBatch + $resultShape)
}

function Invoke-BatchedScalarUnary {
    param($ArrayData, [scriptblock]$Operation)
    $batchShape = @(Get-BatchShape $ArrayData.Shape)
    $leafFactory = {
        param($Prefix)
        [double](& $Operation (Get-TensorNodeByIndices -Node $ArrayData.Data -Indices $Prefix))
    }
    if ($batchShape.Count -eq 0) { return New-ScalarValue (& $leafFactory @()) }
    return New-ArrayValue (New-ArrayData -Data (Build-BatchedTensorData -BatchShape $batchShape -LeafFactory $leafFactory) -Shape $batchShape)
}

function Get-BatchIndexText {
    param([int[]]$Indices)
    if ($Indices.Count -eq 0) { return '' }
    return '[' + (($Indices | ForEach-Object { $_.ToString() }) -join ',') + ']'
}

function New-ScalarValue { param([double]$Value) return @{ Kind = 'Scalar'; Value = [double]$Value } }
function New-ArrayValue { param($Value) return @{ Kind = 'Array'; Value = $Value } }
function New-MatrixValue { param($Value) return (New-ArrayValue $Value) }
function New-TextValue { param([string]$Value) return @{ Kind = 'Text'; Value = $Value } }
function Assert-ArrayValue { param($Value) if ($Value.Kind -ne 'Array') { throw $ui.NeedMatrix } }
function Assert-MatrixValue {
    param($Value)
    Assert-ArrayValue $Value
    if ($Value.Value.Shape.Count -lt 2) { throw $ui.NeedMatrix }
}

function Assert-SameArrayShape {
    param($Left, $Right, [string]$Message)
    if (-not (Test-ShapeEqual -LeftShape $Left.Shape -RightShape $Right.Shape)) {
        throw ($Message + ' (' + (($Left.Shape | ForEach-Object { $_.ToString() }) -join 'x') + ' vs ' + (($Right.Shape | ForEach-Object { $_.ToString() }) -join 'x') + ')')
    }
}

function Invoke-ArrayUnaryOperation {
    param($Node, [scriptblock]$Operation)
    if (-not (Test-IsTensorCollection $Node)) { return (& $Operation ([double]$Node)) }
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in $Node) { $result.Add((Invoke-ArrayUnaryOperation -Node $item -Operation $Operation)) }
    return ,$result
}

function Invoke-ArrayBinaryOperation {
    param($LeftNode, $RightNode, [scriptblock]$Operation)
    if (-not (Test-IsTensorCollection $LeftNode)) { return (& $Operation ([double]$LeftNode) ([double]$RightNode)) }
    $result = New-Object 'System.Collections.Generic.List[object]'
    for ($i = 0; $i -lt $LeftNode.Count; $i++) {
        $result.Add((Invoke-ArrayBinaryOperation -LeftNode $LeftNode[$i] -RightNode $RightNode[$i] -Operation $Operation))
    }
    return ,$result
}

function Add-Matrices {
    param($A, $B)
    return Invoke-ArrayBinaryOperation -LeftNode $A -RightNode $B -Operation { param($LeftValue, $RightValue) $LeftValue + $RightValue }
}

function Subtract-Matrices {
    param($A, $B)
    return Invoke-ArrayBinaryOperation -LeftNode $A -RightNode $B -Operation { param($LeftValue, $RightValue) $LeftValue - $RightValue }
}

function Multiply-Matrices {
    param($A, $B)
    if ($A[0].Count -ne $B.Count) {
        throw ($ui.MulDim + ' (' + $A.Count + 'x' + $A[0].Count + ' * ' + $B.Count + 'x' + $B[0].Count + ')')
    }
    $result = New-ZeroMatrix -Rows $A.Count -Cols $B[0].Count
    for ($r = 0; $r -lt $A.Count; $r++) {
        for ($c = 0; $c -lt $B[0].Count; $c++) {
            $sum = 0.0
            for ($k = 0; $k -lt $B.Count; $k++) { $sum += $A[$r][$k] * $B[$k][$c] }
            $result[$r][$c] = $sum
        }
    }
    return ,$result
}

function Multiply-MatrixByScalar {
    param($Matrix, [double]$Scalar)
    return Invoke-ArrayUnaryOperation -Node $Matrix -Operation { param($Value) $Value * $Scalar }
}

function Transpose-Matrix {
    param($Matrix)
    $result = New-ZeroMatrix -Rows $Matrix[0].Count -Cols $Matrix.Count
    for ($r = 0; $r -lt $Matrix.Count; $r++) { for ($c = 0; $c -lt $Matrix[0].Count; $c++) { $result[$c][$r] = $Matrix[$r][$c] } }
    return ,$result
}

function Get-Rref {
    param($Matrix)
    $working = Clone-Matrix $Matrix
    $rows = $working.Count
    $cols = $working[0].Count
    $lead = 0
    $rank = 0
    $pivotCols = New-Object 'System.Collections.Generic.List[int]'
    for ($r = 0; $r -lt $rows; $r++) {
        if ($lead -ge $cols) { break }
        $pivotRow = $r
        while ($pivotRow -lt $rows -and [math]::Abs($working[$pivotRow][$lead]) -lt 1e-10) { $pivotRow++ }
        while ($pivotRow -ge $rows) {
            $lead++
            if ($lead -ge $cols) { return @{ Matrix = $working; Rank = $rank; PivotCols = $pivotCols } }
            $pivotRow = $r
            while ($pivotRow -lt $rows -and [math]::Abs($working[$pivotRow][$lead]) -lt 1e-10) { $pivotRow++ }
        }
        if ($pivotRow -ne $r) {
            $tmp = $working[$r]
            $working[$r] = $working[$pivotRow]
            $working[$pivotRow] = $tmp
        }
        $pivot = $working[$r][$lead]
        for ($c = 0; $c -lt $cols; $c++) { $working[$r][$c] = $working[$r][$c] / $pivot }
        for ($row = 0; $row -lt $rows; $row++) {
            if ($row -eq $r) { continue }
            $factor = $working[$row][$lead]
            if ([math]::Abs($factor) -lt 1e-10) { continue }
            for ($c = 0; $c -lt $cols; $c++) {
                $working[$row][$c] = $working[$row][$c] - ($factor * $working[$r][$c])
                if ([math]::Abs($working[$row][$c]) -lt 1e-10) { $working[$row][$c] = 0.0 }
            }
        }
        $pivotCols.Add($lead)
        $rank++
        $lead++
    }
    return @{ Matrix = $working; Rank = $rank; PivotCols = $pivotCols }
}

function Get-Ref {
    param($Matrix)
    $working = Clone-Matrix $Matrix
    $rows = $working.Count
    $cols = $working[0].Count
    $lead = 0
    $rank = 0
    for ($r = 0; $r -lt $rows; $r++) {
        if ($lead -ge $cols) { break }
        $pivotRow = $r
        while ($pivotRow -lt $rows -and [math]::Abs($working[$pivotRow][$lead]) -lt 1e-10) { $pivotRow++ }
        while ($pivotRow -ge $rows) {
            $lead++
            if ($lead -ge $cols) { return @{ Matrix = $working; Rank = $rank } }
            $pivotRow = $r
            while ($pivotRow -lt $rows -and [math]::Abs($working[$pivotRow][$lead]) -lt 1e-10) { $pivotRow++ }
        }
        if ($pivotRow -ne $r) {
            $tmp = $working[$r]
            $working[$r] = $working[$pivotRow]
            $working[$pivotRow] = $tmp
        }
        $pivot = $working[$r][$lead]
        if ([math]::Abs($pivot) -gt 1e-10) {
            for ($c = $lead; $c -lt $cols; $c++) { $working[$r][$c] = $working[$r][$c] / $pivot }
        }
        for ($row = $r + 1; $row -lt $rows; $row++) {
            $factor = $working[$row][$lead]
            if ([math]::Abs($factor) -lt 1e-10) { continue }
            for ($c = $lead; $c -lt $cols; $c++) {
                $working[$row][$c] = $working[$row][$c] - ($factor * $working[$r][$c])
                if ([math]::Abs($working[$row][$c]) -lt 1e-10) { $working[$row][$c] = 0.0 }
            }
        }
        $rank++
        $lead++
    }
    return @{ Matrix = $working; Rank = $rank }
}

function Get-Determinant {
    param($Matrix)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw $ui.DetSquare }
    $working = Clone-Matrix $Matrix
    $size = $working.Count
    $swapCount = 0
    for ($col = 0; $col -lt $size; $col++) {
        $pivotRow = $col
        while ($pivotRow -lt $size -and [math]::Abs($working[$pivotRow][$col]) -lt 1e-10) { $pivotRow++ }
        if ($pivotRow -ge $size) { return 0.0 }
        if ($pivotRow -ne $col) {
            $tmp = $working[$col]
            $working[$col] = $working[$pivotRow]
            $working[$pivotRow] = $tmp
            $swapCount++
        }
        for ($row = $col + 1; $row -lt $size; $row++) {
            $factor = $working[$row][$col] / $working[$col][$col]
            for ($c = $col; $c -lt $size; $c++) {
                $working[$row][$c] = $working[$row][$c] - ($factor * $working[$col][$c])
                if ([math]::Abs($working[$row][$c]) -lt 1e-10) { $working[$row][$c] = 0.0 }
            }
        }
    }
    $det = if ($swapCount % 2 -eq 0) { 1.0 } else { -1.0 }
    for ($i = 0; $i -lt $size; $i++) { $det *= $working[$i][$i] }
    return $det
}

function Get-Trace {
    param($Matrix)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw $ui.SquareOnly }
    $sum = 0.0
    for ($i = 0; $i -lt $Matrix.Count; $i++) { $sum += $Matrix[$i][$i] }
    return $sum
}

function Get-FrobeniusNorm {
    param($Matrix)
    $sum = 0.0
    if (-not (Test-IsTensorCollection $Matrix)) { return [math]::Abs([double]$Matrix) }
    foreach ($item in $Matrix) { $sum += [math]::Pow((Get-FrobeniusNorm $item), 2) }
    return [math]::Sqrt($sum)
}

function Get-MinorMatrix {
    param($Matrix, [int]$RemoveRow, [int]$RemoveCol)
    $minor = New-Object 'System.Collections.Generic.List[object]'
    for ($r = 0; $r -lt $Matrix.Count; $r++) {
        if ($r -eq $RemoveRow) { continue }
        $row = New-Object 'System.Collections.Generic.List[double]'
        for ($c = 0; $c -lt $Matrix[0].Count; $c++) {
            if ($c -eq $RemoveCol) { continue }
            $row.Add($Matrix[$r][$c])
        }
        $minor.Add($row)
    }
    return ,$minor
}

function Get-CofactorMatrix {
    param($Matrix)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw $ui.SquareOnly }
    if ($Matrix.Count -eq 1) {
        $single = New-ZeroMatrix -Rows 1 -Cols 1
        $single[0][0] = 1.0
        return ,$single
    }
    $size = $Matrix.Count
    $result = New-ZeroMatrix -Rows $size -Cols $size
    for ($r = 0; $r -lt $size; $r++) {
        for ($c = 0; $c -lt $size; $c++) {
            $minor = Get-MinorMatrix -Matrix $Matrix -RemoveRow $r -RemoveCol $c
            $sign = if ((($r + $c) % 2) -eq 0) { 1.0 } else { -1.0 }
            $result[$r][$c] = $sign * (Get-Determinant $minor)
        }
    }
    return ,$result
}

function Get-AdjugateMatrix {
    param($Matrix)
    return Transpose-Matrix (Get-CofactorMatrix $Matrix)
}

function Get-InverseMatrix {
    param($Matrix)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw $ui.InvSquare }
    $size = $Matrix.Count
    $left = Clone-Matrix $Matrix
    $right = New-IdentityMatrix -Size $size
    for ($col = 0; $col -lt $size; $col++) {
        $pivotRow = $col
        while ($pivotRow -lt $size -and [math]::Abs($left[$pivotRow][$col]) -lt 1e-10) { $pivotRow++ }
        if ($pivotRow -ge $size) { throw $ui.NotInvertible }
        if ($pivotRow -ne $col) {
            $tmp = $left[$col]; $left[$col] = $left[$pivotRow]; $left[$pivotRow] = $tmp
            $tmp = $right[$col]; $right[$col] = $right[$pivotRow]; $right[$pivotRow] = $tmp
        }
        $pivot = $left[$col][$col]
        for ($c = 0; $c -lt $size; $c++) {
            $left[$col][$c] = $left[$col][$c] / $pivot
            $right[$col][$c] = $right[$col][$c] / $pivot
        }
        for ($row = 0; $row -lt $size; $row++) {
            if ($row -eq $col) { continue }
            $factor = $left[$row][$col]
            if ([math]::Abs($factor) -lt 1e-10) { continue }
            for ($c = 0; $c -lt $size; $c++) {
                $left[$row][$c] = $left[$row][$c] - ($factor * $left[$col][$c])
                $right[$row][$c] = $right[$row][$c] - ($factor * $right[$col][$c])
                if ([math]::Abs($left[$row][$c]) -lt 1e-10) { $left[$row][$c] = 0.0 }
                if ([math]::Abs($right[$row][$c]) -lt 1e-10) { $right[$row][$c] = 0.0 }
            }
        }
    }
    return ,$right
}

function Get-NullSpaceValue {
    param($Matrix)
    $rrefData = Get-Rref $Matrix
    $rref = $rrefData.Matrix
    $cols = $Matrix[0].Count
    $pivotCols = @($rrefData.PivotCols)
    $freeCols = @()
    for ($c = 0; $c -lt $cols; $c++) { if ($pivotCols -notcontains $c) { $freeCols += $c } }
    if ($freeCols.Count -eq 0) { return New-TextValue -Value $ui.TrivialNull }
    $basis = New-ZeroMatrix -Rows $cols -Cols $freeCols.Count
    for ($freeIndex = 0; $freeIndex -lt $freeCols.Count; $freeIndex++) {
        $freeCol = $freeCols[$freeIndex]
        $basis[$freeCol][$freeIndex] = 1.0
        for ($row = 0; $row -lt $pivotCols.Count; $row++) {
            $pivotCol = $pivotCols[$row]
            $basis[$pivotCol][$freeIndex] = -1.0 * $rref[$row][$freeCol]
        }
    }
    return New-MatrixValue -Value (New-ArrayData -Data $basis -Shape @($cols, $freeCols.Count))
}

function Get-RowSumsVector {
    param($Matrix)
    $result = New-Object 'System.Collections.Generic.List[double]'
    for ($r = 0; $r -lt $Matrix.Count; $r++) {
        $sum = 0.0
        for ($c = 0; $c -lt $Matrix[$r].Count; $c++) { $sum += [double]$Matrix[$r][$c] }
        $result.Add($sum)
    }
    return ,$result
}

function Get-ColSumsVector {
    param($Matrix)
    $result = New-Object 'System.Collections.Generic.List[double]'
    for ($c = 0; $c -lt $Matrix[0].Count; $c++) {
        $sum = 0.0
        for ($r = 0; $r -lt $Matrix.Count; $r++) { $sum += [double]$Matrix[$r][$c] }
        $result.Add($sum)
    }
    return ,$result
}

function Get-RowMeansVector {
    param($Matrix)
    $sums = Get-RowSumsVector $Matrix
    $result = New-Object 'System.Collections.Generic.List[double]'
    foreach ($sum in $sums) { $result.Add(([double]$sum / [double]$Matrix[0].Count)) }
    return ,$result
}

function Get-ColMeansVector {
    param($Matrix)
    $sums = Get-ColSumsVector $Matrix
    $result = New-Object 'System.Collections.Generic.List[double]'
    foreach ($sum in $sums) { $result.Add(([double]$sum / [double]$Matrix.Count)) }
    return ,$result
}

function Get-RowMaxVector {
    param($Matrix)
    $result = New-Object 'System.Collections.Generic.List[double]'
    foreach ($row in $Matrix) {
        $result.Add(([double[]](Get-EnumerableItems $row) | Measure-Object -Maximum).Maximum)
    }
    return ,$result
}

function Get-ColMaxVector {
    param($Matrix)
    $result = New-Object 'System.Collections.Generic.List[double]'
    for ($c = 0; $c -lt $Matrix[0].Count; $c++) {
        $columnValues = @()
        foreach ($row in $Matrix) { $columnValues += [double]$row[$c] }
        $result.Add(([double[]]$columnValues | Measure-Object -Maximum).Maximum)
    }
    return ,$result
}

function Get-RowMinVector {
    param($Matrix)
    $result = New-Object 'System.Collections.Generic.List[double]'
    foreach ($row in $Matrix) {
        $result.Add(([double[]](Get-EnumerableItems $row) | Measure-Object -Minimum).Minimum)
    }
    return ,$result
}

function Get-ColMinVector {
    param($Matrix)
    $result = New-Object 'System.Collections.Generic.List[double]'
    for ($c = 0; $c -lt $Matrix[0].Count; $c++) {
        $columnValues = @()
        foreach ($row in $Matrix) { $columnValues += [double]$row[$c] }
        $result.Add(([double[]]$columnValues | Measure-Object -Minimum).Minimum)
    }
    return ,$result
}

function Get-MatrixPowerData {
    param($Matrix, [double]$Power)
    if ([math]::Abs($Power - [math]::Round($Power)) -gt 1e-10) { throw $ui.IntegerPowerOnly }
    $n = [int][math]::Round($Power)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw $ui.SquareOnly }
    if ($n -eq 0) { return New-IdentityMatrix -Size $Matrix.Count }
    $baseMatrix = if ($n -lt 0) { Get-InverseMatrix $Matrix } else { Clone-Matrix $Matrix }
    $exp = [math]::Abs($n)
    $result = New-IdentityMatrix -Size $Matrix.Count
    while ($exp -gt 0) {
        if (($exp % 2) -eq 1) { $result = Multiply-Matrices $result $baseMatrix }
        $exp = [math]::Floor($exp / 2)
        if ($exp -gt 0) { $baseMatrix = Multiply-Matrices $baseMatrix $baseMatrix }
    }
    return ,$result
}

function Test-IsSymmetricMatrix {
    param($Matrix, [double]$Tolerance = 1e-8)
    if ($Matrix.Count -ne $Matrix[0].Count) { return $false }
    for ($r = 0; $r -lt $Matrix.Count; $r++) {
        for ($c = $r + 1; $c -lt $Matrix[0].Count; $c++) {
            if ([math]::Abs([double]$Matrix[$r][$c] - [double]$Matrix[$c][$r]) -gt $Tolerance) { return $false }
        }
    }
    return $true
}

function Get-ColumnVector {
    param($Matrix, [int]$ColumnIndex)
    $result = New-Object 'System.Collections.Generic.List[double]'
    for ($r = 0; $r -lt $Matrix.Count; $r++) { $result.Add([double]$Matrix[$r][$ColumnIndex]) }
    return ,$result
}

function Get-VectorDotProduct {
    param([double[]]$Left, [double[]]$Right)
    $sum = 0.0
    for ($i = 0; $i -lt $Left.Count; $i++) { $sum += ([double]$Left[$i] * [double]$Right[$i]) }
    return $sum
}

function Get-VectorNormValue {
    param([double[]]$Vector)
    return [math]::Sqrt((Get-VectorDotProduct -Left $Vector -Right $Vector))
}

function Sort-EigenDecompositionData {
    param([double[]]$Values, $Vectors)
    $order = @()
    for ($i = 0; $i -lt $Values.Count; $i++) { $order += [pscustomobject]@{ Index = $i; Value = [double]$Values[$i] } }
    $order = @($order | Sort-Object -Property Value -Descending)
    $sortedValues = New-Object 'System.Collections.Generic.List[double]'
    $sortedVectors = New-ZeroMatrix -Rows $Vectors.Count -Cols $Vectors[0].Count
    for ($newCol = 0; $newCol -lt $order.Count; $newCol++) {
        $srcCol = $order[$newCol].Index
        $sortedValues.Add([double]$order[$newCol].Value)
        for ($row = 0; $row -lt $Vectors.Count; $row++) { $sortedVectors[$row][$newCol] = [double]$Vectors[$row][$srcCol] }
    }
    return [pscustomobject]@{ EigenValues = $sortedValues; EigenVectors = $sortedVectors }
}

function Get-SymmetricEigenDecomposition {
    param($Matrix, [double]$Tolerance = 1e-10, [int]$MaxIterations = 128)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw 'EIGVALS/EIGVECS 需要方阵。' }
    if (-not (Test-IsSymmetricMatrix -Matrix $Matrix -Tolerance 1e-8)) { throw '当前版本的 EIGVALS / EIGVECS 稳定支持实对称矩阵。' }
    $n = $Matrix.Count
    $a = Clone-Matrix $Matrix
    $vectors = New-IdentityMatrix -Size $n
    for ($iteration = 0; $iteration -lt $MaxIterations; $iteration++) {
        $maxValue = 0.0
        $p = 0
        $q = 1
        for ($r = 0; $r -lt $n; $r++) {
            for ($c = $r + 1; $c -lt $n; $c++) {
                $value = [math]::Abs([double]$a[$r][$c])
                if ($value -gt $maxValue) { $maxValue = $value; $p = $r; $q = $c }
            }
        }
        if ($maxValue -lt $Tolerance) { break }
        $app = [double]$a[$p][$p]
        $aqq = [double]$a[$q][$q]
        $apq = [double]$a[$p][$q]
        $angle = if ([math]::Abs($aqq - $app) -lt $Tolerance) { [math]::PI / 4.0 } else { 0.5 * [math]::Atan2((2.0 * $apq), ($aqq - $app)) }
        $cosValue = [math]::Cos($angle)
        $sinValue = [math]::Sin($angle)
        for ($r = 0; $r -lt $n; $r++) {
            if ($r -eq $p -or $r -eq $q) { continue }
            $arp = [double]$a[$r][$p]
            $arq = [double]$a[$r][$q]
            $a[$r][$p] = ($cosValue * $arp) - ($sinValue * $arq)
            $a[$p][$r] = [double]$a[$r][$p]
            $a[$r][$q] = ($sinValue * $arp) + ($cosValue * $arq)
            $a[$q][$r] = [double]$a[$r][$q]
        }
        $a[$p][$p] = ($cosValue * $cosValue * $app) - (2.0 * $sinValue * $cosValue * $apq) + ($sinValue * $sinValue * $aqq)
        $a[$q][$q] = ($sinValue * $sinValue * $app) + (2.0 * $sinValue * $cosValue * $apq) + ($cosValue * $cosValue * $aqq)
        $a[$p][$q] = 0.0
        $a[$q][$p] = 0.0
        for ($r = 0; $r -lt $n; $r++) {
            $vrp = [double]$vectors[$r][$p]
            $vrq = [double]$vectors[$r][$q]
            $vectors[$r][$p] = ($cosValue * $vrp) - ($sinValue * $vrq)
            $vectors[$r][$q] = ($sinValue * $vrp) + ($cosValue * $vrq)
        }
    }
    $values = New-Object 'System.Collections.Generic.List[double]'
    for ($i = 0; $i -lt $n; $i++) { $values.Add([double]$a[$i][$i]) }
    return Sort-EigenDecompositionData -Values $values.ToArray() -Vectors $vectors
}

function Get-EigenValuesVector {
    param($Matrix)
    $eigen = Get-SymmetricEigenDecomposition $Matrix
    return $eigen.EigenValues
}

function Get-EigenVectorsMatrix {
    param($Matrix)
    $eigen = Get-SymmetricEigenDecomposition $Matrix
    return $eigen.EigenVectors
}

function Get-SingularValuesVector {
    param($Matrix)
    $gram = Multiply-Matrices (Transpose-Matrix $Matrix) $Matrix
    $eigen = Get-SymmetricEigenDecomposition $gram
    $limit = [Math]::Min($Matrix.Count, $Matrix[0].Count)
    $values = New-Object 'System.Collections.Generic.List[double]'
    for ($i = 0; $i -lt $limit; $i++) {
        $raw = if ($i -lt $eigen.EigenValues.Count) { [double]$eigen.EigenValues[$i] } else { 0.0 }
        if ($raw -lt 0.0 -and [math]::Abs($raw) -lt 1e-8) { $raw = 0.0 }
        if ($raw -lt 0.0) { throw 'SVALS 计算失败：A^T A 的特征值出现负值。' }
        $values.Add([math]::Sqrt($raw))
    }
    return ,$values
}

function Get-CholeskyMatrix {
    param($Matrix)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw 'CHOL 需要方阵。' }
    if (-not (Test-IsSymmetricMatrix -Matrix $Matrix -Tolerance 1e-8)) { throw 'CHOL 需要对称正定矩阵。' }
    $n = $Matrix.Count
    $result = New-ZeroMatrix -Rows $n -Cols $n
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -le $i; $j++) {
            $sum = 0.0
            for ($k = 0; $k -lt $j; $k++) { $sum += ([double]$result[$i][$k] * [double]$result[$j][$k]) }
            if ($i -eq $j) {
                $value = [double]$Matrix[$i][$i] - $sum
                if ($value -le 1e-10) { throw 'CHOL 需要对称正定矩阵。' }
                $result[$i][$j] = [math]::Sqrt($value)
            } else {
                if ([math]::Abs([double]$result[$j][$j]) -lt 1e-10) { throw 'CHOL 计算失败：遇到零主元。' }
                $result[$i][$j] = ([double]$Matrix[$i][$j] - $sum) / [double]$result[$j][$j]
            }
        }
    }
    return ,$result
}

function Get-LuDecompositionText {
    param($Matrix)
    if ($Matrix.Count -ne $Matrix[0].Count) { throw 'LU 需要方阵。' }
    $n = $Matrix.Count
    $u = Clone-Matrix $Matrix
    $l = New-ZeroMatrix -Rows $n -Cols $n
    $p = New-IdentityMatrix -Size $n
    for ($i = 0; $i -lt $n; $i++) {
        $pivotRow = $i
        $pivotValue = [math]::Abs([double]$u[$i][$i])
        for ($row = $i + 1; $row -lt $n; $row++) {
            $candidate = [math]::Abs([double]$u[$row][$i])
            if ($candidate -gt $pivotValue) { $pivotValue = $candidate; $pivotRow = $row }
        }
        if ($pivotValue -lt 1e-10) { throw 'LU 分解失败：遇到零主元。' }
        if ($pivotRow -ne $i) {
            $tmp = $u[$i]; $u[$i] = $u[$pivotRow]; $u[$pivotRow] = $tmp
            $tmp = $p[$i]; $p[$i] = $p[$pivotRow]; $p[$pivotRow] = $tmp
            for ($k = 0; $k -lt $i; $k++) {
                $tempValue = [double]$l[$i][$k]
                $l[$i][$k] = [double]$l[$pivotRow][$k]
                $l[$pivotRow][$k] = $tempValue
            }
        }
        $l[$i][$i] = 1.0
        for ($row = $i + 1; $row -lt $n; $row++) {
            $factor = [double]$u[$row][$i] / [double]$u[$i][$i]
            $l[$row][$i] = $factor
            for ($col = $i; $col -lt $n; $col++) {
                $u[$row][$col] = [double]$u[$row][$col] - ($factor * [double]$u[$i][$col])
                if ([math]::Abs([double]$u[$row][$col]) -lt 1e-10) { $u[$row][$col] = 0.0 }
            }
        }
    }
    return ('P =' + [Environment]::NewLine + (Format-Matrix $p) + [Environment]::NewLine + [Environment]::NewLine + 'L =' + [Environment]::NewLine + (Format-Matrix $l) + [Environment]::NewLine + [Environment]::NewLine + 'U =' + [Environment]::NewLine + (Format-Matrix $u))
}

function Get-QrDecompositionText {
    param($Matrix)
    $rows = $Matrix.Count
    $cols = $Matrix[0].Count
    $q = New-ZeroMatrix -Rows $rows -Cols $cols
    $r = New-ZeroMatrix -Rows $cols -Cols $cols
    for ($j = 0; $j -lt $cols; $j++) {
        $sourceColumn = @((Get-ColumnVector -Matrix $Matrix -ColumnIndex $j) | ForEach-Object { [double]$_ })
        $v = @($sourceColumn)
        for ($i = 0; $i -lt $j; $i++) {
            $qColumn = @((Get-ColumnVector -Matrix $q -ColumnIndex $i) | ForEach-Object { [double]$_ })
            $rij = Get-VectorDotProduct -Left $qColumn -Right $sourceColumn
            $r[$i][$j] = $rij
            for ($k = 0; $k -lt $rows; $k++) { $v[$k] = [double]$v[$k] - ($rij * [double]$qColumn[$k]) }
        }
        $norm = Get-VectorNormValue -Vector $v
        if ($norm -lt 1e-10) { throw 'QR 分解失败：检测到线性相关列。' }
        $r[$j][$j] = $norm
        for ($k = 0; $k -lt $rows; $k++) { $q[$k][$j] = [double]$v[$k] / $norm }
    }
    return ('Q =' + [Environment]::NewLine + (Format-Matrix $q) + [Environment]::NewLine + [Environment]::NewLine + 'R =' + [Environment]::NewLine + (Format-Matrix $r))
}

function Invoke-BatchedTextUnary {
    param($ArrayData, [scriptblock]$Operation)
    $batchShape = @(Get-BatchShape $ArrayData.Shape)
    if ($batchShape.Count -eq 0) { return New-TextValue (& $Operation $ArrayData.Data) }
    $parts = New-Object 'System.Collections.Generic.List[string]'
    $leafFactory = {
        param($Prefix)
        $slice = Get-TensorNodeByIndices -Node $ArrayData.Data -Indices $Prefix
        $parts.Add((Get-BatchIndexText -Indices $Prefix) + ':' + [Environment]::NewLine + (& $Operation $slice))
        return $null
    }
    [void](Build-BatchedTensorData -BatchShape $batchShape -LeafFactory $leafFactory)
    return New-TextValue (($parts -join ([Environment]::NewLine + [Environment]::NewLine)))
}
function Add-Values {
    param($Left, $Right)
    if ($Left.Kind -eq 'Scalar' -and $Right.Kind -eq 'Scalar') { return New-ScalarValue ($Left.Value + $Right.Value) }
    if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Array') {
        Assert-SameArrayShape -Left $Left.Value -Right $Right.Value -Message $ui.SameShape
        return New-MatrixValue (New-ArrayData -Data (Add-Matrices $Left.Value.Data $Right.Value.Data) -Shape $Left.Value.Shape)
    }
    throw $ui.UnsupportedOp
}

function Subtract-Values {
    param($Left, $Right)
    if ($Left.Kind -eq 'Scalar' -and $Right.Kind -eq 'Scalar') { return New-ScalarValue ($Left.Value - $Right.Value) }
    if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Array') {
        Assert-SameArrayShape -Left $Left.Value -Right $Right.Value -Message $ui.SameShape
        return New-MatrixValue (New-ArrayData -Data (Subtract-Matrices $Left.Value.Data $Right.Value.Data) -Shape $Left.Value.Shape)
    }
    throw $ui.UnsupportedOp
}

function Multiply-Values {
    param($Left, $Right, [switch]$ElementWise)
    if ($Left.Kind -eq 'Scalar' -and $Right.Kind -eq 'Scalar') { return New-ScalarValue ($Left.Value * $Right.Value) }
    if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Array') {
        if (-not $ElementWise -and $Left.Value.Shape.Count -ge 2 -and $Right.Value.Shape.Count -ge 2) {
            return New-MatrixValue (Invoke-BatchedMatrixBinary -LeftArray $Left.Value -RightArray $Right.Value -Operation { param($LeftMatrix, $RightMatrix) Multiply-Matrices $LeftMatrix $RightMatrix } -ShapeFactory { param($LeftShape, $RightShape) @($LeftShape[0], $RightShape[1]) })
        }
        Assert-SameArrayShape -Left $Left.Value -Right $Right.Value -Message $ui.SameShape
        return New-MatrixValue (New-ArrayData -Data (Invoke-ArrayBinaryOperation -LeftNode $Left.Value.Data -RightNode $Right.Value.Data -Operation { param($LeftValue, $RightValue) $LeftValue * $RightValue }) -Shape $Left.Value.Shape)
    }
    if ($Left.Kind -eq 'Scalar' -and $Right.Kind -eq 'Array') { return New-MatrixValue (New-ArrayData -Data (Multiply-MatrixByScalar $Right.Value.Data $Left.Value) -Shape $Right.Value.Shape) }
    if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Scalar') { return New-MatrixValue (New-ArrayData -Data (Multiply-MatrixByScalar $Left.Value.Data $Right.Value) -Shape $Left.Value.Shape) }
    throw $ui.UnsupportedOp
}

function Divide-Values {
    param($Left, $Right, [switch]$ElementWise)
    if ($ElementWise) {
        if ($Left.Kind -eq 'Scalar' -and $Right.Kind -eq 'Scalar') {
            if ([math]::Abs($Right.Value) -lt 1e-10) { throw $ui.DivZero }
            return New-ScalarValue ($Left.Value / $Right.Value)
        }
        if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Scalar') {
            if ([math]::Abs($Right.Value) -lt 1e-10) { throw $ui.DivZero }
            return New-MatrixValue (New-ArrayData -Data (Multiply-MatrixByScalar $Left.Value.Data (1.0 / $Right.Value)) -Shape $Left.Value.Shape)
        }
        if ($Left.Kind -eq 'Scalar' -and $Right.Kind -eq 'Array') {
            $data = Invoke-ArrayUnaryOperation -Node $Right.Value.Data -Operation {
                param($Value)
                if ([math]::Abs($Value) -lt 1e-10) { throw $ui.DivZero }
                return $Left.Value / $Value
            }
            return New-MatrixValue (New-ArrayData -Data $data -Shape $Right.Value.Shape)
        }
        if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Array') {
            Assert-SameArrayShape -Left $Left.Value -Right $Right.Value -Message $ui.SameShape
            $data = Invoke-ArrayBinaryOperation -LeftNode $Left.Value.Data -RightNode $Right.Value.Data -Operation {
                param($LeftValue, $RightValue)
                if ([math]::Abs($RightValue) -lt 1e-10) { throw $ui.DivZero }
                return $LeftValue / $RightValue
            }
            return New-MatrixValue (New-ArrayData -Data $data -Shape $Left.Value.Shape)
        }
        throw $ui.UnsupportedOp
    }
    if ($Right.Kind -eq 'Scalar') {
        if ([math]::Abs($Right.Value) -lt 1e-10) { throw $ui.DivZero }
        if ($Left.Kind -eq 'Scalar') { return New-ScalarValue ($Left.Value / $Right.Value) }
        if ($Left.Kind -eq 'Array') { return New-MatrixValue (New-ArrayData -Data (Multiply-MatrixByScalar $Left.Value.Data (1.0 / $Right.Value)) -Shape $Left.Value.Shape) }
    }
    if ($Left.Kind -eq 'Array' -and $Right.Kind -eq 'Array' -and $Left.Value.Shape.Count -ge 2 -and $Right.Value.Shape.Count -ge 2) {
        return New-MatrixValue (Invoke-BatchedMatrixBinary -LeftArray $Left.Value -RightArray $Right.Value -Operation { param($LeftMatrix, $RightMatrix) Multiply-Matrices $LeftMatrix (Get-InverseMatrix $RightMatrix) } -ShapeFactory { param($LeftShape, $RightShape) @($LeftShape[0], $RightShape[1]) })
    }
    throw $ui.UnsupportedOp
}

function Power-Values {
    param($Left, $Right)
    if ($Right.Kind -ne 'Scalar') { throw $ui.IntegerPowerOnly }
    $power = $Right.Value
    if ($Left.Kind -eq 'Scalar') { return New-ScalarValue ([math]::Pow($Left.Value, $power)) }
    if ($Left.Kind -ne 'Array') { throw $ui.UnsupportedOp }
    if ($Left.Value.Shape.Count -lt 2) {
        return New-MatrixValue (New-ArrayData -Data (Invoke-ArrayUnaryOperation -Node $Left.Value.Data -Operation { param($Value) [math]::Pow($Value, $power) }) -Shape $Left.Value.Shape)
    }
    return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Left.Value -Operation { param($Matrix) Get-MatrixPowerData -Matrix $Matrix -Power $power } -ShapeFactory { param($InputShape) $InputShape })
}

function Invoke-FunctionValue {
    param([string]$Name, $Argument)
    switch (Get-CanonicalIdentifier $Name) {
        'T' { Assert-MatrixValue $Argument; return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Transpose-Matrix $Matrix } -ShapeFactory { param($InputShape) @($InputShape[1], $InputShape[0]) }) }
        'REF' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) (Get-Ref $Matrix).Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'RREF' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) (Get-Rref $Matrix).Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'RANK' { Assert-MatrixValue $Argument; return Invoke-BatchedScalarUnary -ArrayData $Argument.Value -Operation { param($Matrix) (Get-Rref $Matrix).Rank } }
        'DET' { Assert-MatrixValue $Argument; return Invoke-BatchedScalarUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-Determinant $Matrix } }
        'TR' { Assert-MatrixValue $Argument; return Invoke-BatchedScalarUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-Trace $Matrix } }
        'TRACE' { Assert-MatrixValue $Argument; return Invoke-BatchedScalarUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-Trace $Matrix } }
        'NORM' { Assert-ArrayValue $Argument; return New-ScalarValue (Get-FrobeniusNorm $Argument.Value.Data) }
        'ROWSUMS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-RowSumsVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[0]) })
        }
        'COLSUMS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-ColSumsVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[1]) })
        }
        'ROWMEANS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-RowMeansVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[0]) })
        }
        'COLMEANS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-ColMeansVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[1]) })
        }
        'ROWMAXS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-RowMaxVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[0]) })
        }
        'COLMAXS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-ColMaxVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[1]) })
        }
        'ROWMINS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-RowMinVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[0]) })
        }
        'COLMINS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-ColMinVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[1]) })
        }
        'EIGVALS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-EigenValuesVector $Matrix } -ShapeFactory { param($InputShape) @($InputShape[0]) })
        }
        'EIGVECS' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-EigenVectorsMatrix $Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'SVALS' {
            Assert-MatrixValue $Argument
            return New-ArrayValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-SingularValuesVector $Matrix } -ShapeFactory { param($InputShape) @([Math]::Min($InputShape[0], $InputShape[1])) })
        }
        'CHOL' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-CholeskyMatrix $Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'LU' {
            Assert-MatrixValue $Argument
            return Invoke-BatchedTextUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-LuDecompositionText $Matrix }
        }
        'QR' {
            Assert-MatrixValue $Argument
            return Invoke-BatchedTextUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-QrDecompositionText $Matrix }
        }
        'COF' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-CofactorMatrix $Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'ADJ' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-AdjugateMatrix $Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'INV' {
            Assert-MatrixValue $Argument
            return New-MatrixValue (Invoke-BatchedMatrixUnary -ArrayData $Argument.Value -Operation { param($Matrix) Get-InverseMatrix $Matrix } -ShapeFactory { param($InputShape) $InputShape })
        }
        'NULL' {
            Assert-MatrixValue $Argument
            $batchShape = @(Get-BatchShape $Argument.Value.Shape)
            if ($batchShape.Count -eq 0) { return Get-NullSpaceValue $Argument.Value.Data }
            $parts = New-Object 'System.Collections.Generic.List[string]'
            $leafFactory = {
                param($Prefix)
                $slice = Get-TensorNodeByIndices -Node $Argument.Value.Data -Indices $Prefix
                $formatted = Format-ResultValue (Get-NullSpaceValue $slice)
                $parts.Add((Get-BatchIndexText -Indices $Prefix) + ': ' + [Environment]::NewLine + $formatted)
                return $null
            }
            [void](Build-BatchedTensorData -BatchShape $batchShape -LeafFactory $leafFactory)
            return New-TextValue (($parts -join ([Environment]::NewLine + [Environment]::NewLine)))
        }
        default { throw ($ui.UnsupportedOp + ' ' + $Name) }
    }
}

function Tokenize-Expression {
    param([string]$Expression)
    $Expression = Normalize-CommonSymbols $Expression
    $tokens = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $Expression.Length) {
        $ch = $Expression[$i]
        if ([char]::IsWhiteSpace($ch)) { $i++; continue }
        if ($ch -eq '.' -and $i + 1 -lt $Expression.Length -and '*/'.Contains([string]$Expression[$i + 1])) {
            $tokens.Add(@{ Type = '.' + [string]$Expression[$i + 1]; Value = '.' + [string]$Expression[$i + 1] })
            $i += 2
            continue
        }
        if ($ch -match '[\p{L}_]') {
            $start = $i
            while ($i -lt $Expression.Length -and $Expression[$i] -match '[\p{L}\p{Nd}_]') { $i++ }
            $tokens.Add(@{ Type = 'Identifier'; Value = $Expression.Substring($start, $i - $start) })
            continue
        }
        if ($ch -match '[0-9.]') {
            $start = $i
            if ($ch -eq '.') {
                if ($i + 1 -lt $Expression.Length -and $Expression[$i + 1] -match '[0-9]') {
                    $i++
                    while ($i -lt $Expression.Length -and $Expression[$i] -match '[0-9]') { $i++ }
                    $tokens.Add(@{ Type = 'Number'; Value = $Expression.Substring($start, $i - $start) })
                    continue
                }
            } else {
                while ($i -lt $Expression.Length -and $Expression[$i] -match '[0-9]') { $i++ }
                if ($i + 1 -lt $Expression.Length -and $Expression[$i] -eq '.' -and $Expression[$i + 1] -match '[0-9]') {
                    $i++
                    while ($i -lt $Expression.Length -and $Expression[$i] -match '[0-9]') { $i++ }
                }
                $tokens.Add(@{ Type = 'Number'; Value = $Expression.Substring($start, $i - $start) })
                continue
            }
        }
        if ('+-*/^()'.Contains([string]$ch)) {
            $tokens.Add(@{ Type = [string]$ch; Value = [string]$ch })
            $i++
            continue
        }
        throw $ui.UnexpectedToken
    }
    return $tokens
}

function Get-ReferencedMatrixNames {
    param([string]$Expression)
    $tokens = Tokenize-Expression $Expression
    $names = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $token = $tokens[$i]
        if ($token.Type -ne 'Identifier') { continue }
        $next = $null
        if ($i + 1 -lt $tokens.Count) { $next = $tokens[$i + 1] }
        if ($null -ne $next -and $next.Type -eq '(') { continue }
        $name = Get-CanonicalIdentifier $token.Value
        if (-not $names.Contains($name)) { $names.Add($name) }
    }
    return $names
}

function Parse-ExpressionText {
    param([string]$Expression, $Context)
    $tokens = Tokenize-Expression $Expression
    $script:exprIndex = 0

    function Peek-Token {
        if ($script:exprIndex -lt $tokens.Count) { return $tokens[$script:exprIndex] }
        return $null
    }

    function Read-Token {
        $token = Peek-Token
        if ($null -ne $token) { $script:exprIndex++ }
        return $token
    }

    function Parse-Primary {
        $token = Read-Token
        if ($null -eq $token) { throw $ui.UnexpectedToken }
        if ($token.Type -eq 'Number') { return New-ScalarValue ([double]$token.Value) }
        if ($token.Type -eq 'Identifier') {
            $next = Peek-Token
            if ($null -ne $next -and $next.Type -eq '(') {
                [void](Read-Token)
                $arg = Parse-AddSub
                $closing = Read-Token
                if ($null -eq $closing -or $closing.Type -ne ')') { throw $ui.MissingParen }
                return Invoke-FunctionValue -Name $token.Value -Argument $arg
            }
            $name = Get-CanonicalIdentifier $token.Value
            if (-not $Context.ContainsKey($name)) { throw ($ui.UnknownVar + ' ' + $name) }
            return New-MatrixValue $Context[$name]
        }
        if ($token.Type -eq '(') {
            $value = Parse-AddSub
            $closing = Read-Token
            if ($null -eq $closing -or $closing.Type -ne ')') { throw $ui.MissingParen }
            return $value
        }
        if ($token.Type -eq '-') {
            $value = Parse-Primary
            if ($value.Kind -eq 'Scalar') { return New-ScalarValue (-1 * $value.Value) }
            if ($value.Kind -eq 'Array') { return New-MatrixValue (New-ArrayData -Data (Multiply-MatrixByScalar $value.Value.Data -1.0) -Shape $value.Value.Shape) }
        }
        throw $ui.UnexpectedToken
    }

    function Parse-Power {
        $left = Parse-Primary
        while ($true) {
            $next = Peek-Token
            if ($null -eq $next -or $next.Type -ne '^') { break }
            [void](Read-Token)
            $right = Parse-Primary
            $left = Power-Values $left $right
        }
        return $left
    }

    function Parse-MulDiv {
        $left = Parse-Power
        while ($true) {
            $next = Peek-Token
            if ($null -eq $next -or ($next.Type -ne '*' -and $next.Type -ne '/' -and $next.Type -ne '.*' -and $next.Type -ne './')) { break }
            $op = Read-Token
            $right = Parse-Power
            if ($op.Type -eq '*') {
                $left = Multiply-Values $left $right
            } elseif ($op.Type -eq '/') {
                $left = Divide-Values $left $right
            } elseif ($op.Type -eq '.*') {
                $left = Multiply-Values $left $right -ElementWise
            } else {
                $left = Divide-Values $left $right -ElementWise
            }
        }
        return $left
    }

    function Parse-Mul {
        $left = Parse-Primary
        while ($true) {
            $next = Peek-Token
            if ($null -eq $next -or $next.Type -ne '*') { break }
            [void](Read-Token)
            $right = Parse-Primary
            $left = Multiply-Values $left $right
        }
        return $left
    }

    function Parse-AddSub {
        $left = Parse-MulDiv
        while ($true) {
            $next = Peek-Token
            if ($null -eq $next -or ($next.Type -ne '+' -and $next.Type -ne '-')) { break }
            $op = Read-Token
            $right = Parse-MulDiv
            if ($op.Type -eq '+') { $left = Add-Values $left $right } else { $left = Subtract-Values $left $right }
        }
        return $left
    }

    $result = Parse-AddSub
    if ($script:exprIndex -ne $tokens.Count) { throw $ui.UnexpectedToken }
    return $result
}

function Format-ResultValue {
    param($Result)
    switch ($Result.Kind) {
        'Scalar' { return Format-Number $Result.Value }
        'Array' { return Format-ArrayData $Result.Value }
        'Text' { return $Result.Value }
        default { return [string]$Result.Value }
    }
}

function Set-ButtonStyle {
    param([System.Windows.Forms.Button]$Button, [System.Drawing.Color]$BackColor, [System.Drawing.Color]$ForeColor)
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-FormScaleDefaults {
    param([System.Windows.Forms.Form]$Form)
    if ($null -eq $Form) { return }
    $Form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $Form.AutoSize = $false
}

$matrixStore = New-Object System.Collections.ArrayList
$script:isShuttingDown = $false
$script:helpDocs = @{}
$script:matrixFlow = $null
$script:variableQuickPanel = $null
$script:functionQuickPanel = $null
$script:helpQuickPanel = $null
$script:updateAdaptiveLayout = $null
$script:updateHelpQuickScroll = $null
$script:operationHistory = New-Object System.Collections.ArrayList
$script:historyForm = $null
$script:dataForm = $null
$script:dataTable = $null
$script:dataGridView = $null
$script:dataStatusLabel = $null
$script:dataFileLabel = $null
$script:dataPath = ''
$script:dataGridHost = $null
$script:dataStorageDir = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) 'MatrixWorkspaceApp'
$script:historyFilePath = Join-Path $script:dataStorageDir 'operation_history.json'
$script:desktopPetSettingsPath = Join-Path $script:dataStorageDir 'desktop_pet_settings.json'
$script:desktopPetSettings = $null
$script:desktopPetProcess = $null

function Get-MatrixLabelByIndex {
    param([int]$Index)
    return [char](65 + $Index)
}

function Get-MatrixDisplayText {
    param($Item)
    $label = [string]$Item.Label
    $descriptor = [string]$Item.Descriptor
    if ($Item.Mode -eq 'preset' -and -not [string]::IsNullOrWhiteSpace($descriptor)) {
        if ($descriptor.Length -gt 12) {
            return $label + ':' + $descriptor.Substring(0, 9) + '...'
        }
        return $label + ':' + $descriptor
    }
    return $label
}

function Ensure-AppStorageDirectory {
    if (-not (Test-Path -LiteralPath $script:dataStorageDir)) {
        [void](New-Item -ItemType Directory -Path $script:dataStorageDir -Force)
    }
}

function Get-DesktopPetDefaultRoot {
    return (Join-Path $script:appRoot 'DyberPet-main')
}

function Get-DesktopPetDefaultPythonPath {
    foreach ($cmdName in @('python.exe', 'py.exe')) {
        try {
            $cmd = Get-Command $cmdName -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) {
                return $cmd.Source
            }
        }
        catch {
        }
    }
    return 'python.exe'
}

function Get-DesktopPetDefaultSettings {
    return [ordered]@{
        PetRoot = Get-DesktopPetDefaultRoot
        PythonPath = Get-DesktopPetDefaultPythonPath
        PreferredLaunch = '自动'
        AutoStart = $false
    }
}

function Load-DesktopPetSettings {
    Ensure-AppStorageDirectory
    $defaults = Get-DesktopPetDefaultSettings
    if (-not (Test-Path -LiteralPath $script:desktopPetSettingsPath)) {
        $script:desktopPetSettings = $defaults
        return $script:desktopPetSettings
    }
    try {
        $raw = Get-Content -Path $script:desktopPetSettingsPath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            $script:desktopPetSettings = $defaults
            return $script:desktopPetSettings
        }
        $loaded = ConvertFrom-Json -InputObject $raw
        $script:desktopPetSettings = [ordered]@{
            PetRoot = if ([string]::IsNullOrWhiteSpace([string]$loaded.PetRoot)) { $defaults.PetRoot } else { [string]$loaded.PetRoot }
            PythonPath = if ([string]::IsNullOrWhiteSpace([string]$loaded.PythonPath)) { $defaults.PythonPath } else { [string]$loaded.PythonPath }
            PreferredLaunch = if ([string]::IsNullOrWhiteSpace([string]$loaded.PreferredLaunch)) { $defaults.PreferredLaunch } else { [string]$loaded.PreferredLaunch }
            AutoStart = [bool]$loaded.AutoStart
        }
    }
    catch {
        $script:desktopPetSettings = $defaults
    }
    return $script:desktopPetSettings
}

function Save-DesktopPetSettings {
    param($Settings)
    Ensure-AppStorageDirectory
    $payload = [pscustomobject]@{
        PetRoot = [string]$Settings.PetRoot
        PythonPath = [string]$Settings.PythonPath
        PreferredLaunch = [string]$Settings.PreferredLaunch
        AutoStart = [bool]$Settings.AutoStart
    }
    $payload | ConvertTo-Json -Depth 4 | Set-Content -Path $script:desktopPetSettingsPath -Encoding UTF8
    $script:desktopPetSettings = [ordered]@{
        PetRoot = $payload.PetRoot
        PythonPath = $payload.PythonPath
        PreferredLaunch = $payload.PreferredLaunch
        AutoStart = $payload.AutoStart
    }
}

function Get-DesktopPetLaunchInfo {
    param($Settings)
    if ($null -eq $Settings) { $Settings = Load-DesktopPetSettings }
    $petRoot = [string]$Settings.PetRoot
    $preferredLaunch = [string]$Settings.PreferredLaunch
    $pythonPath = [string]$Settings.PythonPath
    $exePath = Join-Path $petRoot 'run_DyberPet.exe'
    $scriptPath = Join-Path $petRoot 'run_DyberPet.py'
    if (-not (Test-Path -LiteralPath $petRoot)) {
        throw '桌宠项目目录不存在，请先在设置中确认 DyberPet-main 路径。'
    }
    $canUseExe = Test-Path -LiteralPath $exePath
    $canUseScript = Test-Path -LiteralPath $scriptPath
    switch ($preferredLaunch) {
        '优先EXE' {
            if ($canUseExe) { return [pscustomobject]@{ FilePath = $exePath; Arguments = @(); WorkingDirectory = $petRoot; Mode = 'EXE' } }
        }
        '优先Python' {
            if ($canUseScript) { return [pscustomobject]@{ FilePath = $pythonPath; Arguments = @($scriptPath); WorkingDirectory = $petRoot; Mode = 'Python' } }
        }
    }
    if ($canUseExe) { return [pscustomobject]@{ FilePath = $exePath; Arguments = @(); WorkingDirectory = $petRoot; Mode = 'EXE' } }
    if ($canUseScript) { return [pscustomobject]@{ FilePath = $pythonPath; Arguments = @($scriptPath); WorkingDirectory = $petRoot; Mode = 'Python' } }
    throw '未找到 run_DyberPet.exe 或 run_DyberPet.py，请检查桌宠项目目录。'
}

function Get-DesktopPetRuntimeHint {
    param($Settings)
    try {
        $launchInfo = Get-DesktopPetLaunchInfo -Settings $Settings
    }
    catch {
        return .Exception.Message
    }
    if ($launchInfo.Mode -eq 'EXE') {
        return '当前将优先直接启动桌宠 EXE；后续一起打包时，建议把桌宠 EXE 与 res / DyberPet 资源目录一起带上。'
    }
    return '当前将通过 Python 启动 DyberPet；请确认已安装 PySide6、PySide6-Fluent-Widgets、pynput、tendo。'
}

function Start-DesktopPet {
    param($Settings)
    if ($null -eq $Settings) { $Settings = Load-DesktopPetSettings }
    if ($null -ne $script:desktopPetProcess) {
        try {
            if (-not $script:desktopPetProcess.HasExited) {
                return $script:desktopPetProcess
            }
        }
        catch {
        }
    }
    $launchInfo = Get-DesktopPetLaunchInfo -Settings $Settings
    if ($launchInfo.Mode -eq 'Python' -and -not (Test-Path -LiteralPath $launchInfo.FilePath)) {
        throw '未找到可用的 Python 解释器，请在桌宠设置中指定 python.exe 或 py.exe。'
    }
    $startParams = @{
        FilePath = $launchInfo.FilePath
        WorkingDirectory = $launchInfo.WorkingDirectory
        PassThru = $true
    }
    if ($launchInfo.Arguments.Count -gt 0) { $startParams.ArgumentList = $launchInfo.Arguments }
    $script:desktopPetProcess = Start-Process @startParams
    return $script:desktopPetProcess
}

function Stop-DesktopPet {
    if ($null -eq $script:desktopPetProcess) { return }
    try {
        if (-not $script:desktopPetProcess.HasExited) {
            $null = $script:desktopPetProcess.CloseMainWindow()
            Start-Sleep -Milliseconds 800
            if (-not $script:desktopPetProcess.HasExited) {
                Stop-Process -Id $script:desktopPetProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
    }
    finally {
        $script:desktopPetProcess = $null
    }
}

function Show-DesktopPetSettingsDialog {
    $settings = Load-DesktopPetSettings

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = '桌宠设置'
    $dialog.Size = New-Object System.Drawing.Size(760, 520)
    $dialog.MinimumSize = New-Object System.Drawing.Size(680, 460)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'DyberPet 桌宠集成'
    $title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 12, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(20, 18)
    $title.Size = New-Object System.Drawing.Size(260, 28)
    $dialog.Controls.Add($title)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = '当前接入本地 DyberPet 项目。其源码已包含扑鼠标、跟随鼠标、拖拽、点击等互动逻辑；这里主要负责启动、停止、路径设置与后续打包入口。'
    $hint.Location = New-Object System.Drawing.Point(20, 52)
    $hint.Size = New-Object System.Drawing.Size(700, 48)
    $hint.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#657C90')
    $dialog.Controls.Add($hint)

    $petRootLabel = New-Object System.Windows.Forms.Label
    $petRootLabel.Text = '桌宠项目目录'
    $petRootLabel.Location = New-Object System.Drawing.Point(20, 116)
    $petRootLabel.Size = New-Object System.Drawing.Size(120, 24)
    $dialog.Controls.Add($petRootLabel)

    $petRootBox = New-Object System.Windows.Forms.TextBox
    $petRootBox.Location = New-Object System.Drawing.Point(20, 146)
    $petRootBox.Size = New-Object System.Drawing.Size(560, 30)
    $petRootBox.Anchor = 'Top,Left,Right'
    $petRootBox.Text = [string]$settings.PetRoot
    $dialog.Controls.Add($petRootBox)

    $petRootBrowse = New-Object System.Windows.Forms.Button
    $petRootBrowse.Text = '浏览目录'
    $petRootBrowse.Location = New-Object System.Drawing.Point(590, 144)
    $petRootBrowse.Size = New-Object System.Drawing.Size(110, 34)
    $petRootBrowse.Anchor = 'Top,Right'
    Set-ButtonStyle -Button $petRootBrowse -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($petRootBrowse)

    $pythonLabel = New-Object System.Windows.Forms.Label
    $pythonLabel.Text = 'Python / 启动器路径'
    $pythonLabel.Location = New-Object System.Drawing.Point(20, 194)
    $pythonLabel.Size = New-Object System.Drawing.Size(160, 24)
    $dialog.Controls.Add($pythonLabel)

    $pythonBox = New-Object System.Windows.Forms.TextBox
    $pythonBox.Location = New-Object System.Drawing.Point(20, 224)
    $pythonBox.Size = New-Object System.Drawing.Size(560, 30)
    $pythonBox.Anchor = 'Top,Left,Right'
    $pythonBox.Text = [string]$settings.PythonPath
    $dialog.Controls.Add($pythonBox)

    $pythonBrowse = New-Object System.Windows.Forms.Button
    $pythonBrowse.Text = '浏览程序'
    $pythonBrowse.Location = New-Object System.Drawing.Point(590, 222)
    $pythonBrowse.Size = New-Object System.Drawing.Size(110, 34)
    $pythonBrowse.Anchor = 'Top,Right'
    Set-ButtonStyle -Button $pythonBrowse -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($pythonBrowse)

    $modeLabel = New-Object System.Windows.Forms.Label
    $modeLabel.Text = '启动方式'
    $modeLabel.Location = New-Object System.Drawing.Point(20, 272)
    $modeLabel.Size = New-Object System.Drawing.Size(120, 24)
    $dialog.Controls.Add($modeLabel)

    $modeCombo = New-Object System.Windows.Forms.ComboBox
    $modeCombo.DropDownStyle = 'DropDownList'
    $modeCombo.Location = New-Object System.Drawing.Point(20, 302)
    $modeCombo.Size = New-Object System.Drawing.Size(180, 30)
    [void]$modeCombo.Items.AddRange(@('自动', '优先EXE', '优先Python'))
    $modeCombo.SelectedItem = [string]$settings.PreferredLaunch
    if ($modeCombo.SelectedIndex -lt 0) { $modeCombo.SelectedIndex = 0 }
    $dialog.Controls.Add($modeCombo)

    $autoStartCheck = New-Object System.Windows.Forms.CheckBox
    $autoStartCheck.Text = '主程序启动时自动尝试启动桌宠'
    $autoStartCheck.Location = New-Object System.Drawing.Point(220, 304)
    $autoStartCheck.Size = New-Object System.Drawing.Size(280, 26)
    $autoStartCheck.Checked = [bool]$settings.AutoStart
    $dialog.Controls.Add($autoStartCheck)

    $statusBox = New-Object System.Windows.Forms.TextBox
    $statusBox.Location = New-Object System.Drawing.Point(20, 352)
    $statusBox.Size = New-Object System.Drawing.Size(680, 86)
    $statusBox.Multiline = $true
    $statusBox.ReadOnly = $true
    $statusBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FBFCFE')
    $statusBox.Anchor = 'Top,Left,Right'
    $dialog.Controls.Add($statusBox)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = '保存设置'
    $saveButton.Location = New-Object System.Drawing.Point(20, 450)
    $saveButton.Size = New-Object System.Drawing.Size(96, 34)
    $saveButton.Anchor = 'Bottom,Left'
    Set-ButtonStyle -Button $saveButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
    $dialog.Controls.Add($saveButton)

    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = '启动桌宠'
    $startButton.Location = New-Object System.Drawing.Point(128, 450)
    $startButton.Size = New-Object System.Drawing.Size(96, 34)
    $startButton.Anchor = 'Bottom,Left'
    Set-ButtonStyle -Button $startButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
    $dialog.Controls.Add($startButton)

    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Text = '停止桌宠'
    $stopButton.Location = New-Object System.Drawing.Point(236, 450)
    $stopButton.Size = New-Object System.Drawing.Size(96, 34)
    $stopButton.Anchor = 'Bottom,Left'
    Set-ButtonStyle -Button $stopButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F5E7CD')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#6B4B11'))
    $dialog.Controls.Add($stopButton)

    $openDirButton = New-Object System.Windows.Forms.Button
    $openDirButton.Text = '打开目录'
    $openDirButton.Location = New-Object System.Drawing.Point(344, 450)
    $openDirButton.Size = New-Object System.Drawing.Size(96, 34)
    $openDirButton.Anchor = 'Bottom,Left'
    Set-ButtonStyle -Button $openDirButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($openDirButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = '关闭'
    $closeButton.Location = New-Object System.Drawing.Point(604, 450)
    $closeButton.Size = New-Object System.Drawing.Size(96, 34)
    $closeButton.Anchor = 'Bottom,Right'
    Set-ButtonStyle -Button $closeButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($closeButton)

    $buildSettingsObject = {
        return [ordered]@{
            PetRoot = $petRootBox.Text.Trim()
            PythonPath = $pythonBox.Text.Trim()
            PreferredLaunch = [string]$modeCombo.SelectedItem
            AutoStart = $autoStartCheck.Checked
        }
    }

    $refreshPetStatus = {
        $candidate = & $buildSettingsObject
        $lines = New-Object System.Collections.Generic.List[string]
        [void]$lines.Add((Get-DesktopPetRuntimeHint -Settings $candidate))
        [void]$lines.Add(('EXE 存在：{0}' -f (Test-Path -LiteralPath (Join-Path $candidate.PetRoot 'run_DyberPet.exe'))))
        [void]$lines.Add(('Python 脚本存在：{0}' -f (Test-Path -LiteralPath (Join-Path $candidate.PetRoot 'run_DyberPet.py'))))
        if ($null -ne $script:desktopPetProcess) {
            try {
                [void]$lines.Add(('当前桌宠进程状态：{0}' -f $(if ($script:desktopPetProcess.HasExited) { '已退出' } else { '运行中' })))
            }
            catch {
            }
        }
        $statusBox.Text = ($lines -join [Environment]::NewLine)
    }

    $petRootBrowse.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.SelectedPath = $petRootBox.Text
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $petRootBox.Text = $folderDialog.SelectedPath
            & $refreshPetStatus
        }
    })

    $pythonBrowse.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = '可执行程序|*.exe|所有文件|*.*'
        $fileDialog.FileName = $pythonBox.Text
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $pythonBox.Text = $fileDialog.FileName
            & $refreshPetStatus
        }
    })

    $saveButton.Add_Click({
        try {
            $candidate = & $buildSettingsObject
            Save-DesktopPetSettings -Settings $candidate
            & $refreshPetStatus
            [System.Windows.Forms.MessageBox]::Show('桌宠设置已保存。', '桌宠设置', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '桌宠设置', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })

    $startButton.Add_Click({
        try {
            $candidate = & $buildSettingsObject
            Save-DesktopPetSettings -Settings $candidate
            $process = Start-DesktopPet -Settings $candidate
            & $refreshPetStatus
            [System.Windows.Forms.MessageBox]::Show(('桌宠已启动，进程 ID：{0}' -f $process.Id), '桌宠设置', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '桌宠设置', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })

    $stopButton.Add_Click({
        Stop-DesktopPet
        & $refreshPetStatus
    })

    $openDirButton.Add_Click({
        $path = $petRootBox.Text.Trim()
        if (Test-Path -LiteralPath $path) {
            Start-Process explorer.exe -ArgumentList @($path) | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show('当前桌宠目录不存在。', '桌宠设置', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

    $closeButton.Add_Click({ $dialog.Close() })
    $petRootBox.Add_TextChanged($refreshPetStatus)
    $pythonBox.Add_TextChanged($refreshPetStatus)
    $modeCombo.Add_SelectedIndexChanged($refreshPetStatus)
    $autoStartCheck.Add_CheckedChanged($refreshPetStatus)
    & $refreshPetStatus
    [void]$dialog.ShowDialog($form)
    $dialog.Dispose()
}
function Load-OperationHistory {
    Ensure-AppStorageDirectory
    $script:operationHistory.Clear()
    if (-not (Test-Path -LiteralPath $script:historyFilePath)) { return }
    try {
        $raw = Get-Content -Path $script:historyFilePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return }
        $items = ConvertFrom-Json -InputObject $raw
        foreach ($item in @($items)) {
            [void]$script:operationHistory.Add([ordered]@{
                Time = [string]$item.Time
                Expression = [string]$item.Expression
                Status = [string]$item.Status
                Result = [string]$item.Result
            })
        }
    }
    catch {
    }
}

function Save-OperationHistory {
    Ensure-AppStorageDirectory
    $payload = @($script:operationHistory | ForEach-Object {
        [pscustomobject]@{
            Time = $_.Time
            Expression = $_.Expression
            Status = $_.Status
            Result = $_.Result
        }
    })
    $payload | ConvertTo-Json -Depth 4 | Set-Content -Path $script:historyFilePath -Encoding UTF8
}

function Add-OperationHistoryEntry {
    param([string]$Expression, [string]$Status, [string]$ResultText)
    $entry = [ordered]@{
        Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Expression = $Expression
        Status = $Status
        Result = $ResultText
    }
    [void]$script:operationHistory.Insert(0, $entry)
    while ($script:operationHistory.Count -gt 50) { $script:operationHistory.RemoveAt($script:operationHistory.Count - 1) }
    Save-OperationHistory
}

function Get-DataTableColumnNames {
    param([System.Data.DataTable]$Table)
    if ($null -eq $Table) { return @() }
    return @($Table.Columns | ForEach-Object { $_.ColumnName })
}

function Test-IsMissingValue {
    param($Value)
    if ($null -eq $Value -or $Value -eq [System.DBNull]::Value) { return $true }
    $text = [string]$Value
    return [string]::IsNullOrWhiteSpace($text)
}

function Convert-MatrixToDataTable {
    param($ArrayData, [string]$BaseName = 'C')
    $shape = @($ArrayData.Shape)
    if ($shape.Count -ne 2) { throw '当前仅支持将二维矩阵直接转换为表格。' }
    $table = New-Object System.Data.DataTable 'MatrixData'
    for ($c = 0; $c -lt $shape[1]; $c++) {
        [void]$table.Columns.Add(($BaseName + ($c + 1)), [string])
    }
    for ($r = 0; $r -lt $shape[0]; $r++) {
        $row = $table.NewRow()
        for ($c = 0; $c -lt $shape[1]; $c++) {
            $row[$c] = Format-Number $ArrayData.Data[$r][$c]
        }
        [void]$table.Rows.Add($row)
    }
    return ,$table
}

function Set-CurrentDataTable {
    param([System.Data.DataTable]$Table, [string]$SourceLabel = '')
    $script:dataTable = $Table
    if ($null -ne $script:dataGridView) {
        $script:dataGridView.DataSource = $null
        if ($null -ne $script:dataTable) { $script:dataGridView.DataSource = $script:dataTable }
    }
    if ($null -ne $script:dataStatusLabel) {
        if ($null -eq $Table) {
            $script:dataStatusLabel.Text = $ui.DataStatusReady
        } else {
            $script:dataStatusLabel.Text = $ui.DataLoaded + ' - ' + $Table.Rows.Count + ' x ' + $Table.Columns.Count
        }
    }
    if ($null -ne $script:dataFileLabel) {
        $script:dataFileLabel.Text = if ([string]::IsNullOrWhiteSpace($SourceLabel)) { $ui.DataStatusReady } else { $SourceLabel }
    }
}

function Ensure-DataLoaded {
    param([System.Data.DataTable]$Table = $null)
    $target = if ($null -ne $Table) { $Table } else { $script:dataTable }
    if ($null -eq $target) { throw $ui.NoDataLoaded }
}

function Import-DelimitedDataTable {
    param(
        [string]$Path,
        [int]$StartRow,
        [int]$StartCol,
        [int]$EndRow,
        [int]$EndCol,
        [bool]$FirstRowAsHeader
    )
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::UTF8)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    if ([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -eq '.txt') {
        $parser.SetDelimiters("`t", ",")
    } else {
        $parser.SetDelimiters(',')
    }
    $parser.HasFieldsEnclosedInQuotes = $true
    $rows = New-Object 'System.Collections.Generic.List[object]'
    try {
        while (-not $parser.EndOfData) {
            $rows.Add(@($parser.ReadFields()))
        }
    }
    finally {
        $parser.Close()
    }
    $table = Convert-GridRowsToDataTable -Rows $rows -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader
    return ,$table
}

function Convert-GridRowsToDataTable {
    param(
        [object[]]$Rows,
        [int]$StartRow,
        [int]$StartCol,
        [int]$EndRow,
        [int]$EndCol,
        [switch]$FirstRowAsHeader
    )
    $startRowIndex = [Math]::Max(0, $StartRow - 1)
    $startColIndex = [Math]::Max(0, $StartCol - 1)
    $rowLimit = if ($EndRow -le 0) { $Rows.Count } else { [Math]::Min($Rows.Count, $EndRow) }
    $filteredRows = @()
    for ($r = $startRowIndex; $r -lt $rowLimit; $r++) {
        $row = @($Rows[$r])
        $colLimit = if ($EndCol -le 0) { $row.Count } else { [Math]::Min($row.Count, $EndCol) }
        if ($startColIndex -ge $colLimit) { continue }
        $segment = @()
        for ($c = $startColIndex; $c -lt $colLimit; $c++) { $segment += ,$row[$c] }
        $filteredRows += ,$segment
    }
    if ($filteredRows.Count -eq 0) { throw '所选范围内没有数据。' }

    $table = New-Object System.Data.DataTable 'ImportedData'
    $headers = @()
    $dataStart = 0
    if ($FirstRowAsHeader) {
        for ($i = 0; $i -lt $filteredRows[0].Count; $i++) {
            $name = [string]$filteredRows[0][$i]
            if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Column' + ($i + 1) }
            $headers += ,$name
        }
        $dataStart = 1
    } else {
        for ($i = 0; $i -lt $filteredRows[0].Count; $i++) { $headers += ,('Column' + ($i + 1)) }
    }
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $header = $headers[$i]
        if ($table.Columns.Contains($header)) { $header = $header + '_' + ($i + 1) }
        [void]$table.Columns.Add($header, [string])
    }
    for ($r = $dataStart; $r -lt $filteredRows.Count; $r++) {
        $row = $table.NewRow()
        for ($c = 0; $c -lt $table.Columns.Count; $c++) {
            if ($c -lt $filteredRows[$r].Count) { $row[$c] = [string]$filteredRows[$r][$c] } else { $row[$c] = '' }
        }
        [void]$table.Rows.Add($row)
    }
    return ,$table
}

function Import-ExcelDataTable {
    param(
        [string]$Path,
        [string]$SheetName,
        [int]$StartRow,
        [int]$StartCol,
        [int]$EndRow,
        [int]$EndCol,
        [bool]$FirstRowAsHeader
    )
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path, $null, $true)
        $sheet = if ([string]::IsNullOrWhiteSpace($SheetName)) { $workbook.Worksheets.Item(1) } else { $workbook.Worksheets.Item($SheetName) }
        $usedRange = $sheet.UsedRange
        $values = $usedRange.Value2
        $rows = New-Object 'System.Collections.Generic.List[object]'
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        for ($r = 1; $r -le $rowCount; $r++) {
            $line = New-Object 'System.Collections.Generic.List[object]'
            for ($c = 1; $c -le $colCount; $c++) {
                $cellValue = if ($rowCount -eq 1 -and $colCount -eq 1) { $values } else { $values[$r, $c] }
                $cellText = ''
                if ($null -ne $cellValue) { $cellText = [string]$cellValue }
                $line.Add($cellText)
            }
            $rows.Add($line)
        }
        $table = Convert-GridRowsToDataTable -Rows $rows.ToArray() -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader
        return ,$table
    }
    finally {
        if ($workbook) { $workbook.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null }
        foreach ($obj in @($usedRange, $sheet, $workbook, $excel)) {
            if ($null -ne $obj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Import-DataTableFromFile {
    param(
        [string]$Path,
        [int]$StartRow = 1,
        [int]$StartCol = 1,
        [int]$EndRow = 0,
        [int]$EndCol = 0,
        [bool]$FirstRowAsHeader = $false,
        [string]$SheetName = ''
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { throw $ui.NeedFile }
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $table = switch ($ext) {
        '.csv' { Import-DelimitedDataTable -Path $Path -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader }
        '.txt' { Import-DelimitedDataTable -Path $Path -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader }
        '.xlsx' { Import-ExcelDataTable -Path $Path -SheetName $SheetName -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader }
        '.xls' { Import-ExcelDataTable -Path $Path -SheetName $SheetName -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader }
        '.xlsm' { Import-ExcelDataTable -Path $Path -SheetName $SheetName -StartRow $StartRow -StartCol $StartCol -EndRow $EndRow -EndCol $EndCol -FirstRowAsHeader:$FirstRowAsHeader }
        default { throw $ui.FileNotSupported }
    }
    return ,$table
}

function Get-DataOverviewText {
    param([System.Data.DataTable]$Table)
    Ensure-DataLoaded -Table $Table
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('Rows: ' + $Table.Rows.Count)
    $lines.Add('Cols: ' + $Table.Columns.Count)
    $lines.Add('')
    foreach ($column in $Table.Columns) {
        $missing = 0
        foreach ($row in $Table.Rows) { if (Test-IsMissingValue $row[$column.ColumnName]) { $missing++ } }
        $lines.Add($column.ColumnName + ' | Missing: ' + $missing)
    }
    return ($lines -join [Environment]::NewLine)
}

function Get-MissingRatioDataTable {
    param([System.Data.DataTable]$Table)
    Ensure-DataLoaded -Table $Table
    $result = New-Object System.Data.DataTable 'MissingRatio'
    [void]$result.Columns.Add('Column', [string])
    [void]$result.Columns.Add('MissingCount', [string])
    [void]$result.Columns.Add('MissingRatio', [string])
    foreach ($column in $Table.Columns) {
        $missing = 0
        foreach ($row in $Table.Rows) { if (Test-IsMissingValue $row[$column.ColumnName]) { $missing++ } }
        $ratio = if ($Table.Rows.Count -eq 0) { 0.0 } else { [math]::Round(($missing / [double]$Table.Rows.Count) * 100, 2) }
        $newRow = $result.NewRow()
        $newRow['Column'] = $column.ColumnName
        $newRow['MissingCount'] = [string]$missing
        $newRow['MissingRatio'] = $ratio.ToString('0.##') + '%'
        [void]$result.Rows.Add($newRow)
    }
    return ,$result
}

function Remove-MissingRowsFromTable {
    param([System.Data.DataTable]$Table)
    Ensure-DataLoaded -Table $Table
    $clone = $Table.Clone()
    foreach ($row in $Table.Rows) {
        $keep = $true
        foreach ($column in $Table.Columns) {
            if (Test-IsMissingValue $row[$column.ColumnName]) { $keep = $false; break }
        }
        if ($keep) { [void]$clone.ImportRow($row) }
    }
    return ,$clone
}

function Fill-MissingWithZero {
    param([System.Data.DataTable]$Table)
    Ensure-DataLoaded -Table $Table
    $clone = $Table.Copy()
    foreach ($row in $clone.Rows) {
        foreach ($column in $clone.Columns) {
            if (Test-IsMissingValue $row[$column.ColumnName]) { $row[$column.ColumnName] = '0' }
        }
    }
    return ,$clone
}

function Fill-MissingWithMean {
    param([System.Data.DataTable]$Table)
    Ensure-DataLoaded -Table $Table
    $clone = $Table.Copy()
    foreach ($column in $clone.Columns) {
        $values = @()
        foreach ($row in $clone.Rows) {
            $number = 0.0
            if (-not (Test-IsMissingValue $row[$column.ColumnName]) -and [double]::TryParse([string]$row[$column.ColumnName], [ref]$number)) {
                $values += $number
            }
        }
        if ($values.Count -eq 0) { continue }
        $mean = ($values | Measure-Object -Average).Average
        foreach ($row in $clone.Rows) {
            if (Test-IsMissingValue $row[$column.ColumnName]) { $row[$column.ColumnName] = ([math]::Round($mean, 4)).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
        }
    }
    return ,$clone
}

function Select-TableRange {
    param([System.Data.DataTable]$Table, [int]$StartRow, [int]$EndRow, [int]$StartCol, [int]$EndCol)
    Ensure-DataLoaded -Table $Table
    $columnNames = Get-DataTableColumnNames $Table
    $startRowIndex = [Math]::Max(0, $StartRow - 1)
    $endRowIndex = if ($EndRow -le 0) { $Table.Rows.Count - 1 } else { [Math]::Min($Table.Rows.Count - 1, $EndRow - 1) }
    $startColIndex = [Math]::Max(0, $StartCol - 1)
    $endColIndex = if ($EndCol -le 0) { $columnNames.Count - 1 } else { [Math]::Min($columnNames.Count - 1, $EndCol - 1) }
    $result = New-Object System.Data.DataTable 'SelectedData'
    for ($c = $startColIndex; $c -le $endColIndex; $c++) { [void]$result.Columns.Add($columnNames[$c], [string]) }
    for ($r = $startRowIndex; $r -le $endRowIndex; $r++) {
        if ($r -lt 0 -or $r -ge $Table.Rows.Count) { continue }
        $newRow = $result.NewRow()
        for ($c = $startColIndex; $c -le $endColIndex; $c++) {
            $newRow[$columnNames[$c]] = [string]$Table.Rows[$r][$columnNames[$c]]
        }
        [void]$result.Rows.Add($newRow)
    }
    return ,$result
}

function Set-DataCellValue {
    param([System.Data.DataTable]$Table, [int]$RowIndex, [string]$ColumnName, [string]$NewValue)
    Ensure-DataLoaded -Table $Table
    if ($RowIndex -lt 1 -or $RowIndex -gt $Table.Rows.Count) { throw '行号超出范围。' }
    if (-not $Table.Columns.Contains($ColumnName)) { throw '列名不存在。' }
    $clone = $Table.Copy()
    $clone.Rows[$RowIndex - 1][$ColumnName] = $NewValue
    return ,$clone
}

function Get-GroupedStatisticsTable {
    param([System.Data.DataTable]$Table, [string]$GroupColumn, [string]$ValueColumn, [string]$Method)
    Ensure-DataLoaded -Table $Table
    if (-not $Table.Columns.Contains($GroupColumn)) { throw '分组列不存在。' }
    if (-not $Table.Columns.Contains($ValueColumn)) { throw '统计列不存在。' }
    $result = New-Object System.Data.DataTable 'GroupedStats'
    [void]$result.Columns.Add($GroupColumn, [string])
    [void]$result.Columns.Add($Method, [string])
    $groups = @{}
    foreach ($row in $Table.Rows) {
        $groupKey = [string]$row[$GroupColumn]
        if (-not $groups.ContainsKey($groupKey)) { $groups[$groupKey] = @() }
        $groups[$groupKey] += ,$row[$ValueColumn]
    }
    foreach ($groupKey in ($groups.Keys | Sort-Object)) {
        $items = @($groups[$groupKey])
        $numeric = @()
        foreach ($item in $items) {
            $number = 0.0
            if ([double]::TryParse([string]$item, [ref]$number)) { $numeric += $number }
        }
        $value = switch ($Method) {
            'COUNT' { [string]$items.Count }
            'SUM' { if ($numeric.Count -eq 0) { '0' } else { ([math]::Round(($numeric | Measure-Object -Sum).Sum, 4)).ToString() } }
            default { if ($numeric.Count -eq 0) { '0' } else { ([math]::Round(($numeric | Measure-Object -Average).Average, 4)).ToString() } }
        }
        $row = $result.NewRow()
        $row[$GroupColumn] = $groupKey
        $row[$Method] = $value
        [void]$result.Rows.Add($row)
    }
    return ,$result
}

function Show-TextViewerDialog {
    param([string]$Title, [string]$Text)

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(760, 520)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = 'Fill'
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.ScrollBars = 'Both'
    $box.WordWrap = $false
    $box.Font = New-Object System.Drawing.Font('Consolas', 11)
    $box.Text = $Text
    $dialog.Controls.Add($box)
    [void]$dialog.ShowDialog($form)
    $dialog.Dispose()
}

function Show-DataTableDialog {
    param([string]$Title, [System.Data.DataTable]$Table)

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(900, 560)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'DisplayedCells'
    $grid.DataSource = $Table
    $dialog.Controls.Add($grid)
    [void]$dialog.ShowDialog($form)
    $dialog.Dispose()
}

foreach ($moduleName in @('data_profile.ps1', 'data_plotting.ps1')) {
    $modulePath = Join-Path $script:moduleRoot $moduleName
    if (Test-Path -LiteralPath $modulePath) {
        . $modulePath
    }
}

function Get-HelpKey {
    param([string]$Query)
    $normalized = (Normalize-CommonSymbols $Query).Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }
    $upper = $normalized.ToUpperInvariant()
    switch -Regex ($upper) {
        '^\+$' { return '+' }
        '^-$' { return '-' }
        '^\*$' { return '*' }
        '^/$' { return '/' }
        '^\.\*$' { return '.*' }
        '^\./$' { return './' }
        '^\^$' { return '^' }
        '^\^\d+$' { return '^' }
        '^BLOCK$' { return 'BLOCK' }
        '^I\d+$' { return 'SPECIAL_I' }
        '^ZEROS\(' { return 'SPECIAL_ZEROS' }
        '^ONES\(' { return 'SPECIAL_ONES' }
        '^DIAG\(' { return 'SPECIAL_DIAG' }
        '^ROWSUMS?$' { return 'ROWSUMS' }
        '^COLSUMS?$' { return 'COLSUMS' }
        '^ROWMEANS?$' { return 'ROWMEANS' }
        '^COLMEANS?$' { return 'COLMEANS' }
        '^ROWMAXS?$' { return 'ROWMAXS' }
        '^COLMAXS?$' { return 'COLMAXS' }
        '^ROWMINS?$' { return 'ROWMINS' }
        '^COLMINS?$' { return 'COLMINS' }
        '^EIGVALS?$' { return 'EIGVALS' }
        '^EIGVECS?$' { return 'EIGVECS' }
        '^SVALS?$' { return 'SVALS' }
        '^CHOL(ESKY)?$' { return 'CHOL' }
        '^LU$' { return 'LU' }
        '^QR$' { return 'QR' }
        '^DATA$' { return 'DATA' }
        '^HISTORY$' { return 'HISTORY' }
    }
    if ($normalized.Contains('|') -or $normalized.Contains((S '%E5%88%86%E5%9D%97'))) { return 'BLOCK' }
    if ($normalized.Contains((S '%E5%8D%95%E4%BD%8D%E9%98%B5'))) { return 'SPECIAL_I' }
    if ($normalized.Contains((S '%E9%9B%B6%E7%9F%A9%E9%98%B5'))) { return 'SPECIAL_ZEROS' }
    if ($normalized.Contains('ones') -or $normalized.Contains((S '%E5%85%A8%201'))) { return 'SPECIAL_ONES' }
    if ($normalized.Contains((S '%E5%AF%B9%E8%A7%92'))) { return 'SPECIAL_DIAG' }
    if ($normalized.Contains((S '%E6%95%B0%E6%8D%AE%E5%AF%BC%E5%85%A5'))) { return 'DATA' }
    if ($normalized.Contains((S '%E5%8E%86%E5%8F%B2%E8%AE%B0%E5%BD%95'))) { return 'HISTORY' }
    $canonical = Get-CanonicalIdentifier $normalized
    if ($script:helpDocs.ContainsKey($canonical)) { return $canonical }
    return $null
}

function Insert-ExpressionText {
    param([System.Windows.Forms.TextBox]$TextBox, [string]$Text)
    $start = $TextBox.SelectionStart
    $TextBox.Text = $TextBox.Text.Insert($start, $Text)
    $TextBox.SelectionStart = $start + $Text.Length
    $TextBox.SelectionLength = 0
    $TextBox.Focus()
}

function Set-MatrixCardMode {
    param($Item, [string]$Mode, [string]$Descriptor = '')
    if ($null -eq $Item) { return }
    if ($Mode -eq 'preset' -and -not [string]::IsNullOrWhiteSpace($Descriptor)) {
        $Item.Mode = 'preset'
        $Item.Descriptor = $Descriptor
        $Item.ModeValueLabel.Text = $ui.Preset + ': ' + $Descriptor
        $Item.ModeValueLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#1F5E86')
    }
    else {
        $Item.Mode = 'manual'
        $Item.Descriptor = ''
        $Item.ModeValueLabel.Text = $ui.ManualInput
        $Item.ModeValueLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#4F6479')
    }
}

function Refresh-VariableQuickButtons {
    if ($null -eq $script:variableQuickPanel) { return }
    $script:variableQuickPanel.SuspendLayout()
    $script:variableQuickPanel.Controls.Clear()
    foreach ($item in $matrixStore) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = Get-MatrixDisplayText -Item $item
        $button.Size = New-Object System.Drawing.Size(118, 32)
        $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 8)
        Set-ButtonStyle -Button $button -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EAF2FA')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#23435C'))
        $insertValue = [string]$item.Label
        $button.Add_Click({ Insert-ExpressionText -TextBox $expressionTextBox -Text $insertValue }.GetNewClosure())
        $script:variableQuickPanel.Controls.Add($button)
    }
    $script:variableQuickPanel.ResumeLayout()
    if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout }
}

function Update-MatrixCountLabel {
    if ($countValueLabel) { $countValueLabel.Text = $matrixStore.Count.ToString() }
    if ($removeMatrixButton) { $removeMatrixButton.Enabled = ($matrixStore.Count -gt 1) }
    if ($addMatrixButton) { $addMatrixButton.Enabled = ($matrixStore.Count -lt 8) }
    if ($addSpecialButton) { $addSpecialButton.Enabled = ($matrixStore.Count -lt 8) }
}

function Refresh-MatrixCards {
    if ($null -eq $script:matrixFlow) { return }
    $script:matrixFlow.SuspendLayout()
    $script:matrixFlow.Controls.Clear()
    for ($i = 0; $i -lt $matrixStore.Count; $i++) {
        $item = $matrixStore[$i]
        $item.Label = [string](Get-MatrixLabelByIndex $i)
        $item.TitleLabel.Text = 'Matrix ' + $item.Label
        $script:matrixFlow.Controls.Add($item.Card)
    }
    $script:matrixFlow.ResumeLayout()
    Update-MatrixCountLabel
    Refresh-VariableQuickButtons
    if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout }
}

function Build-MatrixContext {
    param([string]$Expression)
    $context = @{}
    $index = @{}
    foreach ($item in $matrixStore) { $index[$item.Label.ToUpperInvariant()] = $item }
    $resolving = New-Object 'System.Collections.Generic.List[string]'

    function Resolve-MatrixValue {
        param([string]$Name)
        $canonical = Get-CanonicalIdentifier $Name
        if ($context.ContainsKey($canonical)) { return $context[$canonical] }
        if ($resolving.Contains($canonical)) {
            throw ('Matrix ' + $canonical + ' ' + 'contains circular block references.')
        }
        if (-not $index.ContainsKey($canonical)) { throw ($ui.UnknownVar + ' ' + $canonical) }
        $resolving.Add($canonical)
        try {
            $resolver = { param($RefName) Resolve-MatrixValue -Name $RefName }
            $context[$canonical] = Parse-Matrix -Text $index[$canonical].TextBox.Text -Name ('Matrix ' + $canonical) -ResolveMatrix $resolver
            return $context[$canonical]
        }
        finally {
            $resolving.Remove($canonical) | Out-Null
        }
    }

    $neededNames = Get-ReferencedMatrixNames $Expression
    foreach ($name in $neededNames) {
        [void](Resolve-MatrixValue -Name $name)
    }
    return $context
}

$script:helpDocs = @{
    '+' = "`r`n含义：逐元素加法。`r`n用法：A + B`r`n要求：两个数组形状完全一致。支持二维矩阵、批量矩阵和更高维数组。"
    '-' = "`r`n含义：逐元素减法。`r`n用法：A - B`r`n要求：两个数组形状完全一致。"
    '*' = "`r`n含义：矩阵乘法。`r`n用法：A * B`r`n规则：如果操作数至少有两维，则把最后两维当作矩阵、前置维度当作批量维；前置维度必须一致，且左矩阵列数等于右矩阵行数。"
    '/' = "`r`n含义：矩阵右除。`r`n用法：A / B`r`n规则：按 A * inv(B) 处理；对批量矩阵逐批执行。"
    '.*' = "`r`n含义：逐元素乘法。`r`n用法：A .* B、A .* 2、2 .* A`r`n要求：两个数组形状一致，或一侧为标量。"
    './' = "`r`n含义：逐元素除法。`r`n用法：A ./ B、A ./ 2、2 ./ A`r`n要求：两个数组形状一致，或一侧为标量。"
    '^' = "`r`n含义：幂运算。`r`n用法：A ^ 2`r`n规则：标量按普通幂运算；至少二维的方阵按矩阵幂并对批量维逐批处理；一维数组按逐元素幂。"
    'T' = "`r`n含义：转置。`r`n用法：T(A) 或 转置(A)`r`n规则：对至少二维输入，把最后两维转置；批量维保持不变。"
    'REF' = "`r`n含义：行阶梯形。`r`n用法：REF(A)`r`n规则：对每个矩阵切片独立计算。"
    'RREF' = "`r`n含义：最简行阶梯形。`r`n用法：RREF(A)`r`n规则：对每个矩阵切片独立计算。"
    'RANK' = "`r`n含义：矩阵秩。`r`n用法：RANK(A)`r`n返回：单个矩阵返回标量；批量矩阵返回对应批量形状的标量数组。"
    'TR' = "`r`n含义：迹。`r`n用法：TR(A) 或 TRACE(A)`r`n返回：按最后两维矩阵求迹，并保留前置批量维。"
    'NORM' = "`r`n含义：范数。`r`n用法：NORM(A)`r`n规则：对任意维数组可用，返回整体 Frobenius 范数。"
    'ROWSUMS' = "`r`n含义：按行求和。`r`n用法：ROWSUMS(A)`r`n返回：每一行的和；批量矩阵返回批量向量。"
    'COLSUMS' = "`r`n含义：按列求和。`r`n用法：COLSUMS(A)`r`n返回：每一列的和；批量矩阵返回批量向量。"
    'ROWMEANS' = "`r`n含义：按行求均值。`r`n用法：ROWMEANS(A)`r`n返回：每一行的平均值；批量矩阵返回批量向量。"
    'COLMEANS' = "`r`n含义：按列求均值。`r`n用法：COLMEANS(A)`r`n返回：每一列的平均值；批量矩阵返回批量向量。"
    'ROWMAXS' = "`r`n含义：按行求最大值。`r`n用法：ROWMAXS(A)`r`n返回：每一行的最大值；批量矩阵返回批量向量。"
    'COLMAXS' = "`r`n含义：按列求最大值。`r`n用法：COLMAXS(A)`r`n返回：每一列的最大值；批量矩阵返回批量向量。"
    'ROWMINS' = "`r`n含义：按行求最小值。`r`n用法：ROWMINS(A)`r`n返回：每一行的最小值；批量矩阵返回批量向量。"
    'COLMINS' = "`r`n含义：按列求最小值。`r`n用法：COLMINS(A)`r`n返回：每一列的最小值；批量矩阵返回批量向量。"
    'INV' = "`r`n含义：逆矩阵。`r`n用法：INV(A)`r`n规则：要求最后两维是可逆方阵；批量矩阵逐批求逆。"
    'EIGVALS' = "`r`n含义：特征值。`r`n用法：EIGVALS(A)`r`n规则：当前稳定支持实对称方阵；按特征值从大到小返回；批量矩阵逐批返回向量。"
    'EIGVECS' = "`r`n含义：特征向量。`r`n用法：EIGVECS(A)`r`n规则：当前稳定支持实对称方阵；返回按特征值降序排列的列向量矩阵。"
    'SVALS' = "`r`n含义：奇异值。`r`n用法：SVALS(A)`r`n规则：通过 A^T A 的特征值计算奇异值，并按从大到小返回。"
    'CHOL' = "`r`n含义：楚列斯基分解。`r`n用法：CHOL(A)`r`n规则：要求输入为对称正定矩阵；返回下三角矩阵 L。"
    'LU' = "`r`n含义：LU 分解。`r`n用法：LU(A)`r`n规则：要求方阵；当前返回文本结果，内容包含 P、L、U 三个矩阵。"
    'QR' = "`r`n含义：QR 分解。`r`n用法：QR(A)`r`n规则：适用于一般矩阵；当前返回文本结果，内容包含 Q、R 两个矩阵。"
    'DET' = "`r`n含义：行列式。`r`n用法：DET(A)`r`n规则：要求最后两维是方阵；批量矩阵返回标量数组。"
    'COF' = "`r`n含义：代数余子式矩阵。`r`n用法：COF(A)`r`n规则：要求最后两维是方阵；批量矩阵逐批计算。"
    'ADJ' = "`r`n含义：伴随矩阵。`r`n用法：ADJ(A)`r`n规则：要求最后两维是方阵；批量矩阵逐批计算。"
    'NULL' = "`r`n含义：零空间。`r`n用法：NULL(A)`r`n规则：单个矩阵返回零空间文本结果；批量矩阵按批次输出每个切片的零空间说明。"
    'BLOCK' = "`r`n含义：分块矩阵输入。`r`n用法：在矩阵输入框写 [A | B; B | A]`r`n规则：同行块的行数必须一致，同列块的列数必须一致；标量会按 1x1 块处理；块中可引用其他矩阵变量或特殊矩阵简写。"
    'SPECIAL_I' = "`r`n含义：单位矩阵简写。`r`n用法：I3、单位阵3`r`n结果：生成 3x3 单位矩阵。"
    'SPECIAL_ZEROS' = "`r`n含义：零矩阵简写。`r`n用法：zeros(2,3)`r`n结果：生成 2x3 零矩阵。"
    'SPECIAL_ONES' = "`r`n含义：全 1 矩阵简写。`r`n用法：ones(2,3)`r`n结果：生成 2x3 全 1 矩阵。"
    'SPECIAL_DIAG' = "`r`n含义：对角矩阵简写。`r`n用法：diag(1,2,3)`r`n结果：生成以给定序列为主对角线的方阵。"
    'DATA' = "`r`n含义：数据导入与处理窗口。`r`n功能：导入本地 CSV / Excel，选择范围，保留为表格格式，并执行缺失值清洗、数据概览、遴选行列、修改数据、分组统计、导出等操作。"
    'HISTORY' = "`r`n含义：历史记录。`r`n功能：保存最近 50 次表达式运算过程，可回看时间、表达式、结果，并重新载入表达式。"
}

function Show-HelpContent {
    param([string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query)) {
        $helpTextBox.Text = $ui.HelpEmpty
        return
    }
    $key = Get-HelpKey -Query $Query
    if ($null -eq $key -or -not $script:helpDocs.ContainsKey($key)) {
        $helpTextBox.Text = $ui.HelpNotFound
        return
    }
    $helpTextBox.Text = ($Query.Trim() + $script:helpDocs[$key])
}

function Add-HelpQuickButton {
    param([string]$Text, [string]$Query)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Size = New-Object System.Drawing.Size(108, 30)
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 8)
    Set-ButtonStyle -Button $button -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $searchValue = $Query
    $button.Add_Click({
        $helpSearchBox.Text = $searchValue
        Show-HelpContent -Query $searchValue
    }.GetNewClosure())
    $script:helpQuickPanel.Controls.Add($button)
    if ($null -ne $script:updateHelpQuickScroll) { & $script:updateHelpQuickScroll }
}

function Execute-Expression {
    try {
        $expr = $expressionTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($expr)) { throw ($ui.Expression + $ui.ParseEmptySuffix) }
        $context = Build-MatrixContext -Expression $expr
        $result = Parse-ExpressionText -Expression $expr -Context $context
        $formattedResult = Format-ResultValue $result
        $resultTextBox.Text = $formattedResult
        $statusLabel.Text = $ui.Done + ': ' + $expr
        Add-OperationHistoryEntry -Expression $expr -Status $statusLabel.Text -ResultText $formattedResult
    }
    catch {
        $statusLabel.Text = $ui.EvalFailed
        Add-OperationHistoryEntry -Expression ($expressionTextBox.Text.Trim()) -Status $ui.EvalFailed -ResultText $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.EvalFailed, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}
function Show-SpecialMatrixDialog {
    param($Item)

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.SpecialConfig + ' - Matrix ' + $Item.Label
    $dialog.Size = New-Object System.Drawing.Size(500, 360)
    $dialog.MinimumSize = New-Object System.Drawing.Size(460, 320)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $dialog.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $dialog.FormBorderStyle = 'Sizable'
    $dialog.MaximizeBox = $true
    $dialog.MinimizeBox = $false

    $dialogBody = New-Object System.Windows.Forms.Panel
    $dialogBody.Dock = 'Fill'
    $dialogBody.Padding = New-Object System.Windows.Forms.Padding(20, 18, 20, 12)
    $dialog.Controls.Add($dialogBody)

    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = 'Bottom'
    $buttonPanel.Height = 52
    $dialogBody.Controls.Add($buttonPanel)

    $typeLabel = New-Object System.Windows.Forms.Label
    $typeLabel.Text = $ui.Type
    $typeLabel.Location = New-Object System.Drawing.Point(20, 22)
    $typeLabel.Size = New-Object System.Drawing.Size(80, 24)
    $dialogBody.Controls.Add($typeLabel)

    $typeCombo = New-Object System.Windows.Forms.ComboBox
    $typeCombo.DropDownStyle = 'DropDownList'
    $typeCombo.Location = New-Object System.Drawing.Point(112, 18)
    $typeCombo.Size = New-Object System.Drawing.Size(328, 30)
    $typeCombo.Anchor = 'Top,Left,Right'
    [void]$typeCombo.Items.AddRange(@($ui.Identity, $ui.ZeroMatrix, $ui.OnesMatrix, $ui.DiagMatrix))
    $typeCombo.SelectedIndex = 0
    $dialogBody.Controls.Add($typeCombo)

    $sizeLabel = New-Object System.Windows.Forms.Label
    $sizeLabel.Text = $ui.Size
    $sizeLabel.Location = New-Object System.Drawing.Point(20, 70)
    $sizeLabel.Size = New-Object System.Drawing.Size(80, 24)
    $dialogBody.Controls.Add($sizeLabel)

    $sizeInput = New-Object System.Windows.Forms.NumericUpDown
    $sizeInput.Location = New-Object System.Drawing.Point(112, 66)
    $sizeInput.Size = New-Object System.Drawing.Size(120, 30)
    $sizeInput.Minimum = 1
    $sizeInput.Maximum = 12
    $sizeInput.Value = 3
    $dialogBody.Controls.Add($sizeInput)

    $rowsLabel = New-Object System.Windows.Forms.Label
    $rowsLabel.Text = $ui.Rows
    $rowsLabel.Location = New-Object System.Drawing.Point(20, 110)
    $rowsLabel.Size = New-Object System.Drawing.Size(80, 24)
    $dialogBody.Controls.Add($rowsLabel)

    $rowsInput = New-Object System.Windows.Forms.NumericUpDown
    $rowsInput.Location = New-Object System.Drawing.Point(112, 106)
    $rowsInput.Size = New-Object System.Drawing.Size(120, 30)
    $rowsInput.Minimum = 1
    $rowsInput.Maximum = 12
    $rowsInput.Value = 3
    $dialogBody.Controls.Add($rowsInput)

    $colsLabel = New-Object System.Windows.Forms.Label
    $colsLabel.Text = $ui.Cols
    $colsLabel.Location = New-Object System.Drawing.Point(20, 150)
    $colsLabel.Size = New-Object System.Drawing.Size(80, 24)
    $dialogBody.Controls.Add($colsLabel)

    $colsInput = New-Object System.Windows.Forms.NumericUpDown
    $colsInput.Location = New-Object System.Drawing.Point(112, 146)
    $colsInput.Size = New-Object System.Drawing.Size(120, 30)
    $colsInput.Minimum = 1
    $colsInput.Maximum = 12
    $colsInput.Value = 3
    $dialogBody.Controls.Add($colsInput)

    $diagLabel = New-Object System.Windows.Forms.Label
    $diagLabel.Text = $ui.DiagValues
    $diagLabel.Location = New-Object System.Drawing.Point(20, 190)
    $diagLabel.Size = New-Object System.Drawing.Size(80, 24)
    $dialogBody.Controls.Add($diagLabel)

    $diagTextBox = New-Object System.Windows.Forms.TextBox
    $diagTextBox.Location = New-Object System.Drawing.Point(112, 186)
    $diagTextBox.Size = New-Object System.Drawing.Size(328, 30)
    $diagTextBox.Anchor = 'Top,Left,Right'
    $diagTextBox.Text = '1,2,3'
    $dialogBody.Controls.Add($diagTextBox)

    $previewLabel = New-Object System.Windows.Forms.Label
    $previewLabel.Text = $ui.SpecialHint
    $previewLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#647A90')
    $previewLabel.Location = New-Object System.Drawing.Point(20, 228)
    $previewLabel.Size = New-Object System.Drawing.Size(420, 34)
    $previewLabel.Anchor = 'Top,Left,Right'
    $dialogBody.Controls.Add($previewLabel)

    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = $ui.Apply
    $applyButton.Location = New-Object System.Drawing.Point(264, 10)
    $applyButton.Size = New-Object System.Drawing.Size(80, 32)
    $applyButton.Anchor = 'Top,Right'
    Set-ButtonStyle -Button $applyButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
    $buttonPanel.Controls.Add($applyButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = $ui.Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(352, 10)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 32)
    $cancelButton.Anchor = 'Top,Right'
    Set-ButtonStyle -Button $cancelButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $buttonPanel.Controls.Add($cancelButton)

    $dialog.AcceptButton = $applyButton
    $dialog.CancelButton = $cancelButton

    $updateControlState = {
        $selectedType = [string]$typeCombo.SelectedItem
        $isIdentity = ($selectedType -eq $ui.Identity)
        $isRectangular = ($selectedType -eq $ui.ZeroMatrix -or $selectedType -eq $ui.OnesMatrix)
        $isDiag = ($selectedType -eq $ui.DiagMatrix)

        $sizeLabel.Visible = $isIdentity
        $sizeInput.Visible = $isIdentity
        $rowsLabel.Visible = $isRectangular
        $rowsInput.Visible = $isRectangular
        $colsLabel.Visible = $isRectangular
        $colsInput.Visible = $isRectangular
        $diagLabel.Visible = $isDiag
        $diagTextBox.Visible = $isDiag
    }
    $typeCombo.Add_SelectedIndexChanged($updateControlState)
    & $updateControlState

    $buttonPanel.Add_Resize({
        $cancelButton.Left = $buttonPanel.ClientSize.Width - $cancelButton.Width - 8
        $applyButton.Left = $cancelButton.Left - $applyButton.Width - 8
    })
    $buttonPanel.PerformLayout()
    $cancelButton.Left = $buttonPanel.ClientSize.Width - $cancelButton.Width - 8
    $applyButton.Left = $cancelButton.Left - $applyButton.Width - 8

    $applyButton.Add_Click({
        try {
            $selectedType = [string]$typeCombo.SelectedItem
            $textValue = ''
            $descriptor = ''
            switch ($selectedType) {
                { $_ -eq $ui.Identity } {
                    $n = [int]$sizeInput.Value
                    $descriptor = 'I' + $n
                    $textValue = $descriptor
                }
                { $_ -eq $ui.ZeroMatrix } {
                    $r = [int]$rowsInput.Value
                    $c = [int]$colsInput.Value
                    $descriptor = 'zeros(' + $r + ',' + $c + ')'
                    $textValue = $descriptor
                }
                { $_ -eq $ui.OnesMatrix } {
                    $r = [int]$rowsInput.Value
                    $c = [int]$colsInput.Value
                    $descriptor = 'ones(' + $r + ',' + $c + ')'
                    $textValue = $descriptor
                }
                { $_ -eq $ui.DiagMatrix } {
                    $parts = @((Normalize-CommonSymbols $diagTextBox.Text) -split '[,;\s]+' | Where-Object { $_ -ne '' })
                    if ($parts.Count -eq 0) { throw $ui.SpecialMatrixInvalid }
                    $descriptor = 'diag(' + ($parts -join ',') + ')'
                    $textValue = $descriptor
                }
                default { throw $ui.SpecialMatrixInvalid }
            }

            $Item.SuppressModeReset = $true
            try {
                $Item.TextBox.Text = $textValue
            }
            finally {
                $Item.SuppressModeReset = $false
            }
            Set-MatrixCardMode -Item $Item -Mode 'preset' -Descriptor $descriptor
            Refresh-VariableQuickButtons
            $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dialog.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.SpecialConfig, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$dialog.ShowDialog($form)
    $dialog.Dispose()
}

function Show-HistoryWindow {
    if ($null -eq $script:historyForm -or $script:historyForm.IsDisposed) {

        Set-FormScaleDefaults -Form $script:historyForm
        $script:historyForm.Text = $ui.HistoryWindow
        $script:historyForm.Size = New-Object System.Drawing.Size(920, 560)
        $script:historyForm.StartPosition = 'CenterParent'
        $script:historyForm.BackColor = [System.Drawing.Color]::White
        $script:historyForm.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

        $historySplit = New-Object System.Windows.Forms.SplitContainer
        $historySplit.Dock = 'Fill'
        $historySplit.SplitterDistance = 560
        $script:historyForm.Controls.Add($historySplit)

        $historyGrid = New-Object System.Windows.Forms.DataGridView
        $historyGrid.Dock = 'Fill'
        $historyGrid.ReadOnly = $true
        $historyGrid.AllowUserToAddRows = $false
        $historyGrid.AllowUserToDeleteRows = $false
        $historyGrid.SelectionMode = 'FullRowSelect'
        $historyGrid.MultiSelect = $false
        $historyGrid.AutoSizeColumnsMode = 'None'
        $historyGrid.ColumnHeadersHeightSizeMode = 'AutoSize'
        $historyGrid.RowHeadersVisible = $false
        $historySplit.Panel1.Controls.Add($historyGrid)

        $detailPanel = New-Object System.Windows.Forms.Panel
        $detailPanel.Dock = 'Fill'
        $historySplit.Panel2.Controls.Add($detailPanel)

        $detailTitle = New-Object System.Windows.Forms.Label
        $detailTitle.Text = $ui.HistoryWindow
        $detailTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
        $detailTitle.Location = New-Object System.Drawing.Point(14, 14)
        $detailTitle.Size = New-Object System.Drawing.Size(160, 24)
        $detailPanel.Controls.Add($detailTitle)

        $loadButton = New-Object System.Windows.Forms.Button
        $loadButton.Text = $ui.LoadExpression
        $loadButton.Location = New-Object System.Drawing.Point(260, 10)
        $loadButton.Size = New-Object System.Drawing.Size(116, 34)
        $loadButton.Anchor = 'Top,Right'
        Set-ButtonStyle -Button $loadButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
        $detailPanel.Controls.Add($loadButton)

        $detailBox = New-Object System.Windows.Forms.TextBox
        $detailBox.Location = New-Object System.Drawing.Point(14, 54)
        $detailBox.Size = New-Object System.Drawing.Size(362, 432)
        $detailBox.Anchor = 'Top,Bottom,Left,Right'
        $detailBox.Multiline = $true
        $detailBox.ReadOnly = $true
        $detailBox.ScrollBars = 'Vertical'
        $detailBox.Font = New-Object System.Drawing.Font('Consolas', 11)
        $detailPanel.Controls.Add($detailBox)

        $refreshHistory = {
            $historyTable = New-Object System.Data.DataTable 'OperationHistory'
            [void]$historyTable.Columns.Add('Time', [string])
            [void]$historyTable.Columns.Add('Status', [string])
            [void]$historyTable.Columns.Add('Expression', [string])
            foreach ($entry in $script:operationHistory) {
                $row = $historyTable.NewRow()
                $row['Time'] = $entry.Time
                $row['Status'] = $entry.Status
                $row['Expression'] = $entry.Expression
                [void]$historyTable.Rows.Add($row)
            }
            $historyGrid.DataSource = $historyTable
            if ($historyGrid.Columns.Contains('Time')) {
                $historyGrid.Columns['Time'].Width = 170
                $historyGrid.Columns['Time'].MinimumWidth = 150
            }
            if ($historyGrid.Columns.Contains('Status')) {
                $historyGrid.Columns['Status'].Width = 210
                $historyGrid.Columns['Status'].MinimumWidth = 180
            }
            if ($historyGrid.Columns.Contains('Expression')) {
                $historyGrid.Columns['Expression'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
                $historyGrid.Columns['Expression'].MinimumWidth = 220
            }
            if ($historyTable.Rows.Count -eq 0) {
                $detailBox.Text = $ui.HistoryEmpty
            } else {
                $historyGrid.ClearSelection()
                $historyGrid.Rows[0].Selected = $true
                $detailBox.Text = "Time: " + $script:operationHistory[0].Time + "`r`nStatus: " + $script:operationHistory[0].Status + "`r`n`r`nExpression:`r`n" + $script:operationHistory[0].Expression + "`r`n`r`nResult:`r`n" + $script:operationHistory[0].Result
            }
        }

        $historyGrid.Add_SelectionChanged({
            if ($historyGrid.SelectedRows.Count -eq 0) { return }
            $index = $historyGrid.SelectedRows[0].Index
            if ($index -ge $script:operationHistory.Count) { return }
            $entry = $script:operationHistory[$index]
            $detailBox.Text = "Time: " + $entry.Time + "`r`nStatus: " + $entry.Status + "`r`n`r`nExpression:`r`n" + $entry.Expression + "`r`n`r`nResult:`r`n" + $entry.Result
        })

        $loadButton.Add_Click({
            if ($historyGrid.SelectedRows.Count -eq 0) { return }
            $index = $historyGrid.SelectedRows[0].Index
            if ($index -ge $script:operationHistory.Count) { return }
            $entry = $script:operationHistory[$index]
            $expressionTextBox.Text = $entry.Expression
            $resultTextBox.Text = $entry.Result
            $statusLabel.Text = $entry.Status
            $script:historyForm.Hide()
            $form.Activate()
        })

        $script:historyForm.Add_Activated($refreshHistory)
        & $refreshHistory
    }
    $script:historyForm.Show($form)
    $script:historyForm.BringToFront()
}

function ConvertTo-CsvCell {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ($text.Contains('"')) { $text = $text.Replace('"', '""') }
    if ($text.IndexOfAny(@([char]',', [char]'"', [char]"`r", [char]"`n")) -ge 0) { return '"' + $text + '"' }
    return $text
}

function Export-DataTableToCsvFile {
    param([System.Data.DataTable]$Table, [string]$Path)
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $headers = @($Table.Columns | ForEach-Object { ConvertTo-CsvCell $_.ColumnName })
    $lines.Add(($headers -join ','))
    foreach ($row in $Table.Rows) {
        $cells = foreach ($column in $Table.Columns) { ConvertTo-CsvCell ([string]$row[$column.ColumnName]) }
        $lines.Add(($cells -join ','))
    }
    [System.IO.File]::WriteAllLines($Path, $lines, [System.Text.Encoding]::UTF8)
}

function Export-ControlToImageFile {
    param([System.Windows.Forms.Control]$Control, [string]$Path)
    $bitmap = New-Object System.Drawing.Bitmap($Control.Width, $Control.Height)
    $Control.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $Control.Width, $Control.Height)))
    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

function Show-ImportDataDialog {

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.ImportFile
    $dialog.Size = New-Object System.Drawing.Size(560, 250)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Text = $ui.ImportFile
    $pathLabel.Location = New-Object System.Drawing.Point(18, 22)
    $pathLabel.Size = New-Object System.Drawing.Size(90, 24)
    $dialog.Controls.Add($pathLabel)

    $pathBox = New-Object System.Windows.Forms.TextBox
    $pathBox.Location = New-Object System.Drawing.Point(112, 18)
    $pathBox.Size = New-Object System.Drawing.Size(320, 30)
    $dialog.Controls.Add($pathBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = $ui.Browse
    $browseButton.Location = New-Object System.Drawing.Point(440, 16)
    $browseButton.Size = New-Object System.Drawing.Size(80, 34)
    Set-ButtonStyle -Button $browseButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($browseButton)

    $sheetLabel = New-Object System.Windows.Forms.Label
    $sheetLabel.Text = $ui.SheetName
    $sheetLabel.Location = New-Object System.Drawing.Point(18, 66)
    $sheetLabel.Size = New-Object System.Drawing.Size(90, 24)
    $dialog.Controls.Add($sheetLabel)

    $sheetBox = New-Object System.Windows.Forms.TextBox
    $sheetBox.Location = New-Object System.Drawing.Point(112, 62)
    $sheetBox.Size = New-Object System.Drawing.Size(160, 30)
    $dialog.Controls.Add($sheetBox)

    $headerCheck = New-Object System.Windows.Forms.CheckBox
    $headerCheck.Text = $ui.FirstRowHeader
    $headerCheck.Location = New-Object System.Drawing.Point(290, 64)
    $headerCheck.Size = New-Object System.Drawing.Size(160, 26)
    $dialog.Controls.Add($headerCheck)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = $ui.ImportAllHint
    $hint.Location = New-Object System.Drawing.Point(18, 110)
    $hint.Size = New-Object System.Drawing.Size(500, 44)
    $hint.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#647A90')
    $dialog.Controls.Add($hint)

    $importButton = New-Object System.Windows.Forms.Button
    $importButton.Text = $ui.ImportFile
    $importButton.Location = New-Object System.Drawing.Point(350, 166)
    $importButton.Size = New-Object System.Drawing.Size(80, 34)
    Set-ButtonStyle -Button $importButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
    $dialog.Controls.Add($importButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = $ui.Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(440, 166)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 34)
    Set-ButtonStyle -Button $cancelButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($cancelButton)

    $browseButton.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = 'Data Files|*.csv;*.txt;*.xlsx;*.xls;*.xlsm|All Files|*.*'
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $pathBox.Text = $fileDialog.FileName
        }
    })

    $importButton.Add_Click({
        try {
            $table = Import-DataTableFromFile -Path $pathBox.Text.Trim() -FirstRowAsHeader:$headerCheck.Checked -SheetName $sheetBox.Text.Trim()
            $script:dataPath = $pathBox.Text.Trim()
            Set-CurrentDataTable -Table $table -SourceLabel ([System.IO.Path]::GetFileName($script:dataPath))
            $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dialog.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.DataImportFailed, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$dialog.ShowDialog($script:dataForm)
    $dialog.Dispose()
}

function Show-CleanMissingDialog {
    Ensure-DataLoaded

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.CleanMissing
    $dialog.Size = New-Object System.Drawing.Size(360, 240)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = $ui.RemoveMissingRows
    $removeButton.Location = New-Object System.Drawing.Point(24, 30)
    $removeButton.Size = New-Object System.Drawing.Size(290, 38)
    Set-ButtonStyle -Button $removeButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
    $dialog.Controls.Add($removeButton)

    $zeroButton = New-Object System.Windows.Forms.Button
    $zeroButton.Text = $ui.FillMissingZero
    $zeroButton.Location = New-Object System.Drawing.Point(24, 82)
    $zeroButton.Size = New-Object System.Drawing.Size(290, 38)
    Set-ButtonStyle -Button $zeroButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($zeroButton)

    $meanButton = New-Object System.Windows.Forms.Button
    $meanButton.Text = $ui.FillMissingMean
    $meanButton.Location = New-Object System.Drawing.Point(24, 134)
    $meanButton.Size = New-Object System.Drawing.Size(290, 38)
    Set-ButtonStyle -Button $meanButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F8EFD9')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#74511A'))
    $dialog.Controls.Add($meanButton)

    $removeButton.Add_Click({ Set-CurrentDataTable -Table (Remove-MissingRowsFromTable $script:dataTable) -SourceLabel $script:dataFileLabel.Text; $dialog.Close() })
    $zeroButton.Add_Click({ Set-CurrentDataTable -Table (Fill-MissingWithZero $script:dataTable) -SourceLabel $script:dataFileLabel.Text; $dialog.Close() })
    $meanButton.Add_Click({ Set-CurrentDataTable -Table (Fill-MissingWithMean $script:dataTable) -SourceLabel $script:dataFileLabel.Text; $dialog.Close() })
    [void]$dialog.ShowDialog($script:dataForm)
    $dialog.Dispose()
}

function Show-SelectRowsColsDialog {
    Ensure-DataLoaded

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.SelectRowsCols
    $dialog.Size = New-Object System.Drawing.Size(420, 240)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $labels = @($ui.StartRow, $ui.EndRow, $ui.StartCol, $ui.EndCol)
    $controls = @()
    for ($i = 0; $i -lt 4; $i++) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labels[$i]
        $label.Location = New-Object System.Drawing.Point -ArgumentList 24, (24 + ($i * 38))
        $label.Size = New-Object System.Drawing.Size(100, 24)
        $dialog.Controls.Add($label)
        $num = New-Object System.Windows.Forms.NumericUpDown
        $num.Location = New-Object System.Drawing.Point -ArgumentList 130, (20 + ($i * 38))
        $num.Minimum = 0
        $num.Maximum = 999999
        $num.Value = if ($i -lt 2) { 1 } else { 1 }
        $dialog.Controls.Add($num)
        $controls += ,$num
    }
    $controls[1].Value = 0
    $controls[3].Value = 0

    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = $ui.Apply
    $applyButton.Location = New-Object System.Drawing.Point(224, 170)
    $applyButton.Size = New-Object System.Drawing.Size(74, 32)
    Set-ButtonStyle -Button $applyButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
    $dialog.Controls.Add($applyButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = $ui.Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(306, 170)
    $cancelButton.Size = New-Object System.Drawing.Size(74, 32)
    Set-ButtonStyle -Button $cancelButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($cancelButton)

    $applyButton.Add_Click({
        $selected = Select-TableRange -Table $script:dataTable -StartRow ([int]$controls[0].Value) -EndRow ([int]$controls[1].Value) -StartCol ([int]$controls[2].Value) -EndCol ([int]$controls[3].Value)
        Set-CurrentDataTable -Table $selected -SourceLabel $script:dataFileLabel.Text
        $dialog.Close()
    })
    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$dialog.ShowDialog($script:dataForm)
    $dialog.Dispose()
}

function Show-EditDataDialog {
    Ensure-DataLoaded

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.EditData
    $dialog.Size = New-Object System.Drawing.Size(420, 240)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $rowLabel = New-Object System.Windows.Forms.Label
    $rowLabel.Text = $ui.Rows
    $rowLabel.Location = New-Object System.Drawing.Point(24, 28)
    $dialog.Controls.Add($rowLabel)
    $rowInput = New-Object System.Windows.Forms.NumericUpDown
    $rowInput.Location = New-Object System.Drawing.Point(110, 24)
    $rowInput.Minimum = 1
    $rowInput.Maximum = [Math]::Max(1, $script:dataTable.Rows.Count)
    $dialog.Controls.Add($rowInput)

    $columnLabel = New-Object System.Windows.Forms.Label
    $columnLabel.Text = $ui.Cols
    $columnLabel.Location = New-Object System.Drawing.Point(24, 74)
    $dialog.Controls.Add($columnLabel)
    $columnCombo = New-Object System.Windows.Forms.ComboBox
    $columnCombo.Location = New-Object System.Drawing.Point(110, 70)
    $columnCombo.Size = New-Object System.Drawing.Size(260, 30)
    $columnCombo.DropDownStyle = 'DropDownList'
    [void]$columnCombo.Items.AddRange((Get-DataTableColumnNames $script:dataTable))
    if ($columnCombo.Items.Count -gt 0) { $columnCombo.SelectedIndex = 0 }
    $dialog.Controls.Add($columnCombo)

    $valueLabel = New-Object System.Windows.Forms.Label
    $valueLabel.Text = $ui.EditData
    $valueLabel.Location = New-Object System.Drawing.Point(24, 120)
    $dialog.Controls.Add($valueLabel)
    $valueBox = New-Object System.Windows.Forms.TextBox
    $valueBox.Location = New-Object System.Drawing.Point(110, 116)
    $valueBox.Size = New-Object System.Drawing.Size(260, 30)
    $dialog.Controls.Add($valueBox)

    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = $ui.Apply
    $applyButton.Location = New-Object System.Drawing.Point(214, 168)
    $applyButton.Size = New-Object System.Drawing.Size(74, 32)
    Set-ButtonStyle -Button $applyButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
    $dialog.Controls.Add($applyButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = $ui.Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(296, 168)
    $cancelButton.Size = New-Object System.Drawing.Size(74, 32)
    Set-ButtonStyle -Button $cancelButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($cancelButton)

    $applyButton.Add_Click({
        $updated = Set-DataCellValue -Table $script:dataTable -RowIndex ([int]$rowInput.Value) -ColumnName ([string]$columnCombo.SelectedItem) -NewValue $valueBox.Text
        Set-CurrentDataTable -Table $updated -SourceLabel $script:dataFileLabel.Text
        $dialog.Close()
    })
    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$dialog.ShowDialog($script:dataForm)
    $dialog.Dispose()
}

function Show-GroupStatsDialog {
    Ensure-DataLoaded

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.GroupStats
    $dialog.Size = New-Object System.Drawing.Size(440, 250)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $columns = Get-DataTableColumnNames $script:dataTable
    $groupLabel = New-Object System.Windows.Forms.Label
    $groupLabel.Text = $ui.GroupBy
    $groupLabel.Location = New-Object System.Drawing.Point(24, 26)
    $dialog.Controls.Add($groupLabel)
    $groupCombo = New-Object System.Windows.Forms.ComboBox
    $groupCombo.Location = New-Object System.Drawing.Point(120, 22)
    $groupCombo.Size = New-Object System.Drawing.Size(270, 30)
    $groupCombo.DropDownStyle = 'DropDownList'
    [void]$groupCombo.Items.AddRange($columns)
    if ($groupCombo.Items.Count -gt 0) { $groupCombo.SelectedIndex = 0 }
    $dialog.Controls.Add($groupCombo)

    $valueLabel = New-Object System.Windows.Forms.Label
    $valueLabel.Text = $ui.AggregateColumn
    $valueLabel.Location = New-Object System.Drawing.Point(24, 78)
    $dialog.Controls.Add($valueLabel)
    $valueCombo = New-Object System.Windows.Forms.ComboBox
    $valueCombo.Location = New-Object System.Drawing.Point(120, 74)
    $valueCombo.Size = New-Object System.Drawing.Size(270, 30)
    $valueCombo.DropDownStyle = 'DropDownList'
    [void]$valueCombo.Items.AddRange($columns)
    if ($valueCombo.Items.Count -gt 0) { $valueCombo.SelectedIndex = 0 }
    $dialog.Controls.Add($valueCombo)

    $methodLabel = New-Object System.Windows.Forms.Label
    $methodLabel.Text = $ui.AggregateMethod
    $methodLabel.Location = New-Object System.Drawing.Point(24, 130)
    $dialog.Controls.Add($methodLabel)
    $methodCombo = New-Object System.Windows.Forms.ComboBox
    $methodCombo.Location = New-Object System.Drawing.Point(120, 126)
    $methodCombo.Size = New-Object System.Drawing.Size(150, 30)
    $methodCombo.DropDownStyle = 'DropDownList'
    [void]$methodCombo.Items.AddRange(@('COUNT', 'SUM', 'MEAN'))
    $methodCombo.SelectedIndex = 0
    $dialog.Controls.Add($methodCombo)

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = $ui.GroupStats
    $runButton.Location = New-Object System.Drawing.Point(220, 176)
    $runButton.Size = New-Object System.Drawing.Size(80, 32)
    Set-ButtonStyle -Button $runButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
    $dialog.Controls.Add($runButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = $ui.Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(310, 176)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 32)
    Set-ButtonStyle -Button $cancelButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
    $dialog.Controls.Add($cancelButton)

    $runButton.Add_Click({
        $result = Get-GroupedStatisticsTable -Table $script:dataTable -GroupColumn ([string]$groupCombo.SelectedItem) -ValueColumn ([string]$valueCombo.SelectedItem) -Method ([string]$methodCombo.SelectedItem)
        Show-DataTableDialog -Title $ui.GroupStats -Table $result
    })
    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$dialog.ShowDialog($script:dataForm)
    $dialog.Dispose()
}

function Show-ExportDataDialog {
    Ensure-DataLoaded

    Set-FormScaleDefaults -Form $dialog
    $dialog.Text = $ui.ExportData
    $dialog.Size = New-Object System.Drawing.Size(360, 220)
    $dialog.StartPosition = 'CenterParent'
    $dialog.BackColor = [System.Drawing.Color]::White
    $dialog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $csvButton = New-Object System.Windows.Forms.Button
    $csvButton.Text = $ui.ExportCsv
    $csvButton.Location = New-Object System.Drawing.Point(24, 34)
    $csvButton.Size = New-Object System.Drawing.Size(290, 40)
    Set-ButtonStyle -Button $csvButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
    $dialog.Controls.Add($csvButton)

    $imgButton = New-Object System.Windows.Forms.Button
    $imgButton.Text = $ui.ExportImage
    $imgButton.Location = New-Object System.Drawing.Point(24, 92)
    $imgButton.Size = New-Object System.Drawing.Size(290, 40)
    Set-ButtonStyle -Button $imgButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F8EFD9')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#74511A'))
    $dialog.Controls.Add($imgButton)

    $csvButton.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = 'CSV Files|*.csv'
        $saveDialog.FileName = 'table_export.csv'
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-DataTableToCsvFile -Table $script:dataTable -Path $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show($ui.DataExported, $ui.ExportData, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

    $imgButton.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = 'PNG Files|*.png'
        $saveDialog.FileName = 'table_snapshot.png'
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-ControlToImageFile -Control $script:dataGridHost -Path $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show($ui.DataExported, $ui.ExportData, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

    [void]$dialog.ShowDialog($script:dataForm)
    $dialog.Dispose()
}

function Show-DataWindow {
    if ($null -eq $script:dataForm -or $script:dataForm.IsDisposed) {

        Set-FormScaleDefaults -Form $script:dataForm
        $script:dataForm.Text = $ui.DataWindow
        $script:dataForm.Size = New-Object System.Drawing.Size(1220, 720)
        $script:dataForm.StartPosition = 'CenterParent'
        $script:dataForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#EFF3F8')
        $script:dataForm.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

        $rootPanel = New-Object System.Windows.Forms.Panel
        $rootPanel.Dock = 'Fill'
        $rootPanel.Padding = New-Object System.Windows.Forms.Padding(12)
        $script:dataForm.Controls.Add($rootPanel)

        $topCard = New-Object System.Windows.Forms.Panel
        $topCard.Dock = 'Top'
        $topCard.Height = 196
        $topCard.BackColor = [System.Drawing.Color]::White
        $topCard.BorderStyle = 'FixedSingle'
        $topCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
        $rootPanel.Controls.Add($topCard)

        $script:dataGridHost = New-Object System.Windows.Forms.Panel
        $script:dataGridHost.Dock = 'Fill'
        $script:dataGridHost.Padding = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
        $rootPanel.Controls.Add($script:dataGridHost)

        $title = New-Object System.Windows.Forms.Label
        $title.Text = $ui.DataWindow
        $title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 12, [System.Drawing.FontStyle]::Bold)
        $title.Location = New-Object System.Drawing.Point(16, 12)
        $title.Size = New-Object System.Drawing.Size(280, 24)
        $topCard.Controls.Add($title)

        $script:dataFileLabel = New-Object System.Windows.Forms.Label
        $script:dataFileLabel.Text = $ui.DataStatusReady
        $script:dataFileLabel.Location = New-Object System.Drawing.Point(18, 42)
        $script:dataFileLabel.Size = New-Object System.Drawing.Size(1080, 22)
        $script:dataFileLabel.Anchor = 'Top,Left,Right'
        $topCard.Controls.Add($script:dataFileLabel)

        $script:dataStatusLabel = New-Object System.Windows.Forms.Label
        $script:dataStatusLabel.Text = $ui.DataStatusReady
        $script:dataStatusLabel.Location = New-Object System.Drawing.Point(18, 68)
        $script:dataStatusLabel.Size = New-Object System.Drawing.Size(1080, 22)
        $script:dataStatusLabel.Anchor = 'Top,Left,Right'
        $topCard.Controls.Add($script:dataStatusLabel)
        $buttonHost = New-Object System.Windows.Forms.Panel
        $buttonHost.Location = New-Object System.Drawing.Point(18, 102)
        $buttonHost.Size = New-Object System.Drawing.Size(1132, 84)
        $buttonHost.Anchor = 'Top,Left,Right'
        $buttonHost.AutoScroll = $true
        $buttonHost.BorderStyle = 'None'
        $topCard.Controls.Add($buttonHost)

        $buttonFlow = New-Object System.Windows.Forms.FlowLayoutPanel
        $buttonFlow.Location = New-Object System.Drawing.Point(0, 0)
        $buttonFlow.Size = New-Object System.Drawing.Size(1114, 76)
        $buttonFlow.Anchor = 'Top,Left,Right'
        $buttonFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        $buttonFlow.WrapContents = $true
        $buttonFlow.AutoScroll = $false
        $buttonFlow.Padding = New-Object System.Windows.Forms.Padding(0)
        $buttonFlow.AutoSize = $true
        $buttonFlow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
        $buttonHost.Controls.Add($buttonFlow)
        function New-DataActionButton {
            param([string]$Text, [System.Drawing.Color]$BackColor, [System.Drawing.Color]$ForeColor)
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text = $Text
            $btn.Size = New-Object System.Drawing.Size(126, 32)
            $btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 8)
            Set-ButtonStyle -Button $btn -BackColor $BackColor -ForeColor $ForeColor
            return $btn
        }

        $importButton = New-DataActionButton -Text $ui.ImportFile -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
        $overviewButton = New-DataActionButton -Text $ui.DataOverview -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
        $missingButton = New-DataActionButton -Text $ui.MissingRatio -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
        $cleanButton = New-DataActionButton -Text $ui.CleanMissing -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
        $selectButton = New-DataActionButton -Text $ui.SelectRowsCols -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
        $editButton = New-DataActionButton -Text $ui.EditData -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
        $groupButton = New-DataActionButton -Text $ui.GroupStats -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F8EFD9')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#74511A'))
        $plotButton = New-DataActionButton -Text $ui.PlotData -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F8EFD9')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#74511A'))
        $exportButton = New-DataActionButton -Text $ui.ExportData -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
        foreach ($button in @($importButton, $overviewButton, $missingButton, $cleanButton, $selectButton, $editButton, $groupButton, $plotButton, $exportButton)) { $buttonFlow.Controls.Add($button) }

        $updateDataTopLayout = {
            $buttonHost.Width = [Math]::Max(320, $topCard.ClientSize.Width - 36)
            $buttonFlow.Width = [Math]::Max(300, $buttonHost.ClientSize.Width - 18)
            $buttonHost.Height = [Math]::Min(132, [Math]::Max(52, $buttonFlow.PreferredSize.Height + 12))
            $buttonHost.AutoScrollMinSize = [System.Drawing.Size]::new([Math]::Max($buttonHost.ClientSize.Width - 6, $buttonFlow.PreferredSize.Width + 8), [Math]::Max(44, $buttonFlow.PreferredSize.Height + 8))
            $topCard.Height = [Math]::Max(194, $buttonHost.Bottom + 14)
        }

        $topCard.Add_Resize({
            & $updateDataTopLayout
        })

        & $updateDataTopLayout

        $gridCard = New-Object System.Windows.Forms.Panel
        $gridCard.Dock = 'Fill'
        $gridCard.BackColor = [System.Drawing.Color]::White
        $gridCard.BorderStyle = 'FixedSingle'
        $gridCard.Padding = New-Object System.Windows.Forms.Padding(10)
        $script:dataGridHost.Controls.Add($gridCard)

        $script:dataGridView = New-Object System.Windows.Forms.DataGridView
        $script:dataGridView.Dock = 'Fill'
        $script:dataGridView.AllowUserToAddRows = $false
        $script:dataGridView.AllowUserToDeleteRows = $false
        $script:dataGridView.AutoSizeColumnsMode = 'DisplayedCells'
        $script:dataGridView.RowHeadersVisible = $true
        $script:dataGridView.BackgroundColor = [System.Drawing.Color]::White
        $script:dataGridView.BorderStyle = 'None'
        $script:dataGridView.ColumnHeadersHeightSizeMode = 'AutoSize'
        $gridCard.Controls.Add($script:dataGridView)

        $importButton.Add_Click({ Show-ImportDataDialog })
        $overviewButton.Add_Click({
            try { Show-TextViewerDialog -Title $ui.DataOverview -Text (Get-DataOverviewReportText $script:dataTable) } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.DataOverview, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null }
        })
        $missingButton.Add_Click({
            try { Show-MissingRatioDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.MissingRatio, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null }
        })
        $cleanButton.Add_Click({ try { Show-CleanMissingDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.CleanMissing, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } })
        $selectButton.Add_Click({ try { Show-SelectRowsColsDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.SelectRowsCols, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } })
        $editButton.Add_Click({ try { Show-EditDataDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.EditData, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } })
        $groupButton.Add_Click({ try { Show-GroupStatsDialogEx } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.GroupStats, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } })
        $plotButton.Add_Click({ try { Show-PlotCenterWindow } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.PlotData, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } })
        $exportButton.Add_Click({ try { Show-ExportDataDialog } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.ExportData, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } })
    }

    if ($null -eq $script:dataTable) { Set-CurrentDataTable -Table $null -SourceLabel '' }
    $script:dataForm.Show($form)
    $script:dataForm.BringToFront()
}

function New-MatrixCard {
    param([string]$Label, [string]$InitialText)

    $card = New-Object System.Windows.Forms.Panel
    $card.Size = New-Object System.Drawing.Size(430, 388)
    $card.Margin = New-Object System.Windows.Forms.Padding(0, 0, 14, 14)
    $card.BackColor = [System.Drawing.Color]::White
    $card.BorderStyle = 'FixedSingle'

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Matrix ' + $Label
    $title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(14, 12)
    $title.Size = New-Object System.Drawing.Size(140, 24)
    $card.Controls.Add($title)

    $specialButton = New-Object System.Windows.Forms.Button
    $specialButton.Text = $ui.SpecialButtonText
    $specialButton.Location = New-Object System.Drawing.Point(342, 10)
    $specialButton.Size = New-Object System.Drawing.Size(70, 28)
    Set-ButtonStyle -Button $specialButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
    $card.Controls.Add($specialButton)

    $modeLabel = New-Object System.Windows.Forms.Label
    $modeLabel.Text = $ui.InputMode
    $modeLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#72849A')
    $modeLabel.Location = New-Object System.Drawing.Point(14, 44)
    $modeLabel.Size = New-Object System.Drawing.Size(70, 22)
    $card.Controls.Add($modeLabel)

    $modeValue = New-Object System.Windows.Forms.Label
    $modeValue.Text = $ui.ManualInput
    $modeValue.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#4F6479')
    $modeValue.Location = New-Object System.Drawing.Point(88, 44)
    $modeValue.Size = New-Object System.Drawing.Size(324, 22)
    $card.Controls.Add($modeValue)

    $tip = New-Object System.Windows.Forms.Label
    $tip.Text = $ui.ManualEditHint
    $tip.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $tip.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#72849A')
    $tip.Location = New-Object System.Drawing.Point(14, 70)
    $tip.Size = New-Object System.Drawing.Size(398, 62)
    $card.Controls.Add($tip)

    $blockTip = New-Object System.Windows.Forms.Label
    $blockTip.Text = $ui.BlockHint
    $blockTip.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $blockTip.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#72849A')
    $blockTip.Location = New-Object System.Drawing.Point(14, 132)
    $blockTip.Size = New-Object System.Drawing.Size(398, 40)
    $card.Controls.Add($blockTip)

    $specialTip = New-Object System.Windows.Forms.Label
    $specialTip.Text = $ui.SpecialHint
    $specialTip.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $specialTip.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#72849A')
    $specialTip.Location = New-Object System.Drawing.Point(14, 174)
    $specialTip.Size = New-Object System.Drawing.Size(398, 42)
    $card.Controls.Add($specialTip)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Vertical'
    $textBox.BorderStyle = 'FixedSingle'
    $textBox.Font = New-Object System.Drawing.Font('Consolas', 12)
    $textBox.Location = New-Object System.Drawing.Point(14, 226)
    $textBox.Size = New-Object System.Drawing.Size(398, 144)
    $textBox.Anchor = 'Top,Bottom,Left,Right'
    $textBox.Text = $InitialText
    $card.Controls.Add($textBox)

    $item = @{
        Label = $Label
        Card = $card
        TitleLabel = $title
        ModeValueLabel = $modeValue
        TextBox = $textBox
        SpecialButton = $specialButton
        Mode = 'manual'
        Descriptor = ''
        SuppressModeReset = $false
    }

    $state = $item
    $textBox.Add_TextChanged({
        if (-not $state.SuppressModeReset -and $state.Mode -ne 'manual') {
            Set-MatrixCardMode -Item $state -Mode 'manual' -Descriptor ''
            Refresh-VariableQuickButtons
        }
    }.GetNewClosure())
    $specialButton.Add_Click({ Show-SpecialMatrixDialog -Item $state }.GetNewClosure())

    return $item
}

function Add-MatrixCard {
    param([string]$InitialText = '', [switch]$OpenSpecial)
    if ($matrixStore.Count -ge 8) { return }
    $label = [string](Get-MatrixLabelByIndex $matrixStore.Count)
    $item = New-MatrixCard -Label $label -InitialText $InitialText
    [void]$matrixStore.Add($item)
    Refresh-MatrixCards
    if ($OpenSpecial) { Show-SpecialMatrixDialog -Item $item }
}

function Remove-MatrixCard {
    if ($matrixStore.Count -le 1) { return }
    $matrixStore.RemoveAt($matrixStore.Count - 1)
    Refresh-MatrixCards
}

function Add-ExpressionQuickButton {
    param([string]$Text, [string]$InsertText, [string]$BackHex, [string]$ForeHex)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Size = New-Object System.Drawing.Size(102, 32)
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 8)
    Set-ButtonStyle -Button $button -BackColor ([System.Drawing.ColorTranslator]::FromHtml($BackHex)) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml($ForeHex))
    $insertValue = $InsertText
    $button.Add_Click({ Insert-ExpressionText -TextBox $expressionTextBox -Text $insertValue }.GetNewClosure())
    $script:functionQuickPanel.Controls.Add($button)
}

$form = New-Object System.Windows.Forms.Form
Set-FormScaleDefaults -Form $form
$form.Text = $ui.Title
$form.Size = New-Object System.Drawing.Size(1460, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1180, 760)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#EFF3F8')
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
Set-FormScaleDefaults -Form $form

$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = New-Object System.Windows.Forms.Padding(12, 12, 12, 12)
$form.Controls.Add($contentPanel)

$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.SplitterWidth = 6
$mainSplit.SplitterDistance = 840
$mainSplit.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#D8E2ED')
$contentPanel.Controls.Add($mainSplit)

$toolbarCard = New-Object System.Windows.Forms.Panel
$toolbarCard.Dock = 'Top'
$toolbarCard.Height = 94
$toolbarCard.BackColor = [System.Drawing.Color]::White
$toolbarCard.BorderStyle = 'FixedSingle'
$contentPanel.Controls.Add($toolbarCard)

$contentPanel.Controls.SetChildIndex($toolbarCard, 0)

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Text = $ui.MatrixCount
$countLabel.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
$countLabel.Location = New-Object System.Drawing.Point(14, 18)
$countLabel.Size = New-Object System.Drawing.Size(110, 24)
$toolbarCard.Controls.Add($countLabel)

$countValueLabel = New-Object System.Windows.Forms.Label
$countValueLabel.Text = '0'
$countValueLabel.Font = New-Object System.Drawing.Font('Consolas', 15, [System.Drawing.FontStyle]::Bold)
$countValueLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#1F5E86')
$countValueLabel.Location = New-Object System.Drawing.Point(126, 16)
$countValueLabel.Size = New-Object System.Drawing.Size(40, 28)
$toolbarCard.Controls.Add($countValueLabel)

$addMatrixButton = New-Object System.Windows.Forms.Button
$addMatrixButton.Text = $ui.AddMatrix
$addMatrixButton.Location = New-Object System.Drawing.Point(178, 14)
$addMatrixButton.Size = New-Object System.Drawing.Size(92, 34)
Set-ButtonStyle -Button $addMatrixButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
$toolbarCard.Controls.Add($addMatrixButton)

$addSpecialButton = New-Object System.Windows.Forms.Button
$addSpecialButton.Text = $ui.AddSpecial
$addSpecialButton.Location = New-Object System.Drawing.Point(282, 14)
$addSpecialButton.Size = New-Object System.Drawing.Size(118, 34)
Set-ButtonStyle -Button $addSpecialButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
$toolbarCard.Controls.Add($addSpecialButton)

$removeMatrixButton = New-Object System.Windows.Forms.Button
$removeMatrixButton.Text = $ui.RemoveMatrix
$removeMatrixButton.Location = New-Object System.Drawing.Point(412, 14)
$removeMatrixButton.Size = New-Object System.Drawing.Size(92, 34)
Set-ButtonStyle -Button $removeMatrixButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#DCE8F3')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#284761'))
$toolbarCard.Controls.Add($removeMatrixButton)

$dataButton = New-Object System.Windows.Forms.Button
$dataButton.Text = $ui.DataImport
$dataButton.Location = New-Object System.Drawing.Point(516, 14)
$dataButton.Size = New-Object System.Drawing.Size(132, 34)
Set-ButtonStyle -Button $dataButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#E6F1EF')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#1F5A54'))
$toolbarCard.Controls.Add($dataButton)

$historyButton = New-Object System.Windows.Forms.Button
$historyButton.Text = $ui.History
$historyButton.Location = New-Object System.Drawing.Point(660, 14)
$historyButton.Size = New-Object System.Drawing.Size(96, 34)
Set-ButtonStyle -Button $historyButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F5E7CD')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#6B4B11'))
$toolbarCard.Controls.Add($historyButton)

$petButton = New-Object System.Windows.Forms.Button
$petButton.Text = '桌宠'
$petButton.Location = New-Object System.Drawing.Point(768, 14)
$petButton.Size = New-Object System.Drawing.Size(82, 34)
Set-ButtonStyle -Button $petButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#FCEEE8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#8B4A34'))
$toolbarCard.Controls.Add($petButton)

$petSettingsButton = New-Object System.Windows.Forms.Button
$petSettingsButton.Text = '桌宠设置'
$petSettingsButton.Location = New-Object System.Drawing.Point(860, 14)
$petSettingsButton.Size = New-Object System.Drawing.Size(110, 34)
Set-ButtonStyle -Button $petSettingsButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
$toolbarCard.Controls.Add($petSettingsButton)

$syntaxLabel = New-Object System.Windows.Forms.Label
$syntaxLabel.Text = $ui.Syntax
$syntaxLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#62778E')
$syntaxLabel.Location = New-Object System.Drawing.Point(14, 58)
$syntaxLabel.Size = New-Object System.Drawing.Size(1400, 24)
$syntaxLabel.Anchor = 'Top,Left,Right'
$toolbarCard.Controls.Add($syntaxLabel)

$leftSplit = New-Object System.Windows.Forms.SplitContainer
$leftSplit.Dock = 'Fill'
$leftSplit.Orientation = 'Horizontal'
$leftSplit.SplitterWidth = 6
$leftSplit.SplitterDistance = 520
$leftSplit.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#D8E2ED')
$mainSplit.Panel1.Controls.Add($leftSplit)

$rightSplit = New-Object System.Windows.Forms.SplitContainer
$rightSplit.Dock = 'Fill'
$rightSplit.Orientation = 'Horizontal'
$rightSplit.SplitterWidth = 6
$rightSplit.SplitterDistance = 350
$rightSplit.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#D8E2ED')
$mainSplit.Panel2.Controls.Add($rightSplit)

$workspaceCard = New-Object System.Windows.Forms.Panel
$workspaceCard.Dock = 'Fill'
$workspaceCard.BackColor = [System.Drawing.Color]::White
$workspaceCard.BorderStyle = 'FixedSingle'
$leftSplit.Panel1.Controls.Add($workspaceCard)

$workspaceTitle = New-Object System.Windows.Forms.Label
$workspaceTitle.Text = $ui.Workspace
$workspaceTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
$workspaceTitle.Location = New-Object System.Drawing.Point(14, 16)
$workspaceTitle.Size = New-Object System.Drawing.Size(120, 24)
$workspaceCard.Controls.Add($workspaceTitle)

$workspaceHint = New-Object System.Windows.Forms.Label
$workspaceHint.Text = $ui.InputTip
$workspaceHint.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#687C92')
$workspaceHint.Location = New-Object System.Drawing.Point(14, 44)
$workspaceHint.Size = New-Object System.Drawing.Size(790, 22)
$workspaceHint.Anchor = 'Top,Left,Right'
$workspaceCard.Controls.Add($workspaceHint)

$workspaceBlockHint = New-Object System.Windows.Forms.Label
$workspaceBlockHint.Text = $ui.BlockHint
$workspaceBlockHint.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#687C92')
$workspaceBlockHint.Location = New-Object System.Drawing.Point(14, 68)
$workspaceBlockHint.Size = New-Object System.Drawing.Size(790, 22)
$workspaceBlockHint.Anchor = 'Top,Left,Right'
$workspaceCard.Controls.Add($workspaceBlockHint)

$workspaceBody = New-Object System.Windows.Forms.Panel
$workspaceBody.Location = New-Object System.Drawing.Point(12, 118)
$workspaceBody.Size = New-Object System.Drawing.Size(806, 396)
$workspaceBody.Anchor = 'Top,Bottom,Left,Right'
$workspaceCard.Controls.Add($workspaceBody)

$script:matrixFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$script:matrixFlow.Dock = 'Fill'
$script:matrixFlow.AutoScroll = $true
$script:matrixFlow.WrapContents = $true
$script:matrixFlow.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F8FBFD')
$script:matrixFlow.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 8)
$workspaceBody.Controls.Add($script:matrixFlow)

$resultCard = New-Object System.Windows.Forms.Panel
$resultCard.Dock = 'Fill'
$resultCard.BackColor = [System.Drawing.Color]::White
$resultCard.BorderStyle = 'FixedSingle'
$leftSplit.Panel2.Controls.Add($resultCard)

$resultTitle = New-Object System.Windows.Forms.Label
$resultTitle.Text = $ui.Result
$resultTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
$resultTitle.Location = New-Object System.Drawing.Point(14, 18)
$resultTitle.Size = New-Object System.Drawing.Size(80, 24)
$resultCard.Controls.Add($resultTitle)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = $ui.Ready
$statusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#6A7F95')
$statusLabel.Location = New-Object System.Drawing.Point(102, 20)
$statusLabel.Size = New-Object System.Drawing.Size(600, 22)
$resultCard.Controls.Add($statusLabel)

$resultTextBox = New-Object System.Windows.Forms.TextBox
$resultTextBox.Multiline = $true
$resultTextBox.ReadOnly = $true
$resultTextBox.ScrollBars = 'Both'
$resultTextBox.WordWrap = $false
$resultTextBox.BorderStyle = 'FixedSingle'
$resultTextBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FBFCFE')
$resultTextBox.Font = New-Object System.Drawing.Font('Consolas', 12)
$resultTextBox.Location = New-Object System.Drawing.Point(12, 58)
$resultTextBox.Size = New-Object System.Drawing.Size(718, 164)
$resultTextBox.Anchor = 'Top,Bottom,Left,Right'
$resultCard.Controls.Add($resultTextBox)

$expressionCard = New-Object System.Windows.Forms.Panel
$expressionCard.Dock = 'Fill'
$expressionCard.BackColor = [System.Drawing.Color]::White
$expressionCard.BorderStyle = 'FixedSingle'
$rightSplit.Panel1.Controls.Add($expressionCard)

$expressionLabel = New-Object System.Windows.Forms.Label
$expressionLabel.Text = $ui.Expression
$expressionLabel.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
$expressionLabel.Location = New-Object System.Drawing.Point(14, 26)
$expressionLabel.Size = New-Object System.Drawing.Size(120, 24)
$expressionCard.Controls.Add($expressionLabel)

$expressionTextBox = New-Object System.Windows.Forms.TextBox
$expressionTextBox.Font = New-Object System.Drawing.Font('Consolas', 15)
$expressionTextBox.Location = New-Object System.Drawing.Point(14, 68)
$expressionTextBox.Size = New-Object System.Drawing.Size(470, 34)
$expressionTextBox.Anchor = 'Top,Left,Right'
$expressionTextBox.Text = 'A*B'
$expressionCard.Controls.Add($expressionTextBox)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = $ui.Run
$runButton.Location = New-Object System.Drawing.Point(492, 66)
$runButton.Size = New-Object System.Drawing.Size(112, 36)
$runButton.Anchor = 'Top,Right'
Set-ButtonStyle -Button $runButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#1F7A8C')) -ForeColor ([System.Drawing.Color]::White)
$expressionCard.Controls.Add($runButton)

$sampleButton = New-Object System.Windows.Forms.Button
$sampleButton.Text = $ui.Sample
$sampleButton.Location = New-Object System.Drawing.Point(610, 66)
$sampleButton.Size = New-Object System.Drawing.Size(76, 36)
$sampleButton.Anchor = 'Top,Right'
Set-ButtonStyle -Button $sampleButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#F5E7CD')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#6B4B11'))
$expressionCard.Controls.Add($sampleButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = $ui.Clear
$clearButton.Location = New-Object System.Drawing.Point(692, 66)
$clearButton.Size = New-Object System.Drawing.Size(76, 36)
$clearButton.Anchor = 'Top,Right'
Set-ButtonStyle -Button $clearButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
$expressionCard.Controls.Add($clearButton)

$expressionHintLabel = New-Object System.Windows.Forms.Label
$expressionHintLabel.Text = $ui.ExprHint
$expressionHintLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#5F758C')
$expressionHintLabel.Location = New-Object System.Drawing.Point(14, 112)
$expressionHintLabel.Size = New-Object System.Drawing.Size(760, 22)
$expressionHintLabel.Anchor = 'Top,Left,Right'
$expressionCard.Controls.Add($expressionHintLabel)

$variableLabel = New-Object System.Windows.Forms.Label
$variableLabel.Text = 'Variables'
$variableLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#5F758C')
$variableLabel.Location = New-Object System.Drawing.Point(14, 146)
$variableLabel.Size = New-Object System.Drawing.Size(120, 22)
$expressionCard.Controls.Add($variableLabel)

$script:variableQuickPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:variableQuickPanel.Location = New-Object System.Drawing.Point(14, 174)
$script:variableQuickPanel.Size = New-Object System.Drawing.Size(760, 52)
$script:variableQuickPanel.Anchor = 'Top,Left,Right'
$script:variableQuickPanel.WrapContents = $true
$script:variableQuickPanel.AutoScroll = $true
$script:variableQuickPanel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
$expressionCard.Controls.Add($script:variableQuickPanel)

$operatorLabel = New-Object System.Windows.Forms.Label
$operatorLabel.Text = 'Functions & Operators'
$operatorLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#5F758C')
$operatorLabel.Location = New-Object System.Drawing.Point(14, 228)
$operatorLabel.Size = New-Object System.Drawing.Size(180, 22)
$expressionCard.Controls.Add($operatorLabel)

$script:functionQuickPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:functionQuickPanel.Location = New-Object System.Drawing.Point(14, 256)
$script:functionQuickPanel.Size = New-Object System.Drawing.Size(760, 92)
$script:functionQuickPanel.Anchor = 'Top,Bottom,Left,Right'
$script:functionQuickPanel.WrapContents = $true
$script:functionQuickPanel.AutoScroll = $true
$expressionCard.Controls.Add($script:functionQuickPanel)

$helpCard = New-Object System.Windows.Forms.Panel
$helpCard.Dock = 'Fill'
$helpCard.BackColor = [System.Drawing.Color]::White
$helpCard.BorderStyle = 'FixedSingle'
$rightSplit.Panel2.Controls.Add($helpCard)

$helpTitle = New-Object System.Windows.Forms.Label
$helpTitle.Text = $ui.Help
$helpTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
$helpTitle.Location = New-Object System.Drawing.Point(14, 12)
$helpTitle.Size = New-Object System.Drawing.Size(120, 24)
$helpCard.Controls.Add($helpTitle)

$helpSearchBox = New-Object System.Windows.Forms.TextBox
$helpSearchBox.Location = New-Object System.Drawing.Point(14, 42)
$helpSearchBox.Size = New-Object System.Drawing.Size(510, 30)
$helpSearchBox.Anchor = 'Top,Left,Right'
$helpSearchBox.Text = ''
$helpCard.Controls.Add($helpSearchBox)

$helpSearchButton = New-Object System.Windows.Forms.Button
$helpSearchButton.Text = $ui.HelpSearch
$helpSearchButton.Location = New-Object System.Drawing.Point(532, 40)
$helpSearchButton.Size = New-Object System.Drawing.Size(84, 32)
$helpSearchButton.Anchor = 'Top,Right'
Set-ButtonStyle -Button $helpSearchButton -BackColor ([System.Drawing.ColorTranslator]::FromHtml('#EDF3F8')) -ForeColor ([System.Drawing.ColorTranslator]::FromHtml('#35516B'))
$helpCard.Controls.Add($helpSearchButton)

$helpPromptLabel = New-Object System.Windows.Forms.Label
$helpPromptLabel.Text = $ui.HelpPrompt
$helpPromptLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#687C92')
$helpPromptLabel.Location = New-Object System.Drawing.Point(14, 78)
$helpPromptLabel.Size = New-Object System.Drawing.Size(600, 22)
$helpPromptLabel.Anchor = 'Top,Left,Right'
$helpCard.Controls.Add($helpPromptLabel)

$helpQuickViewport = New-Object System.Windows.Forms.Panel
$helpQuickViewport.Location = New-Object System.Drawing.Point(14, 112)
$helpQuickViewport.Size = New-Object System.Drawing.Size(760, 38)
$helpQuickViewport.Anchor = 'Top,Left,Right'
$helpQuickViewport.BorderStyle = 'FixedSingle'
$helpQuickViewport.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FBFCFE')
$helpQuickViewport.AutoScroll = $false
$helpQuickViewport.TabStop = $true
$helpCard.Controls.Add($helpQuickViewport)

$script:helpQuickPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:helpQuickPanel.Location = New-Object System.Drawing.Point(0, 0)
$script:helpQuickPanel.AutoSize = $true
$script:helpQuickPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$script:helpQuickPanel.WrapContents = $false
$script:helpQuickPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$script:helpQuickPanel.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$helpQuickViewport.Controls.Add($script:helpQuickPanel)

$helpQuickScrollBar = New-Object System.Windows.Forms.HScrollBar
$helpQuickScrollBar.Location = New-Object System.Drawing.Point(14, 152)
$helpQuickScrollBar.Size = New-Object System.Drawing.Size(760, 12)
$helpQuickScrollBar.Anchor = 'Top,Left,Right'
$helpQuickScrollBar.Visible = $false
$helpCard.Controls.Add($helpQuickScrollBar)

$helpTextBox = New-Object System.Windows.Forms.TextBox
$helpTextBox.Multiline = $true
$helpTextBox.ReadOnly = $true
$helpTextBox.ScrollBars = 'Vertical'
$helpTextBox.BorderStyle = 'FixedSingle'
$helpTextBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FBFCFE')
$helpTextBox.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$helpTextBox.Location = New-Object System.Drawing.Point(14, 172)
$helpTextBox.Size = New-Object System.Drawing.Size(760, 142)
$helpTextBox.Anchor = 'Top,Bottom,Left,Right'
$helpTextBox.Text = $ui.HelpEmpty
$helpCard.Controls.Add($helpTextBox)

$script:updateHelpQuickScroll = {
    if ($null -eq $script:helpQuickPanel) { return }
    $contentWidth = [Math]::Max($script:helpQuickPanel.PreferredSize.Width + 8, $helpQuickViewport.ClientSize.Width)
    $script:helpQuickPanel.Height = [Math]::Max($script:helpQuickPanel.PreferredSize.Height, 34)
    $viewportWidth = [Math]::Max(1, $helpQuickViewport.ClientSize.Width)
    $needsScroll = ($contentWidth -gt $viewportWidth)

    if ($needsScroll) {
        $maxOffset = [Math]::Max(0, $contentWidth - $viewportWidth)
        $helpQuickScrollBar.Visible = $true
        $helpQuickScrollBar.SmallChange = 48
        $helpQuickScrollBar.LargeChange = $viewportWidth
        $helpQuickScrollBar.Maximum = $maxOffset + $helpQuickScrollBar.LargeChange - 1
        if ($helpQuickScrollBar.Value -gt $maxOffset) { $helpQuickScrollBar.Value = $maxOffset }
    } else {
        $helpQuickScrollBar.Visible = $false
        $helpQuickScrollBar.Value = 0
    }

    $script:helpQuickPanel.Left = -1 * $helpQuickScrollBar.Value
    $helpTextBox.Top = $(if ($helpQuickScrollBar.Visible) { 174 } else { 168 })
    $helpTextBox.Height = [Math]::Max(110, $helpCard.ClientSize.Height - $helpTextBox.Top - 12)
}

$script:updateAdaptiveLayout = {
    $workspaceInset = 12
    $workspaceBody.Width = [Math]::Max(260, $workspaceCard.ClientSize.Width - ($workspaceInset * 2))
    $workspaceBody.Height = [Math]::Max(180, $workspaceCard.ClientSize.Height - 140)
    $workspaceBody.Left = $workspaceInset
    $workspaceBody.Top = 134

    $cardOuterWidth = 444
    $availableWidth = [Math]::Max(240, $script:matrixFlow.ClientSize.Width)
    $columns = [Math]::Max(1, [int][Math]::Floor(($availableWidth + 14) / $cardOuterWidth))
    $usedWidth = ($columns * $cardOuterWidth) - 14
    $leftPadding = [Math]::Max(12, [int][Math]::Floor(($availableWidth - $usedWidth) / 2))
    $script:matrixFlow.Padding = New-Object System.Windows.Forms.Padding($leftPadding, 12, $leftPadding, 8)

    $exprClientWidth = $expressionCard.ClientSize.Width
    $contentWidth = [Math]::Min(760, [Math]::Max(360, $exprClientWidth - 28))
    $exprLeft = [Math]::Max(14, [int][Math]::Floor(($exprClientWidth - $contentWidth) / 2))
    $buttonGap = 8
    $runWidth = 112
    $sampleWidth = 76
    $clearWidth = 76
    $buttonWidth = $runWidth + $sampleWidth + $clearWidth + ($buttonGap * 2)
    $exprInputWidth = [Math]::Max(180, $contentWidth - $buttonWidth - $buttonGap)
    $rowWidth = $exprInputWidth + $buttonGap + $buttonWidth
    $rowLeft = [Math]::Max($exprLeft, [int][Math]::Floor(($exprClientWidth - $rowWidth) / 2))

    $expressionLabel.Left = $exprLeft
    $expressionLabel.Top = 54
    $expressionTextBox.Left = $rowLeft
    $expressionTextBox.Top = 108
    $expressionTextBox.Width = $exprInputWidth
    $runButton.Left = $expressionTextBox.Right + $buttonGap
    $runButton.Top = 106
    $sampleButton.Left = $runButton.Right + $buttonGap
    $sampleButton.Top = 106
    $clearButton.Left = $sampleButton.Right + $buttonGap
    $clearButton.Top = 106
    $expressionHintLabel.Left = $exprLeft
    $expressionHintLabel.Top = 160
    $expressionHintLabel.Width = $contentWidth
    $variableLabel.Left = $exprLeft
    $variableLabel.Top = 198
    $script:variableQuickPanel.Left = $exprLeft
    $script:variableQuickPanel.Top = 228
    $script:variableQuickPanel.Width = $contentWidth
    $buttonOuterWidth = 126
    $buttonsPerRow = [Math]::Max(1, [int][Math]::Floor(($contentWidth + 8) / $buttonOuterWidth))
    $variableRows = [Math]::Max(1, [int][Math]::Ceiling([Math]::Max(1, $matrixStore.Count) / [double]$buttonsPerRow))
    $variablePanelHeight = [Math]::Min(96, [Math]::Max(44, ($variableRows * 36) + 8))
    $script:variableQuickPanel.Height = $variablePanelHeight
    $operatorLabel.Left = $exprLeft
    $operatorLabel.Top = $script:variableQuickPanel.Bottom + 16
    $script:functionQuickPanel.Left = $exprLeft
    $script:functionQuickPanel.Top = $operatorLabel.Bottom + 8
    $script:functionQuickPanel.Width = $contentWidth
    $functionButtonOuterWidth = 110
    $functionButtonsPerRow = [Math]::Max(1, [int][Math]::Floor(($contentWidth + 8) / $functionButtonOuterWidth))
    $functionButtonCount = [Math]::Max(1, $script:functionQuickPanel.Controls.Count)
    $functionRows = [Math]::Max(1, [int][Math]::Ceiling($functionButtonCount / [double]$functionButtonsPerRow))
    $visibleFunctionRows = [Math]::Min(4, $functionRows)
    $desiredFunctionHeight = [Math]::Max(76, ($visibleFunctionRows * 40) + 8)
    $availableFunctionHeight = [Math]::Max(76, $expressionCard.ClientSize.Height - $script:functionQuickPanel.Top - 16)
    $script:functionQuickPanel.Height = [Math]::Min($availableFunctionHeight, $desiredFunctionHeight)

    $helpClientWidth = $helpCard.ClientSize.Width
    $helpContentWidth = [Math]::Min(760, [Math]::Max(320, $helpClientWidth - 28))
    $helpLeft = [Math]::Max(14, [int][Math]::Floor(($helpClientWidth - $helpContentWidth) / 2))
    $helpTitle.Left = $helpLeft
    $helpSearchBox.Left = $helpLeft
    $helpSearchBox.Top = 42
    $helpSearchButton.Left = $helpLeft + $helpContentWidth - $helpSearchButton.Width
    $helpSearchButton.Top = 40
    $helpSearchBox.Width = [Math]::Max(180, $helpSearchButton.Left - $helpLeft - 8)
    $helpPromptLabel.Left = $helpLeft
    $helpPromptLabel.Top = 80
    $helpPromptLabel.Width = $helpContentWidth
    $helpQuickViewport.Left = $helpLeft
    $helpQuickViewport.Top = 114
    $helpQuickViewport.Width = $helpContentWidth
    $helpQuickScrollBar.Left = $helpLeft
    $helpQuickScrollBar.Top = 156
    $helpQuickScrollBar.Width = $helpContentWidth
    $helpTextBox.Left = $helpLeft
    $helpTextBox.Width = $helpContentWidth
    & $script:updateHelpQuickScroll

    $resultClientWidth = $resultCard.ClientSize.Width
    $resultContentWidth = [Math]::Min(760, [Math]::Max(320, $resultClientWidth - 28))
    $resultLeft = [Math]::Max(12, [int][Math]::Floor(($resultClientWidth - $resultContentWidth) / 2))
    $resultTitle.Left = $resultLeft
    $resultTitle.Top = 24
    $statusLabel.Left = $resultLeft + 78
    $statusLabel.Top = 26
    $statusLabel.Width = [Math]::Max(180, $resultContentWidth - 90)
    $resultTextBox.Left = $resultLeft
    $resultTextBox.Top = 72
    $resultTextBox.Width = $resultContentWidth
    $resultTextBox.Height = [Math]::Max(120, $resultCard.ClientSize.Height - 84)
}

$helpQuickScrollBar.Add_ValueChanged({
    if ($null -ne $script:updateHelpQuickScroll) { & $script:updateHelpQuickScroll }
})

$helpQuickViewport.Add_MouseEnter({ $helpQuickViewport.Focus() | Out-Null })
$helpQuickViewport.Add_MouseWheel({
    param($sender, $args)
    if (-not $helpQuickScrollBar.Visible) { return }
    $delta = if ($args.Delta -lt 0) { $helpQuickScrollBar.SmallChange } else { -1 * $helpQuickScrollBar.SmallChange }
    $newValue = [Math]::Max($helpQuickScrollBar.Minimum, [Math]::Min($helpQuickScrollBar.Maximum - $helpQuickScrollBar.LargeChange + 1, $helpQuickScrollBar.Value + $delta))
    $helpQuickScrollBar.Value = $newValue
})

Add-ExpressionQuickButton -Text '+' -InsertText '+' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text '-' -InsertText '-' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text '*' -InsertText '*' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text '/' -InsertText '/' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text '.*' -InsertText '.*' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text './' -InsertText './' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text '^2' -InsertText '^2' -BackHex '#EAF2FA' -ForeHex '#23435C'
Add-ExpressionQuickButton -Text 'T()' -InsertText 'T(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'REF()' -InsertText 'REF(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'RREF()' -InsertText 'RREF(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'RANK()' -InsertText 'RANK(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'TR()' -InsertText 'TR(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'NORM()' -InsertText 'NORM(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'INV()' -InsertText 'INV(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'EIGVALS()' -InsertText 'EIGVALS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'EIGVECS()' -InsertText 'EIGVECS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'SVALS()' -InsertText 'SVALS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'CHOL()' -InsertText 'CHOL(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'LU()' -InsertText 'LU(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'QR()' -InsertText 'QR(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'ROWSUMS()' -InsertText 'ROWSUMS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'COLSUMS()' -InsertText 'COLSUMS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'ROWMEANS()' -InsertText 'ROWMEANS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'COLMEANS()' -InsertText 'COLMEANS(A)' -BackHex '#E0F1EE' -ForeHex '#1D5B55'
Add-ExpressionQuickButton -Text 'DET()' -InsertText 'DET(A)' -BackHex '#F8EFD9' -ForeHex '#74511A'
Add-ExpressionQuickButton -Text 'COF()' -InsertText 'COF(A)' -BackHex '#F8EFD9' -ForeHex '#74511A'
Add-ExpressionQuickButton -Text 'ADJ()' -InsertText 'ADJ(A)' -BackHex '#F8EFD9' -ForeHex '#74511A'
Add-ExpressionQuickButton -Text 'NULL()' -InsertText 'NULL(A)' -BackHex '#F8EFD9' -ForeHex '#74511A'

Add-HelpQuickButton -Text 'T()' -Query 'T'
Add-HelpQuickButton -Text 'RREF()' -Query 'RREF'
Add-HelpQuickButton -Text 'DET()' -Query 'DET'
Add-HelpQuickButton -Text 'INV()' -Query 'INV'
Add-HelpQuickButton -Text 'EIGVALS' -Query 'EIGVALS'
Add-HelpQuickButton -Text 'EIGVECS' -Query 'EIGVECS'
Add-HelpQuickButton -Text 'SVALS' -Query 'SVALS'
Add-HelpQuickButton -Text 'CHOL' -Query 'CHOL'
Add-HelpQuickButton -Text 'LU' -Query 'LU'
Add-HelpQuickButton -Text 'QR' -Query 'QR'
Add-HelpQuickButton -Text 'NULL()' -Query 'NULL'
Add-HelpQuickButton -Text 'ROWSUMS' -Query 'ROWSUMS'
Add-HelpQuickButton -Text 'COLSUMS' -Query 'COLSUMS'
Add-HelpQuickButton -Text 'ROWMEANS' -Query 'ROWMEANS'
Add-HelpQuickButton -Text 'COLMEANS' -Query 'COLMEANS'
Add-HelpQuickButton -Text 'ROWMAXS' -Query 'ROWMAXS'
Add-HelpQuickButton -Text 'COLMAXS' -Query 'COLMAXS'
Add-HelpQuickButton -Text 'ROWMINS' -Query 'ROWMINS'
Add-HelpQuickButton -Text 'COLMINS' -Query 'COLMINS'
Add-HelpQuickButton -Text '.*' -Query '.*'
Add-HelpQuickButton -Text './' -Query './'
Add-HelpQuickButton -Text 'Block' -Query 'BLOCK'
Add-HelpQuickButton -Text 'I3' -Query 'I3'
Add-HelpQuickButton -Text 'diag' -Query 'diag(1,2,3)'
Add-HelpQuickButton -Text 'Data' -Query 'DATA'
Add-HelpQuickButton -Text 'History' -Query 'HISTORY'
$addMatrixButton.Add_Click({ Add-MatrixCard -InitialText '[1 0; 0 1]' })
$addSpecialButton.Add_Click({ Add-MatrixCard -InitialText 'I3' -OpenSpecial })
$removeMatrixButton.Add_Click({ Remove-MatrixCard })
$dataButton.Add_Click({ Show-DataWindow })
$historyButton.Add_Click({ Show-HistoryWindow })
$petButton.Add_Click({
    try {
        $settings = Load-DesktopPetSettings
        $process = Start-DesktopPet -Settings $settings
        $statusLabel.Text = '桌宠已启动，进程 ID: ' + $process.Id
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '桌宠', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
})
$petSettingsButton.Add_Click({ Show-DesktopPetSettingsDialog })
$runButton.Add_Click({ Execute-Expression })
$expressionTextBox.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $_.SuppressKeyPress = $true; Execute-Expression } })

$sampleButton.Add_Click({
    while ($matrixStore.Count -gt 3) { Remove-MatrixCard }
    while ($matrixStore.Count -lt 3) { Add-MatrixCard -InitialText '[1 0; 0 1]' }

    $matrixStore[0].SuppressModeReset = $true
    $matrixStore[1].SuppressModeReset = $true
    $matrixStore[2].SuppressModeReset = $true
    try {
        $matrixStore[0].TextBox.Text = 'I3'
        $matrixStore[1].TextBox.Text = 'ones(3,3)'
        $matrixStore[2].TextBox.Text = '[A | B; B | A]'
    }
    finally {
        $matrixStore[0].SuppressModeReset = $false
        $matrixStore[1].SuppressModeReset = $false
        $matrixStore[2].SuppressModeReset = $false
    }

    Set-MatrixCardMode -Item $matrixStore[0] -Mode 'preset' -Descriptor 'I3'
    Set-MatrixCardMode -Item $matrixStore[1] -Mode 'preset' -Descriptor 'ones(3,3)'
    Set-MatrixCardMode -Item $matrixStore[2] -Mode 'manual' -Descriptor ''
    $expressionTextBox.Text = 'DET(C)'
    $statusLabel.Text = $ui.Ready
    Refresh-MatrixCards
})

$clearButton.Add_Click({
    foreach ($item in $matrixStore) {
        $item.SuppressModeReset = $true
        try {
            $item.TextBox.Clear()
        }
        finally {
            $item.SuppressModeReset = $false
        }
        Set-MatrixCardMode -Item $item -Mode 'manual' -Descriptor ''
    }
    $expressionTextBox.Clear()
    $resultTextBox.Clear()
    $statusLabel.Text = $ui.Ready
    Refresh-VariableQuickButtons
})

$helpSearchButton.Add_Click({ Show-HelpContent -Query $helpSearchBox.Text })
$helpSearchBox.Add_TextChanged({ Show-HelpContent -Query $helpSearchBox.Text })
$helpSearchBox.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $_.SuppressKeyPress = $true; Show-HelpContent -Query $helpSearchBox.Text } })

$form.Add_Resize({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$mainSplit.Add_SplitterMoved({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$leftSplit.Add_SplitterMoved({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$rightSplit.Add_SplitterMoved({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$workspaceCard.Add_Resize({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$expressionCard.Add_Resize({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$helpCard.Add_Resize({ if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout } })
$helpQuickViewport.Add_Resize({ if ($null -ne $script:updateHelpQuickScroll) { & $script:updateHelpQuickScroll } })

Load-OperationHistory
Load-DesktopPetSettings | Out-Null
Add-MatrixCard -InitialText '[1 2; 3 4]'
Add-MatrixCard -InitialText '[5 6; 7 8]'
Refresh-MatrixCards
Show-HelpContent -Query 'BLOCK'
if ($null -ne $script:updateAdaptiveLayout) { & $script:updateAdaptiveLayout }
$form.AcceptButton = $runButton
$form.Add_Shown({
    try {
        $settings = Load-DesktopPetSettings
        if ($settings.AutoStart) {
            $process = Start-DesktopPet -Settings $settings
            $statusLabel.Text = '桌宠已自动启动，进程 ID: ' + $process.Id
        }
    }
    catch {
        $statusLabel.Text = '桌宠自动启动失败：' + $_.Exception.Message
    }
})
$form.Add_FormClosing({ $script:isShuttingDown = $true })

[System.Windows.Forms.Application]::Run($form)
















