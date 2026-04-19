$script:appRoot = if ($env:MATRIX_WORKSPACE_ROOT -and (Test-Path -LiteralPath $env:MATRIX_WORKSPACE_ROOT)) {
    $env:MATRIX_WORKSPACE_ROOT
} elseif ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

$script:moduleRoot = Join-Path $script:appRoot 'modules'
if (-not (Test-Path -LiteralPath $script:moduleRoot)) {
    $fallbackRoot = if ($PSScriptRoot) {
        $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        (Get-Location).Path
    }
    $script:moduleRoot = Join-Path $fallbackRoot 'modules'
}

. (Join-Path $script:moduleRoot 'matrix_core.ps1')
. (Join-Path $script:moduleRoot 'matrix_ui_helpers.ps1')
. (Join-Path $script:moduleRoot 'matrix_dialogs.ps1')
. (Join-Path $script:moduleRoot 'matrix_form.ps1')
