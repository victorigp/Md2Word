param(
    [Parameter(Mandatory=$true)]
    [string]$MdFile
)

$content = Get-Content -Path $MdFile -Encoding UTF8
foreach ($line in $content) {
    if ($line -match '^\s*#\s+(.+)$') {
        $title = $Matches[1].Trim()
        # Remove inline formatting
        $title = $title -replace '\*\*(.+?)\*\*', '$1'
        $title = $title -replace '\*(.+?)\*', '$1'
        $title = $title -replace '~~(.+?)~~', '$1'
        $title = $title -replace '`(.+?)`', '$1'
        # Remove invalid filename characters
        $title = $title -replace '[\\/:*?"<>|]', '_'
        Write-Output $title
        exit 0
    }
}

Write-Output "Documento"
exit 0
