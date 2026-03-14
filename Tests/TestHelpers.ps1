function New-TestOfficeImageFile {
    param(
        [Parameter(Mandatory)]
        [string] $Directory,

        [string] $Name = 'OfficeIMO.bmp'
    )

    if (-not (Test-Path -Path $Directory)) {
        $null = New-Item -Path $Directory -ItemType Directory -Force
    }

    $path = Join-Path $Directory $Name
    [byte[]] $bytes = 0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B,
        0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
        0xFF, 0x00
    [System.IO.File]::WriteAllBytes($path, $bytes)

    $path
}

function Get-ZipXmlDocumentLocal {
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [string] $Entry
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $zipEntry = $archive.GetEntry($Entry)
        if (-not $zipEntry) {
            throw "Zip entry '$Entry' not found in '$Path'."
        }

        $stream = $zipEntry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream)
            try {
                return [xml] $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $stream.Dispose()
        }
    } finally {
        $archive.Dispose()
    }
}

function Get-ZipEntriesLocal {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $archive.Entries) {
            $entry.FullName
        }
    } finally {
        $archive.Dispose()
    }
}
