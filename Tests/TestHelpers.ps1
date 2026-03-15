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

function Start-TestHttpFileServer {
    param(
        [Parameter(Mandatory)]
        [string] $FilePath,

        [string] $ContentType = 'application/octet-stream'
    )

    $port = Get-Random -Minimum 20000 -Maximum 40000
    $prefix = "http://127.0.0.1:$port/"
    $job = Start-Job -ScriptBlock {
        param(
            [string] $JobPrefix,
            [string] $JobFilePath,
            [string] $JobContentType
        )

        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add($JobPrefix)
        $listener.Start()

        try {
            try {
                $context = $listener.GetContext()
            } catch [System.Net.HttpListenerException] {
                return
            } catch [System.ObjectDisposedException] {
                return
            }

            try {
                $bytes = [System.IO.File]::ReadAllBytes($JobFilePath)
                $context.Response.StatusCode = 200
                $context.Response.ContentType = $JobContentType
                $context.Response.ContentLength64 = $bytes.Length
                $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            } finally {
                $context.Response.OutputStream.Close()
                $context.Response.Close()
            }
        } finally {
            try {
                $listener.Stop()
            } catch {
            }
            $listener.Close()
        }
    } -ArgumentList $prefix, $FilePath, $ContentType

    Start-Sleep -Milliseconds 300

    [PSCustomObject]@{
        Url = "${prefix}file"
        Job = $job
    }
}

function Stop-TestHttpFileServer {
    param(
        [Parameter(Mandatory)]
        $Server
    )

    if ($Server.Job) {
        try {
            Wait-Job -Job $Server.Job -Timeout 2 | Out-Null
        } catch {
        }

        try {
            Stop-Job -Job $Server.Job -ErrorAction SilentlyContinue | Out-Null
        } catch {
        }

        try {
            Remove-Job -Job $Server.Job -Force -ErrorAction SilentlyContinue | Out-Null
        } catch {
        }
    }
}
