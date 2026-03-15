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

    $probe = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $probe.Start()
    try {
        $port = ([System.Net.IPEndPoint] $probe.LocalEndpoint).Port
    } finally {
        $probe.Stop()
    }

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = 'file.bin'
    }

    $url = "http://127.0.0.1:$port/$fileName"
    $job = Start-Job -ScriptBlock {
        param(
            [int] $JobPort,
            [string] $JobFilePath,
            [string] $JobContentType
        )

        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $JobPort)
        $listener.Start()

        try {
            try {
                $client = $listener.AcceptTcpClient()
            } catch {
                return
            }

            try {
                $bytes = [System.IO.File]::ReadAllBytes($JobFilePath)
                $stream = $client.GetStream()
                try {
                    $stream.ReadTimeout = 5000
                    $stream.WriteTimeout = 5000

                    $requestBuffer = New-Object byte[] 4096
                    $requestBytes = 0
                    do {
                        $read = $stream.Read($requestBuffer, $requestBytes, $requestBuffer.Length - $requestBytes)
                        if ($read -le 0) {
                            break
                        }

                        $requestBytes += $read
                        if ($requestBytes -ge 4) {
                            $requestText = [System.Text.Encoding]::ASCII.GetString($requestBuffer, 0, $requestBytes)
                            if ($requestText.Contains("`r`n`r`n")) {
                                break
                            }
                        }
                    } while ($requestBytes -lt $requestBuffer.Length)

                    $header = "HTTP/1.1 200 OK`r`nContent-Type: $JobContentType`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
                    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
                    $stream.Write($headerBytes, 0, $headerBytes.Length)
                    $stream.Write($bytes, 0, $bytes.Length)
                    $stream.Flush()
                } finally {
                    $stream.Dispose()
                }
            } finally {
                $client.Dispose()
            }
        } finally {
            try {
                $listener.Stop()
            } catch {
            }
        }
    } -ArgumentList $port, $FilePath, $ContentType

    Start-Sleep -Milliseconds 300

    [PSCustomObject]@{
        Url = $url
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
