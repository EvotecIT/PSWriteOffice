using System;
using System.IO;

namespace PSWriteOffice.Services;

/// <summary>Commits task output through same-directory temporary files.</summary>
internal static class AtomicFileWriter
{
    internal static string WriteUnique(string directory, string fileName, byte[] bytes)
    {
        if (bytes == null)
        {
            throw new ArgumentNullException(nameof(bytes));
        }

        Directory.CreateDirectory(directory);
        var safeName = GetSafeFileName(fileName);
        var extension = Path.GetExtension(safeName);
        var stem = Path.GetFileNameWithoutExtension(safeName);

        for (var index = 1; ; index++)
        {
            var candidateName = index == 1 ? safeName : $"{stem}-{index}{extension}";
            var candidatePath = Path.Combine(directory, candidateName);
            var temporaryPath = CreateTemporaryPath(directory, candidateName);

            try
            {
                WriteTemporary(temporaryPath, bytes);
                try
                {
                    File.Move(temporaryPath, candidatePath);
                    return candidatePath;
                }
                catch (IOException) when (File.Exists(candidatePath))
                {
                    // Another writer won this name. Retry with the next deterministic suffix.
                }
            }
            finally
            {
                DeleteTemporary(temporaryPath);
            }
        }
    }

    internal static void Write(string path, byte[] bytes, bool overwrite)
    {
        if (bytes == null)
        {
            throw new ArgumentNullException(nameof(bytes));
        }

        Write(path, overwrite, temporaryPath => WriteTemporary(temporaryPath, bytes));
    }

    internal static void Write(string path, bool overwrite, Action<string> writeTemporaryFile)
    {
        if (writeTemporaryFile == null)
        {
            throw new ArgumentNullException(nameof(writeTemporaryFile));
        }

        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        var temporaryPath = CreateTemporaryPath(directory, Path.GetFileName(fullPath));

        try
        {
            writeTemporaryFile(temporaryPath);
            if (!File.Exists(temporaryPath))
            {
                throw new IOException($"The output writer did not create temporary file '{temporaryPath}'.");
            }

            if (!overwrite)
            {
                File.Move(temporaryPath, fullPath);
                return;
            }

            try
            {
                File.Move(temporaryPath, fullPath);
            }
            catch (IOException) when (File.Exists(fullPath))
            {
                File.Replace(temporaryPath, fullPath, destinationBackupFileName: null);
            }
        }
        finally
        {
            DeleteTemporary(temporaryPath);
        }
    }

    internal static string GetSafeFileName(string fileName, string fallbackName = "output.bin")
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return fallbackName;
        }

        var separatorIndex = Math.Max(fileName.LastIndexOf('/'), fileName.LastIndexOf('\\'));
        var name = separatorIndex >= 0 ? fileName.Substring(separatorIndex + 1) : fileName;
        const string portableInvalidCharacters = "<>:\"/\\|?*";
        var characters = name.ToCharArray();
        var platformInvalidCharacters = Path.GetInvalidFileNameChars();
        for (var index = 0; index < characters.Length; index++)
        {
            var character = characters[index];
            if (char.IsControl(character) ||
                portableInvalidCharacters.IndexOf(character) >= 0 ||
                Array.IndexOf(platformInvalidCharacters, character) >= 0)
            {
                characters[index] = '_';
            }
        }

        var safeName = new string(characters).TrimEnd(' ', '.');
        if (string.IsNullOrWhiteSpace(safeName) || safeName == "." || safeName == "..")
        {
            return fallbackName;
        }

        var firstDotIndex = safeName.IndexOf('.');
        var deviceToken = firstDotIndex >= 0 ? safeName.Substring(0, firstDotIndex) : safeName;
        if (IsReservedWindowsFileName(deviceToken))
        {
            safeName = "_" + safeName;
        }
        return safeName;
    }

    private static bool IsReservedWindowsFileName(string stem)
    {
        if (stem.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("NUL", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (stem.Length == 4 && stem[3] >= '1' && stem[3] <= '9')
        {
            return stem.StartsWith("COM", StringComparison.OrdinalIgnoreCase) ||
                stem.StartsWith("LPT", StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    private static string CreateTemporaryPath(string directory, string fileName)
    {
        var extension = Path.GetExtension(fileName);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        return Path.Combine(directory, $".{stem}.{Guid.NewGuid():N}.tmp{extension}");
    }

    private static void WriteTemporary(string path, byte[] bytes)
    {
        using var stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 81920,
            options: FileOptions.WriteThrough);
        stream.Write(bytes, 0, bytes.Length);
        stream.Flush(flushToDisk: true);
    }

    private static void DeleteTemporary(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
