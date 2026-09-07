using System.Text;
using PSWriteOffice.Services;

namespace PSWriteOffice.Tests;

public sealed class AtomicFileWriterTests
{
    [Fact]
    public void WriteUnique_writes_payload_once_when_names_collide()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            File.WriteAllText(Path.Combine(directory, "report.txt"), "existing-1");
            File.WriteAllText(Path.Combine(directory, "report-2.txt"), "existing-2");
            var writeCount = 0;

            var outputPath = AtomicFileWriter.WriteUnique(directory, "report.txt", temporaryPath =>
            {
                writeCount++;
                File.WriteAllText(temporaryPath, "new-payload", new UTF8Encoding(false));
            });

            Assert.Equal(1, writeCount);
            Assert.Equal("report-3.txt", Path.GetFileName(outputPath));
            Assert.Equal("new-payload", File.ReadAllText(outputPath));
            Assert.Equal("existing-1", File.ReadAllText(Path.Combine(directory, "report.txt")));
            Assert.Equal("existing-2", File.ReadAllText(Path.Combine(directory, "report-2.txt")));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_cleans_temporary_file_when_writer_fails()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                AtomicFileWriter.WriteUnique(directory, "report.txt", temporaryPath =>
                {
                    File.WriteAllText(temporaryPath, "partial");
                    throw new InvalidOperationException("writer failed");
                }));

            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
