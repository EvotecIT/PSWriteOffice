using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Pdf;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class DocumentPdfCmdletTests {
    [Fact]
    public async Task StoppingDuringPdfGenerationDoesNotPublishTheDestination() {
        string root = Path.Combine(Path.GetTempPath(), "pswriteoffice-pdf-stop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        using var shaper = new PausedShaper();
        try {
            string font = new[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf"),
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
            }.First(File.Exists);
            var options = new PdfOptions { DefaultFont = PdfStandardFont.Courier, TextShapingProvider = shaper }
                .EmbedStandardFont(PdfStandardFont.Courier, font);
            var conversion = PdfPlainTextConverter.ToPdfDocumentResult("Stop during actual PDF generation", new PdfPlainTextOptions { PdfOptions = options });
            string output = Path.Combine(root, "stopped.pdf");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("Document", conversion).AddParameter("Path", output);
            IAsyncResult invocation = shell.BeginInvoke();
            Assert.True(await Task.Run(() => shaper.Entered.Wait(TimeSpan.FromSeconds(15))), "The real font shaper was not reached during save.");
            IAsyncResult stopping = shell.BeginStop(null, null);
            Assert.True(SpinWait.SpinUntil(() => shell.InvocationStateInfo.State == PSInvocationState.Stopping, TimeSpan.FromSeconds(5)));
            shaper.Release.Set();
            await Task.Run(() => shell.EndStop(stopping));
            Assert.Throws<PipelineStoppedException>(() => shell.EndInvoke(invocation));
            Assert.False(File.Exists(output));
            Assert.Empty(Directory.GetFiles(root, "*.tmp"));
        } finally { shaper.Release.Set(); Directory.Delete(root, true); }
    }

    private sealed class PausedShaper : OfficeIMO.Drawing.IOfficeTextShapingProvider, IDisposable {
        public ManualResetEventSlim Entered { get; } = new();
        public ManualResetEventSlim Release { get; } = new();
        public OfficeIMO.Drawing.OfficeTextShapingResult? ShapeText(OfficeIMO.Drawing.OfficeTextShapingRequest request) {
            Entered.Set();
            if (!Release.Wait(TimeSpan.FromSeconds(20))) throw new TimeoutException("Font shaping cancellation test timed out.");
            return null;
        }
        public void Dispose() { Entered.Dispose(); Release.Dispose(); }
    }

    [Theory]
    [InlineData(".txt")]
    [InlineData(".doc")]
    public void FileExportProducesReopenablePdfAndWhatIfPreservesTheDestination(string extension) {
        string root = Path.Combine(Path.GetTempPath(), "pswriteoffice-pdf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "output.pdf");
            if (extension == ".txt") File.WriteAllText(input, "  # Literal text\n<b>kept as text</b>");
            else { using var word = WordDocument.Create(); word.AddParagraph("Legacy DOC evidence"); word.Save(input); }
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPath", input).AddParameter("Path", output).AddParameter("WhatIf");
            shell.Invoke(); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error)); Assert.False(File.Exists(output));
            shell.Commands.Clear();
            byte[] original = File.ReadAllBytes(input);
            using (var rejecting = PowerShell.Create()) {
                rejecting.Runspace = runspace;
                rejecting.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPath", input).AddParameter("Path", input);
                Assert.Throws<CmdletInvocationException>(() => rejecting.Invoke());
                Assert.Equal(original, File.ReadAllBytes(input));
            }
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPath", input).AddParameter("Path", output).AddParameter("PassThru");
            Assert.Single(shell.Invoke()); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            Assert.Single(PdfDocument.Load(output).Read().Pages);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ArchiveCmdletResumesAndWhatIfCreatesNoState() {
        string root = Path.Combine(Path.GetTempPath(), "pswriteoffice-archive-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "source"), output = Path.Combine(root, "output"), checkpoint = Path.Combine(root, "checkpoint");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "one.txt"), "Archive cmdlet evidence");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficePdfArchive", typeof(ExportOfficePdfArchiveCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            void AddCommand() => shell.AddCommand("Export-OfficePdfArchive").AddParameter("InputDirectory", input)
                .AddParameter("OutputDirectory", output).AddParameter("CheckpointDirectory", checkpoint);
            AddCommand(); shell.AddParameter("WhatIf"); shell.Invoke();
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(checkpoint));
            shell.Commands.Clear(); AddCommand();
            Assert.Single(shell.Invoke()); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            shell.Commands.Clear(); AddCommand();
            var resumed = Assert.IsType<OfficeIMO.Workflows.OfficePdfArchiveResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, resumed.Reused);
        } finally { Directory.Delete(root, true); }
    }
}
