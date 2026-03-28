using FluentAssertions;
using OfficeCli.Core;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for batch-resident IPC reliability.
/// Bug report: test-samples/bugs/batch-resident-failure/
/// XLSX batch+resident has ~21% failure rate ("Failed to send to resident"),
/// while non-resident batch has 0% failure.
/// </summary>
public class ResidentBatchReliabilityTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private readonly ITestOutputHelper _output;

    public ResidentBatchReliabilityTests(ITestOutputHelper output)
    {
        _output = output;
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
    }

    public void Dispose()
    {
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    [Fact]
    public async Task Xlsx_BatchViaResident_RapidFireOperations_ShouldNotDropMessages()
    {
        // Reproduces: XLSX batch + resident ~21% failure rate
        // Root cause: each TrySend creates a new pipe connection with 200ms timeout.
        // Under rapid sequential sends, the server may not accept fast enough.

        BlankDocCreator.Create(_xlsxPath);
        using var server = new ResidentServer(_xlsxPath, editable: true);
        var serverTask = Task.Run(() => server.RunAsync());
        await Task.Delay(300);

        const int totalOps = 30;
        int nullResponses = 0;
        int successResponses = 0;

        // Simulate what batch command does: sequential TrySend in a tight loop
        for (int i = 1; i <= totalOps; i++)
        {
            var req = new ResidentRequest
            {
                Command = "set",
                Args = { ["path"] = $"/Sheet1/A{i}" },
                Props = new[] { $"value=Item{i}" }
            };

            var response = ResidentClient.TrySend(_xlsxPath, req);
            if (response == null)
                nullResponses++;
            else
                successResponses++;
        }

        _output.WriteLine($"Results: {successResponses}/{totalOps} succeeded, {nullResponses} failed (null response)");

        // The bug: some operations get null response = "Failed to send to resident"
        // Expected: 0 failures. If this assertion fails, it confirms the IPC reliability bug.
        nullResponses.Should().Be(0,
            $"all {totalOps} operations should succeed via resident IPC, " +
            $"but {nullResponses} returned null ('Failed to send to resident')");
    }

    [Fact]
    public async Task Pptx_BatchViaResident_RapidFireOperations_ShouldNotDropMessages()
    {
        // PPTX has ~1.5% failure rate (much lower than XLSX but still nonzero)

        BlankDocCreator.Create(_pptxPath);
        using var server = new ResidentServer(_pptxPath, editable: true);
        var serverTask = Task.Run(() => server.RunAsync());
        await Task.Delay(300);

        // Add a slide first
        var addReq = new ResidentRequest
        {
            Command = "add",
            Args = { ["parent"] = "/" },
            Props = new[] { "type=slide" }
        };
        var addResp = ResidentClient.TrySend(_pptxPath, addReq);
        addResp.Should().NotBeNull("adding a slide should succeed");

        const int totalOps = 30;
        int nullResponses = 0;
        int successResponses = 0;

        for (int i = 1; i <= totalOps; i++)
        {
            var req = new ResidentRequest
            {
                Command = "add",
                Args = { ["parent"] = "/slide[1]", ["type"] = "shape" },
                Props = new[] { $"text=Shape{i}", "type=rect", "x=1cm", "y=1cm", "w=3cm", "h=2cm" }
            };

            var response = ResidentClient.TrySend(_pptxPath, req);
            if (response == null)
                nullResponses++;
            else
                successResponses++;
        }

        _output.WriteLine($"Results: {successResponses}/{totalOps} succeeded, {nullResponses} failed (null response)");

        nullResponses.Should().Be(0,
            $"all {totalOps} operations should succeed via resident IPC, " +
            $"but {nullResponses} returned null ('Failed to send to resident')");
    }

    [Fact]
    public async Task Xlsx_BatchViaResident_VerifyDataIntegrity_AfterRapidWrites()
    {
        // Even if TrySend returns a response, verify the data actually persisted.
        // This catches cases where the server accepted but silently dropped the write.

        BlankDocCreator.Create(_xlsxPath);
        using var server = new ResidentServer(_xlsxPath, editable: true);
        var serverTask = Task.Run(() => server.RunAsync());
        await Task.Delay(300);

        const int totalOps = 20;
        var writtenCells = new List<string>();

        for (int i = 1; i <= totalOps; i++)
        {
            var req = new ResidentRequest
            {
                Command = "set",
                Args = { ["path"] = $"/Sheet1/A{i}" },
                Props = new[] { $"value=Data{i}" }
            };

            var response = ResidentClient.TrySend(_xlsxPath, req);
            if (response != null)
                writtenCells.Add($"A{i}");
        }

        _output.WriteLine($"Written {writtenCells.Count}/{totalOps} cells");

        // Now read back each cell that got a successful response
        int verifiedCount = 0;
        int missingCount = 0;
        foreach (var cell in writtenCells)
        {
            var getReq = new ResidentRequest
            {
                Command = "get",
                Args = { ["path"] = $"/Sheet1/{cell}", ["depth"] = "1" }
            };
            var getResp = ResidentClient.TrySend(_xlsxPath, getReq);
            if (getResp != null && !string.IsNullOrEmpty(getResp.Stdout))
                verifiedCount++;
            else
                missingCount++;
        }

        _output.WriteLine($"Verified: {verifiedCount}/{writtenCells.Count}, Missing: {missingCount}");

        missingCount.Should().Be(0,
            "all cells that received a successful write response should be readable");
    }
}
