using MK.ExcelViewer.Sessions;
using Xunit;

namespace MK.ExcelViewer.Tests;

public class SessionTests
{
    [Fact]
    public void NewSession_IsWaiting_BeforeAnyBytesArrive()
    {
        var session = new SessionStore().Create("report.xlsx", 1000);
        var s = session.Snapshot();

        Assert.Equal(UploadStage.Waiting, s.Stage);
        Assert.Equal("report.xlsx", s.FileName);
        Assert.Equal(0, s.ReceivedBytes);
    }

    [Fact]
    public void Progress_ReportsAPercentage_WhenTheSenderDeclaredASize()
    {
        var session = new SessionStore().Create("r.xlsx", 1000);
        session.BeginReceiving("r.xlsx", 1000);
        session.Progress(430);

        var s = session.Snapshot();
        Assert.Equal(UploadStage.Receiving, s.Stage);
        Assert.Equal(43, s.Percent);
    }

    [Fact]
    public void Percent_IsNull_WhenNoSizeWasDeclared()
    {
        // No denominator means no honest percentage — the page must show motion, not a made-up number.
        var session = new SessionStore().Create("r.xlsx", null);
        session.BeginReceiving("r.xlsx", null);
        session.Progress(5000);

        Assert.Null(session.Snapshot().Percent);
    }

    [Fact]
    public void Percent_NeverExceeds100_EvenIfTheSenderUnderstatedTheSize()
    {
        var session = new SessionStore().Create("r.xlsx", 100);
        session.BeginReceiving("r.xlsx", 100);
        session.Progress(250);              // sender lied, or Content-Length was wrong

        Assert.Equal(100, session.Snapshot().Percent);
    }

    [Fact]
    public void StagesAdvance_WaitingToReceivingToOpeningToReady()
    {
        var session = new SessionStore().Create("r.xlsx", 10);

        Assert.Equal(UploadStage.Waiting, session.Snapshot().Stage);
        session.BeginReceiving("r.xlsx", 10);
        Assert.Equal(UploadStage.Receiving, session.Snapshot().Stage);
        session.Opening();
        Assert.Equal(UploadStage.Opening, session.Snapshot().Stage);
        session.Ready("abc123");
        Assert.Equal(UploadStage.Ready, session.Snapshot().Stage);
        Assert.Equal("abc123", session.Snapshot().Hash);
    }

    [Fact]
    public void Failure_CarriesTheReason_SoThePageCanShowIt()
    {
        var session = new SessionStore().Create("r.xlsx", 10);
        session.Fail("That's a Word document, not a workbook.");

        var s = session.Snapshot();
        Assert.Equal(UploadStage.Failed, s.Stage);
        Assert.Equal("That's a Word document, not a workbook.", s.Error);
    }

    [Fact]
    public void SessionIds_AreUnguessable()
    {
        // The id sits in a URL a human opens. Knowing one must not let you watch — or hijack —
        // someone else's publish.
        var store = new SessionStore();
        var ids = Enumerable.Range(0, 50).Select(_ => store.Create(null, null).Id).ToList();

        Assert.All(ids, id => Assert.Equal(32, id.Length));           // 128 bits, hex
        Assert.All(ids, id => Assert.True(id.All(Uri.IsHexDigit)));
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void UnknownSession_IsNull_NotAThrow() =>
        Assert.Null(new SessionStore().Get("nope"));
}
