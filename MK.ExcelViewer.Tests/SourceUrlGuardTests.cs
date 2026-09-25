using MK.ExcelViewer.Security;
using Xunit;

namespace MK.ExcelViewer.Tests;

public class SourceUrlGuardTests
{
    private static readonly string[] Hosts = ["files.phoebus.asia"];

    [Theory]
    [InlineData("https://files.phoebus.asia/api/v4/file/content/p7Fp/0/report.xlsx?sign=abc")]
    [InlineData("https://FILES.phoebus.asia/x.xlsx")]
    [InlineData("https://files.phoebus.asia:443/x.xlsx")]
    public void Allows_Https_OnAWhitelistedHost(string src)
    {
        Assert.True(SourceUrlGuard.TryParse(src, Hosts, out _));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("not a url")]
    [InlineData("/relative/x.xlsx")]
    [InlineData("http://files.phoebus.asia/x.xlsx")]            // not https
    [InlineData("https://files.phoebus.asia:8443/x.xlsx")]      // not the default port
    [InlineData("https://user:pw@files.phoebus.asia/x.xlsx")]   // credentials
    [InlineData("https://evil-files.phoebus.asia/x.xlsx")]
    [InlineData("https://files.phoebus.asia.evil.com/x.xlsx")]
    [InlineData("https://phoebus.asia/x.xlsx")]
    [InlineData("https://127.0.0.1/x.xlsx")]
    [InlineData("file:///etc/passwd")]
    public void Refuses_EverythingElse(string? src)
    {
        Assert.False(SourceUrlGuard.TryParse(src, Hosts, out _));
    }

    [Fact]
    public void Refuses_EverythingWhenNoHostIsWhitelisted()
    {
        Assert.False(SourceUrlGuard.TryParse("https://files.phoebus.asia/x.xlsx", [], out _));
    }

    [Theory]
    [InlineData("Bảng so sánh.xlsx", "Bảng so sánh.xlsx")]
    [InlineData("../../etc/passwd", "passwd")]
    [InlineData("..\\..\\boot.ini", "boot.ini")]
    [InlineData(null, "from-url.xlsx")]
    [InlineData("", "from-url.xlsx")]
    public void FileName_PrefersTheHint_AndNeverReturnsAPath(string? hint, string expected)
    {
        var src = new Uri("https://files.phoebus.asia/api/v4/file/content/p7Fp/0/from-url.xlsx?sign=abc");
        Assert.Equal(expected, SourceUrlGuard.FileName(hint, src, "workbook.xlsx"));
    }

    [Fact]
    public void FileName_FallsBack_WhenTheUrlHasNoName()
    {
        Assert.Equal("workbook.xlsx", SourceUrlGuard.FileName(null, new Uri("https://files.phoebus.asia/"), "workbook.xlsx"));
    }
}
