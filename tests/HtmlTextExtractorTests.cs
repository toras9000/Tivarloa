using AwesomeAssertions;

namespace Tivarloa.Tests;

[TestClass]
public class HtmlTextExtractorTests
{
    [TestMethod]
    public async Task ExtractAsync()
    {
        var url = new Uri("https://devblogs.microsoft.com/dotnet/");

        var normalText = default(HtmlDocument);
        {
            var options = new HtmlTextExtractorOptions();
            var extractor = new HtmlTextExtractor();
            normalText = await extractor.ExtractAsync(url, options);
        }

        var trimmedText = default(HtmlDocument);
        {
            var options = new HtmlTextExtractorOptions(Trimming: true);
            var extractor = new HtmlTextExtractor();
            trimmedText = await extractor.ExtractAsync(url, options);
        }

        var jsText = default(HtmlDocument);
        {
            var options = new HtmlTextExtractorOptions(UseJs: true);
            var extractor = new HtmlTextExtractor();
            jsText = await extractor.ExtractAsync(url, options);
        }

        normalText.Content.Should().NotBeEmpty();
        trimmedText.Content.Should().NotBeEmpty();
        jsText.Content.Should().NotBeEmpty();
    }
}
