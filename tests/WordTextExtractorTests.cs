using AwesomeAssertions;
using Lestaly;

namespace Tivarloa.Tests;

[TestClass]
public class WordTextExtractorTests
{
    [TestMethod]
    public async Task Extract()
    {
        var url = new Uri("https://github.com/opavon/ThesisTemplate/raw/refs/heads/main/Thesis_and_Cover_Template.docx");

        var normalText = default(string);
        {
            var options = new WordTextExtractorOptions();
            var extractor = new WordTextExtractor();
            normalText = await extractor.ExtractAsync(url, options);
        }

        var footnotText = default(string);
        {
            var options = new WordTextExtractorOptions(WithFootnote: true);
            var extractor = new WordTextExtractor();
            footnotText = await extractor.ExtractAsync(url, options);
        }

        normalText.Should().NotBeEmpty();
        footnotText.Should().NotBeEmpty();
    }

    [TestMethod]
    public async Task ExtractOutline()
    {
        var url = new Uri("https://github.com/opavon/ThesisTemplate/raw/refs/heads/main/Thesis_and_Cover_Template.docx");

        var normalText = default(List<WordOutlineBlock>);
        {
            var options = new WordTextExtractorOptions();
            var extractor = new WordTextExtractor();
            normalText = await extractor.ExtractOutlineAsync(url, options);
        }

        var footnotText = default(List<WordOutlineBlock>);
        {
            var options = new WordTextExtractorOptions(WithFootnote: true);
            var extractor = new WordTextExtractor();
            footnotText = await extractor.ExtractOutlineAsync(url, options);
        }

        normalText.Should().NotBeEmpty();
        footnotText.Should().NotBeEmpty();
    }

}
