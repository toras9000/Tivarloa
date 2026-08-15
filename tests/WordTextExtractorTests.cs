using AwesomeAssertions;
using Lestaly;
using Tivarloa.Tests._Helper;

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

        var normalText = default(WordOutlineBlock[]);
        {
            var options = new WordTextExtractorOptions();
            var extractor = new WordTextExtractor();
            normalText = await extractor.ExtractOutlineAsync(url, options);
        }

        var footnotText = default(WordOutlineBlock[]);
        {
            var options = new WordTextExtractorOptions(WithFootnote: true);
            var extractor = new WordTextExtractor();
            footnotText = await extractor.ExtractOutlineAsync(url, options);
        }

        normalText.Should().NotBeEmpty();
        footnotText.Should().NotBeEmpty();
    }

    [TestMethod]
    public async Task ExtractOutline_Generate()
    {
        using var tempDir = new TempDir();
        var testFile = WordDocHelper.CreateTestDocument(tempDir.Info.RelativeFile("test.docx"));

        var options = new WordTextExtractorOptions();
        var extractor = new WordTextExtractor();
        var outline = extractor.ExtractOutline(testFile, options);

        outline[0].Number.Should().Be("%%[1]%%");
        outline[1].Number.Should().Be("<1 - 1>");
        outline[2].Number.Should().Be("<1 - 2>");
        outline[3].Number.Should().Be("%%[2]%%");
    }
}
