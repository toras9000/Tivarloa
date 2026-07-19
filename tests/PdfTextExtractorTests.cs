using AwesomeAssertions;

namespace Tivarloa.Tests;

[TestClass]
public class PdfTextExtractorTests
{
    [TestMethod]
    public async Task Extract()
    {
        var url = new Uri("https://pdfobject.com/pdf/sample.pdf");

        var normalPages = default(string);
        {
            var options = new PdfTextExtractorOptions();
            var extractor = new PdfTextExtractor();
            normalPages = await extractor.ExtractAsync(url, options);
        }

        var actualTextPages = default(string);
        {
            var options = new PdfTextExtractorOptions(UseActualText: true);
            var extractor = new PdfTextExtractor();
            actualTextPages = await extractor.ExtractAsync(url, options);
        }

        var whiteCompactPages = default(string);
        {
            var options = new PdfTextExtractorOptions(WhitespaceCompaction: true);
            var extractor = new PdfTextExtractor();
            whiteCompactPages = await extractor.ExtractAsync(url, options);
        }

        var paraSepaPages = default(string);
        {
            var options = new PdfTextExtractorOptions(ParagraphSeparation: true);
            var extractor = new PdfTextExtractor();
            paraSepaPages = await extractor.ExtractAsync(url, options);
        }

        normalPages.Should().NotBeEmpty();
        actualTextPages.Should().NotBeEmpty();
        whiteCompactPages.Should().NotBeEmpty();
        paraSepaPages.Should().NotBeEmpty();
    }

    [TestMethod]
    public async Task ExtractPages()
    {
        var url = new Uri("https://pdfobject.com/pdf/sample.pdf");

        var normalPages = default(List<PdfPage>);
        {
            var options = new PdfTextExtractorOptions();
            var extractor = new PdfTextExtractor();
            normalPages = await extractor.ExtractPagesAsync(url, options);
        }

        var actualTextPages = default(List<PdfPage>);
        {
            var options = new PdfTextExtractorOptions(UseActualText: true);
            var extractor = new PdfTextExtractor();
            actualTextPages = await extractor.ExtractPagesAsync(url, options);
        }

        var whiteCompactPages = default(List<PdfPage>);
        {
            var options = new PdfTextExtractorOptions(WhitespaceCompaction: true);
            var extractor = new PdfTextExtractor();
            whiteCompactPages = await extractor.ExtractPagesAsync(url, options);
        }

        var paraSepaPages = default(List<PdfPage>);
        {
            var options = new PdfTextExtractorOptions(ParagraphSeparation: true);
            var extractor = new PdfTextExtractor();
            paraSepaPages = await extractor.ExtractPagesAsync(url, options);
        }

        normalPages.Should().HaveCountGreaterThan(0);
        actualTextPages.Should().HaveCountGreaterThan(0);
        whiteCompactPages.Should().HaveCountGreaterThan(0);
        paraSepaPages.Should().HaveCountGreaterThan(0);
    }
}
