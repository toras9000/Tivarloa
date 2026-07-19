using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Tivarloa;

/// <summary>PDFページテキスト</summary>
/// <param name="Number">ページ番号</param>
/// <param name="Text">ページテキスト</param>
public record PdfPage(int Number, string Text);

/// <summary>PDFテキスト抽出オプション</summary>
/// <param name="Password">パスワード</param>
/// <param name="UseActualText">パスワード</param>
/// <param name="WhitespaceCompaction">空白文字を単一のスペースに置き換える</param>
/// <param name="ParagraphSeparation">段落区切りに2つの改行を挿入する</param>
public record PdfTextExtractorOptions(string? Password = default, bool UseActualText = false, bool WhitespaceCompaction = false, bool ParagraphSeparation = false);

/// <summary>PDFテキスト抽出</summary>
public class PdfTextExtractor : IDisposable
{
    // 公開メソッド
    #region テキスト抽出：文書全体
    /// <summary>PDF文書全体のテキストを取得する</summary>
    /// <param name="url">PDFのWeb URL</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>全体のテキスト</returns>
    public async ValueTask<string> ExtractAsync(Uri url, PdfTextExtractorOptions? options = default, CancellationToken cancelToken = default)
    {
        using var downloadStream = await this.http.Value.GetStreamAsync(url, cancelToken);
        var pages = ExtractPages(downloadStream, options);
        return string.Join(Environment.NewLine, pages.Select(p => p.Text));
    }

    /// <summary>PDF文書全体のテキストを取得する</summary>
    /// <param name="file">PDFファイル</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>全体のテキスト</returns>
    public string Extract(FileInfo file, PdfTextExtractorOptions? options = default)
    {
        using var fileStream = file.OpenRead();
        var pages = ExtractPages(fileStream, options);
        return string.Join(Environment.NewLine, pages.Select(p => p.Text));
    }

    /// <summary>PDF文書全体のテキストを取得する</summary>
    /// <param name="stream">PDFストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>全体のテキスト</returns>
    public string Extract(Stream stream, PdfTextExtractorOptions? options = default)
    {
        var pages = ExtractPages(stream, options);
        return string.Join(Environment.NewLine, pages.Select(p => p.Text));
    }
    #endregion

    #region テキスト抽出：ページ毎
    /// <summary>PDFのページ毎テキストを抽出する</summary>
    /// <param name="url">PDFのWeb URL</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>ページ毎テキスト</returns>
    public async ValueTask<List<PdfPage>> ExtractPagesAsync(Uri url, PdfTextExtractorOptions? options = default, CancellationToken cancelToken = default)
    {
        using var downloadStream = await this.http.Value.GetStreamAsync(url, cancelToken);
        return ExtractPages(downloadStream, options);
    }

    /// <summary>PDFのページ毎テキストを抽出する</summary>
    /// <param name="file">PDFファイル</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>ページ毎テキスト</returns>
    public List<PdfPage> ExtractPages(FileInfo file, PdfTextExtractorOptions? options = default)
    {
        using var fileStream = file.OpenRead();
        return ExtractPages(fileStream, options);
    }

    /// <summary>PDFのページ毎テキストを抽出する</summary>
    /// <param name="stream">PDFストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>ページ毎テキスト</returns>
    public List<PdfPage> ExtractPages(Stream stream, PdfTextExtractorOptions? options = default)
    {
        // パースオプション
        var parseOptions = default(ParsingOptions);

        // テキスト抽出オプション
        var textOptions = new ContentOrderTextExtractor.Options();

        // 抽出オプションを(あれば)適用
        if (options != null)
        {
            // PDFパースオプション
            parseOptions ??= new();
            parseOptions.UseActualText = options.UseActualText;
            if (options.Password != null) parseOptions.Password = options.Password;
            // テキスト抽出
            textOptions.ReplaceWhitespaceWithSpace = options.WhitespaceCompaction;
            textOptions.SeparateParagraphsWithDoubleNewline = options.ParagraphSeparation;
        }

        // ページ毎にテキスト抽出
        var blocks = new List<PdfPage>();
        using var document = PdfDocument.Open(stream, parseOptions);
        foreach (var page in document.GetPages())
        {
            var text = ContentOrderTextExtractor.GetText(page, textOptions);
            blocks.Add(new(page.Number, text));
        }

        return blocks;
    }
    #endregion

    #region 破棄
    /// <summary>リソース破棄</summary>
    public void Dispose()
    {
        this.http.Value.Dispose();
    }
    #endregion

    // 非公開フィールド
    #region 内部リソース
    /// <summary>HTTPクライアント</summary>
    private Lazy<HttpClient> http = new(() => new());
    #endregion
}
