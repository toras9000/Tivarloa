
using System.Text;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Io;

namespace Tivarloa;

/// <summary>HTMLページテキスト</summary>
/// <param name="Title">ページタイトル</param>
/// <param name="Content">ページテキスト</param>
public record HtmlDocument(string Title, string Content);

/// <summary>HTMLテキスト抽出オプション</summary>
/// <param name="Timeout">Webリソースの読み取りタイムアウト時間</param>
/// <param name="UseJs">JavaScriptを有効にするか否か</param>
/// <param name="Trimming">テキストの空白トリムを行うか</param>
public record HtmlTextExtractorOptions(TimeSpan? Timeout = default, bool UseJs = false, bool Trimming = false);

/// <summary>HTMLテキスト抽出</summary>
public class HtmlTextExtractor
{
    // 公開メソッド
    #region テキスト抽出
    /// <summary>HTMLからテキストを抽出する</summary>
    /// <param name="url">HTML Web URL</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>抽出テキスト</returns>
    public async Task<HtmlDocument> ExtractAsync(Uri url, HtmlTextExtractorOptions? options = null, CancellationToken cancelToken = default)
    {
        // リクエスタ設定
        var requester = new DefaultHttpRequester();
        if (options?.Timeout != null) requester.Timeout = options.Timeout.Value;

        // コンテキストの構成
        var config = Configuration.Default.With(requester).WithDefaultLoader();
        if (options?.UseJs == true) config = config.WithJs();

        // コンテキスト生成
        using var context = BrowsingContext.New(config);

        // ドキュメントを取得
        using var document = await context.OpenAsync(url.AbsoluteUri, cancelToken);

        // テキストを抽出
        var page = extractDocumentText(document, options);

        return page;
    }

    /// <summary>HTMLからテキストを抽出する</summary>
    /// <param name="file">HTMLファイル</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>抽出テキスト</returns>
    public async Task<HtmlDocument> ExtractAsync(FileInfo file, HtmlTextExtractorOptions? options = null, CancellationToken cancelToken = default)
    {
        await using var fileStream = file.OpenRead();
        return await ExtractAsync(fileStream, options);
    }

    /// <summary>HTMLからテキストを抽出する</summary>
    /// <param name="stream">HTMLストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>抽出テキスト</returns>
    public async Task<HtmlDocument> ExtractAsync(Stream stream, HtmlTextExtractorOptions? options = null, CancellationToken cancelToken = default)
    {
        // コンテキストの構成
        var config = Configuration.Default;
        if (options?.UseJs == true) config = config.WithJs();

        // コンテキスト生成
        using var context = BrowsingContext.New(config);

        // パーササービスを取得
        var parser = context.GetService<IHtmlParser>() ?? throw new InvalidOperationException();

        // ドキュメントをパース
        using var document = await parser.ParseDocumentAsync(stream);

        // テキストを抽出
        var page = extractDocumentText(document, options);

        return page;
    }
    #endregion

    // 非公開メソッド
    #region テキスト抽出
    /// <summary>HTMLからテキストを抽出する</summary>
    /// <param name="document">HTMLドキュメント</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>抽出テキスト</returns>
    private HtmlDocument extractDocumentText(IDocument document, HtmlTextExtractorOptions? options)
    {
        // テキスト化から除外する要素を削除
        foreach (var element in document.QuerySelectorAll("script, style"))
        {
            element.Remove();
        }

        // テキスト取得
        var title = document.Title ?? "";
        var body = document.Body?.TextContent ?? "";

        // トリミングオプションが有効であれば空白をトリムする
        if (0 < body.Length && options?.Trimming == true)
        {
            // 行内の前後空白をトリムし、空行は1行にまとめる
            var trimming = new StringBuilder(capacity: body.Length);
            var whiteLine = false;
            foreach (var line in body.EnumerateLines())
            {
                if (line.IsWhiteSpace())
                {
                    if (!whiteLine) trimming.AppendLine();
                    whiteLine = true;
                }
                else
                {
                    whiteLine = false;
                    trimming.Append(line.Trim());
                    trimming.AppendLine();
                }
            }
            // トリム文字列で置き換え
            body = trimming.ToString();
        }

        return new(title, body);
    }
    #endregion
}
