using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Tivarloa;

/// <summary>アウトラインブロック</summary>
/// <param name="Level">アウトラインレベル。</param>
/// <param name="Caption">アウトラインキャプション</param>
/// <param name="Content">アウトラインブロック内容</param>
public record WordOutlineBlock(int Level, string Caption, string Content);

/// <summary>抽出オプション</summary>
/// <param name="WithFootnote">脚注を出力するか否か</param>
public record WordTextExtractorOptions(bool WithFootnote = false);

/// <summary>Wordテキスト抽出</summary>
public class WordTextExtractor
{
    // 公開メソッド
    #region テキスト抽出：文書全体
    /// <summary>Word文書全体のテキストを取得する</summary>
    /// <param name="url">Word文書のWeb URL</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>抽出テキスト</returns>
    public async ValueTask<string> ExtractAsync(Uri url, WordTextExtractorOptions? options = null, CancellationToken cancelToken = default)
    {
        using var downloadStream = await this.http.Value.GetStreamAsync(url, cancelToken);
        return Extract(downloadStream, options);
    }

    /// <summary>Word文書全体のテキストを取得する</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>抽出テキスト</returns>
    public string Extract(FileInfo file, WordTextExtractorOptions? options = null)
    {
        using var fileStream = file.OpenRead();
        return Extract(fileStream, options);
    }

    /// <summary>Word文書全体のテキストを取得する</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>抽出テキスト</returns>
    public string Extract(Stream stream, WordTextExtractorOptions? options = null)
    {
        var builder = new StringBuilder();
        ExtractWrite(stream, builder, options);
        return builder.ToString();
    }
    #endregion

    #region テキスト抽出：文書全体 (書き出し)
    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="url">Word文書のWeb URL</param>
    /// <param name="builder">書き込み先文字列ビルダ</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    public ValueTask ExtractWriteAsync(Uri url, StringBuilder builder, WordTextExtractorOptions? options = null, CancellationToken cancelToken = default)
        => extractWriteAsync(url, new StringBuilderFacade(builder), options, cancelToken);

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="url">Word文書のWeb URL</param>
    /// <param name="writer">書き込み先テキストライタ</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    public ValueTask ExtractWriteAsync(Uri url, TextWriter writer, WordTextExtractorOptions? options = null, CancellationToken cancelToken = default)
        => extractWriteAsync(url, new TextWriterFacade(writer), options, cancelToken);

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="builder">書き込み先文字列ビルダ</param>
    /// <param name="options">抽出オプション</param>
    public void ExtractWrite(FileInfo file, StringBuilder builder, WordTextExtractorOptions? options = null)
        => extractWrite(file, new StringBuilderFacade(builder), options);

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="writer">書き込み先テキストライタ</param>
    /// <param name="options">抽出オプション</param>
    public void ExtractWrite(FileInfo file, TextWriter writer, WordTextExtractorOptions? options = null)
        => extractWrite(file, new TextWriterFacade(writer), options);

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="builder">書き込み先文字列ビルダ</param>
    /// <param name="options">抽出オプション</param>
    public void ExtractWrite(Stream stream, StringBuilder builder, WordTextExtractorOptions? options = null)
        => extractWrite(stream, new StringBuilderFacade(builder), options);

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="writer">書き込み先テキストライタ</param>
    /// <param name="options">抽出オプション</param>
    public void ExtractWrite(Stream stream, TextWriter writer, WordTextExtractorOptions? options = null)
        => extractWrite(stream, new TextWriterFacade(writer), options);
    #endregion

    #region テキスト抽出：アウトライン毎
    /// <summary>Word文書全体をアウトラインブロック毎にテキスト化する</summary>
    /// <param name="url">Word文書のWeb URL</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    /// <returns>アウトライン毎のテキスト情報</returns>
    public async ValueTask<List<WordOutlineBlock>> ExtractOutlineAsync(Uri url, WordTextExtractorOptions? options = null, CancellationToken cancelToken = default)
    {
        using var downloadStream = await this.http.Value.GetStreamAsync(url, cancelToken);
        return extractOutline(downloadStream, options);
    }

    /// <summary>Word文書全体をアウトラインブロック毎にテキスト化する</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>アウトライン毎のテキスト情報</returns>
    public List<WordOutlineBlock> ExtractOutline(FileInfo file, WordTextExtractorOptions? options = null)
    {
        using var fileStream = file.OpenRead();
        return extractOutline(fileStream, options);
    }

    /// <summary>Word文書全体をアウトラインブロック毎にテキスト化する</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>アウトライン毎のテキスト情報</returns>
    public List<WordOutlineBlock> ExtractOutline(Stream stream, WordTextExtractorOptions? options = null)
        => extractOutline(stream, options);
    #endregion

    #region 破棄
    /// <summary>リソース破棄</summary>
    public void Dispose()
    {
        this.http.Value.Dispose();
    }
    #endregion

    // 非公開型
    #region 状態管理
    /// <summary>抽出コンテキスト情報</summary>
    /// <param name="document">抽出対象ドキュメント</param>
    private class ExtractContext(WordprocessingDocument document)
    {
        /// <summary>使用している脚注ID</summary>
        public List<long> UsedFootnotes { get; } = new();

        /// <summary>ドキュメントのスタイル定義辞書</summary>
        public IReadOnlyDictionary<string, Style> Styles => this.styleCache.Value;

        /// <summary>ドキュメントの脚注定義辞書</summary>
        public IReadOnlyDictionary<long, Footnote> Footnotes => this.footnotesCache.Value;

        /// <summary>スタイル定義辞書キャッシュ</summary>
        private readonly Lazy<Dictionary<string, Style>> styleCache = new(() =>
        {
            return document?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .Where(s => s.StyleId?.HasValue != null)
                .ToDictionary(s => s.StyleId!.Value!) ?? [];
        });

        /// <summary>脚注定義辞書キャッシュ</summary>
        private readonly Lazy<Dictionary<long, Footnote>> footnotesCache = new(() =>
        {
            return document.MainDocumentPart?.FootnotesPart?.Footnotes?.Elements<Footnote>()
                .Where(f => f.Id?.HasValue == true)
                .ToDictionary(f => f.Id!.Value!) ?? [];
        });
    }
    #endregion

    // 非公開フィールド
    #region 内部リソース
    /// <summary>HTTPクライアント</summary>
    private Lazy<HttpClient> http = new(() => new());
    #endregion

    #region テキスト抽出：文書全体 (書き出し)
    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="url">Word文書のWeb URL</param>
    /// <param name="writer">書き込み処理</param>
    /// <param name="options">抽出オプション</param>
    /// <param name="cancelToken">キャンセルトークン</param>
    private async ValueTask extractWriteAsync(Uri url, IWriterFacade writer, WordTextExtractorOptions? options, CancellationToken cancelToken)
    {
        using var downloadStream = await this.http.Value.GetStreamAsync(url, cancelToken);
        extractWrite(downloadStream, writer, options);
    }

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="writer">書き込み処理</param>
    /// <param name="options">抽出オプション</param>
    private void extractWrite(FileInfo file, IWriterFacade writer, WordTextExtractorOptions? options)
    {
        using var fileStream = file.OpenRead();
        extractWrite(fileStream, writer, options);
    }

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="writer">書き込み処理</param>
    /// <param name="options">抽出オプション</param>
    private void extractWrite(Stream stream, IWriterFacade writer, WordTextExtractorOptions? options)
    {
        // 文書オープン
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        if (doc.MainDocumentPart == null) return;

        // 本文参照
        var docBody = doc.MainDocumentPart.Document?.Body;
        if (docBody == null) return;

        // 抽出コンテキストの生成
        var context = new ExtractContext(doc);

        // 文書のテキスト化
        foreach (var element in enumerateTargetElements(docBody))
        {
            switch (element)
            {
            case Paragraph paragraph:   // 段落
                extractParagraphText(context, paragraph, writer, options);
                break;

            case Table table:           // テーブル
                extractTableText(context, table, writer, options);
                break;

            default:
                break;
            }
        }

        // 脚注出力が有効であればテキスト化
        if (options?.WithFootnote == true)
        {
            writer.WriteLine("----------------------------------------");
            extractFootnoteText(context, writer);
        }
    }
    #endregion

    #region テキスト抽出：アウトライン毎
    /// <summary>Word文書全体をアウトラインブロック毎にテキスト化する</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>アウトライン毎のテキスト情報</returns>
    private List<WordOutlineBlock> extractOutline(Stream stream, WordTextExtractorOptions? options = null)
    {
        // 文書オープン
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        if (doc.MainDocumentPart == null) return [];

        // 本文参照
        var docBody = doc.MainDocumentPart.Document?.Body;
        if (docBody == null) return [];

        // 抽出コンテキストの生成
        var context = new ExtractContext(doc);

        // アウトライン毎のテキスト
        var outline = new List<WordOutlineBlock>();

        // 現在の収集情報
        var level = -1;
        var caption = new StringBuilderFacade(new());
        var content = new StringBuilderFacade(new());

        // 現在の収集情報を確定するローカルメソッド
        void commitOutlineBlock()
        {
            if (0 < content.Builder.Length || 0 < caption.Builder.Length)
            {
                outline.Add(new(level, caption.Builder.ToString(), content.Builder.ToString()));
                level = -1;
                caption.Builder.Clear();
                content.Builder.Clear();
            }
        }

        // 文書の各アウトラインブロックをテキスト化
        foreach (var element in enumerateTargetElements(docBody))
        {
            switch (element)
            {
            case Paragraph paragraph:
                if (extractOutlineLevel(context, paragraph) is int lv)
                {
                    commitOutlineBlock();
                    level = lv;
                    extractParagraphText(context, paragraph, caption, options);
                }
                else
                {
                    extractParagraphText(context, paragraph, content, options);
                }
                break;

            case Table table:
                extractTableText(context, table, content, options);
                break;

            default:
                break;
            }
        }

        // 最後のブロックを確定
        commitOutlineBlock();

        // 脚注出力が有効であればテキスト化
        if (options?.WithFootnote == true && 0 < context.UsedFootnotes.Count)
        {
            level = -1;
            caption.Write("Footnotes");
            extractFootnoteText(context, content);
            commitOutlineBlock();
        }

        return outline;
    }
    #endregion

    #region ドキュメント処理
    /// <summary>指定要素の配下からテキスト化対象要素を列挙する。</summary>
    /// <remarks>
    /// このメソッドでは Paragraph と Table をテキスト化対象とする。
    /// 入れ子の要素を多重処理しないよう、見つけた Paragraph と Table の子孫要素は列挙されない。
    /// </remarks>
    /// <param name="origin">起点となる要素</param>
    /// <returns>テキスト化対象要素を列挙するシーケンス</returns>
    private IEnumerable<OpenXmlElement> enumerateTargetElements(OpenXmlElement origin)
    {
        foreach (var element in origin.Elements<OpenXmlElement>())
        {
            // テキスト化対象であれば列挙
            if (element is Paragraph or Table)
            {
                yield return element;
                continue;
            }

            // 目的の要素でなければその配下を再帰的に列挙
            if (element.HasChildren)
            {
                foreach (var sub in enumerateTargetElements(element))
                {
                    yield return sub;
                }
            }
        }
    }

    /// <summary>段落のアウトラインレベルの取得を試みる</summary>
    /// <param name="context">抽出コンテキスト</param>
    /// <param name="paragraph">段落</param>
    /// <returns>段落がアウトラインレベルを持つ場合はそのレベル値。そうでない場合は null を返却</returns>
    private int? extractOutlineLevel(ExtractContext context, Paragraph paragraph)
    {
        // 段落プロパティの参照
        var paraProp = paragraph.ParagraphProperties;
        if (paraProp == null) return null;

        // 直接アウトライン設定されていれば利用
        var lv = paraProp.OutlineLevel?.Val;
        if (lv?.HasValue == true) return lv.Value;

        // 段落スタイルIDを取得
        var styleId = paraProp.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        // 段落スタイルを参照
        var style = context.Styles.GetValueOrDefault(styleId);
        if (style == null) return null;

        // スタイルのアウトラインがあれば利用
        var styleLv = style.StyleParagraphProperties?.OutlineLevel?.Val;
        if (styleLv?.HasValue == true) return styleLv.Value;

        return null;
    }

    /// <summary>段落テキストを抽出する</summary>
    /// <param name="context">抽出コンテキスト</param>
    /// <param name="paragraph">段落オブジェクト</param>
    /// <param name="writer">書き込み処理</param>
    /// <param name="options">抽出オプション</param>
    private void extractParagraphText(ExtractContext context, Paragraph paragraph, IWriterFacade writer, WordTextExtractorOptions? options = null)
    {
        // 配下の要素を処理
        foreach (var element in paragraph.Descendants<OpenXmlElement>())
        {
            switch (element)
            {
            case Text text: // テキスト
                writer.Write(text.Text);
                break;

            case Break:     // 改行
                writer.WriteLine();
                break;

            case FootnoteReference fnRef when options?.WithFootnote == true && fnRef.Id?.Value is long refId:   // 脚注ID
                context.UsedFootnotes.Add(refId);
                writer.Write($"[{refId}]");
                break;

            default:
                break;
            }
        }

        // 段落末尾に改行
        writer.WriteLine();
    }

    /// <summary>テーブルテキストを抽出する</summary>
    /// <param name="context">抽出コンテキスト</param>
    /// <param name="table">テーブルオブジェク</param>
    /// <param name="writer">書き込み処理</param>
    /// <param name="options">抽出オプション</param>
    private void extractTableText(ExtractContext context, Table table, IWriterFacade writer, WordTextExtractorOptions? options = null)
    {
        // 行を処理
        foreach (var row in table.Elements<TableRow>())
        {
            // セルを処理
            foreach (var cell in row.Elements<TableCell>())
            {
                writer.Write("|");
                foreach (var element in cell.Descendants<OpenXmlElement>())
                {
                    switch (element)
                    {
                    case Text text: // テキスト
                        writer.Write(text.Text);
                        break;

                    case Break:     // 改行
                        writer.Write(" ");
                        break;

                    case FootnoteReference fnRef when options?.WithFootnote == true && fnRef.Id?.Value is long refId:   // 脚注ID
                        context.UsedFootnotes.Add(refId);
                        writer.Write($"[{refId}]");
                        break;

                    default:
                        break;
                    }
                }
            }
            writer.WriteLine("|");
        }
    }

    /// <summary>脚注をテキストを抽出する</summary>
    /// <param name="context">抽出コンテキスト</param>
    /// <param name="writer">書き込み処理</param>
    private void extractFootnoteText(ExtractContext context, IWriterFacade writer)
    {
        // 脚注出力する内容がない場合は
        if (context.UsedFootnotes.Count <= 0) return;

        // 参照(使用)しているIDの脚注を出力
        foreach (var id in context.UsedFootnotes.Distinct())
        {
            // 脚注番号
            writer.Write($"[{id}]:");

            // 対応する脚注取得
            var lineEnded = false;
            var footnote = context.Footnotes.GetValueOrDefault(id);
            if (footnote == null)
            {
                // 脚注が見つからない場合はその旨を出力
                writer.Write(" ** Missing footnote. **");
            }
            else
            {
                // 脚注を出力
                foreach (var element in footnote.Descendants<OpenXmlElement>())
                {
                    switch (element)
                    {
                    case Text text:
                        if (text.Text is string elemText && 0 < elemText.Length)
                        {
                            writer.Write(text.Text);
                            lineEnded = elemText[^1] is '\r' or '\n';
                        }
                        break;

                    case Break:
                        writer.WriteLine();
                        lineEnded = true;
                        break;

                    default:
                        break;
                    }
                }
            }

            // 脚注の末尾に改行が無ければ付与しておく
            if (!lineEnded)
            {
                writer.WriteLine();
            }
        }
    }
    #endregion

}
