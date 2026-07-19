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
/// <param name="WhitespaceCompaction">空白文字を単一のスペースに置き換える</param>
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
    public async ValueTask ExtractWriteAsync(Uri url, StringBuilder builder, WordTextExtractorOptions? options = null, CancellationToken cancelToken = default)
    {
        using var downloadStream = await this.http.Value.GetStreamAsync(url, cancelToken);
        ExtractWrite(downloadStream, builder, options);
    }

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="builder">書き込み先文字列ビルダ</param>
    /// <param name="options">抽出オプション</param>
    public void ExtractWrite(FileInfo file, StringBuilder builder, WordTextExtractorOptions? options = null)
    {
        using var fileStream = file.OpenRead();
        ExtractWrite(fileStream, builder, options);
    }

    /// <summary>Word文書全体のテキストを文字列ビルダに書き込む</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="builder">書き込み先文字列ビルダ</param>
    /// <param name="options">抽出オプション</param>
    public void ExtractWrite(Stream stream, StringBuilder builder, WordTextExtractorOptions? options = null)
    {
        // 文書オープン
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        if (doc.MainDocumentPart == null) return;

        // 本文参照
        var docBody = doc.MainDocumentPart.Document?.Body;
        if (docBody == null) return;

        // 文書のテキスト化
        var footnoteIds = new List<long>();
        foreach (var element in docBody.Elements<OpenXmlElement>())
        {
            switch (element)
            {
            case Paragraph paragraph:   // 段落
                extractParagraphText(paragraph, builder, footnoteIds, options);
                break;

            case Table table:           // テーブル
                extractTableText(table, builder, footnoteIds, options);
                break;

            default:
                break;
            }
        }

        // 脚注出力が有効であればテキスト化
        if (options?.WithFootnote == true)
        {
            builder.AppendLine("----------------------------------------");
            extractFootnoteText(doc.MainDocumentPart.FootnotesPart, footnoteIds, builder);
        }
    }
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
        return ExtractOutline(downloadStream, options);
    }
    /// <summary>Word文書全体をアウトラインブロック毎にテキスト化する</summary>
    /// <param name="file">Word文書ファイル</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>アウトライン毎のテキスト情報</returns>
    public List<WordOutlineBlock> ExtractOutline(FileInfo file, WordTextExtractorOptions? options = null)
    {
        using var fileStream = file.OpenRead();
        return ExtractOutline(fileStream, options);
    }

    /// <summary>Word文書全体をアウトラインブロック毎にテキスト化する</summary>
    /// <param name="stream">Word文書ストリーム</param>
    /// <param name="options">抽出オプション</param>
    /// <returns>アウトライン毎のテキスト情報</returns>
    public List<WordOutlineBlock> ExtractOutline(Stream stream, WordTextExtractorOptions? options = null)
    {
        // 文書オープン
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        if (doc.MainDocumentPart == null) return [];

        // 本文参照
        var docBody = doc.MainDocumentPart.Document?.Body;
        if (docBody == null) return [];

        // スタイル参照用に辞書化
        var styles = doc.MainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>()
            ?.Where(s => s.StyleId?.Value != null)
            ?.ToDictionary(s => s.StyleId!.Value!) ?? [];

        // アウトライン毎のテキスト
        var outline = new List<WordOutlineBlock>();

        // 現在の収集情報
        var level = -1;
        var caption = new StringBuilder();
        var content = new StringBuilder();

        // 現在の収集情報を確定するローカルメソッド
        void commitOutlineBlock()
        {
            if (0 < content.Length || 0 < caption.Length)
            {
                outline.Add(new(level, caption.ToString(), content.ToString()));
                level = -1;
                caption.Clear();
                content.Clear();
            }
        }

        // 文書の各アウトラインブロックをテキスト化
        var footnoteIds = new List<long>();
        foreach (var element in docBody.Elements<OpenXmlElement>())
        {
            switch (element)
            {
            case Paragraph paragraph:
                if (extractOutlineLevel(paragraph, styles) is int lv)
                {
                    commitOutlineBlock();
                    level = lv;
                    extractParagraphText(paragraph, caption, footnoteIds, options);
                }
                else
                {
                    extractParagraphText(paragraph, content, footnoteIds, options);
                }
                break;

            case Table table:
                extractTableText(table, content, footnoteIds, options);
                break;

            default:
                break;
            }
        }

        // 最後のブロックを確定
        commitOutlineBlock();

        // 脚注出力が有効であればテキスト化
        if (options?.WithFootnote == true && 0 < footnoteIds.Count)
        {
            level = -1;
            caption.Append("Footnotes");
            extractFootnoteText(doc.MainDocumentPart.FootnotesPart, footnoteIds, content);
            commitOutlineBlock();
        }

        return outline;
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

    #region ドキュメント処理
    /// <summary>段落のアウトラインレベルの取得を試みる</summary>
    /// <param name="paragraph">段落</param>
    /// <param name="styles">ドキュメントで定義されているスタイルの辞書</param>
    /// <returns>段落がアウトラインレベルを持つ場合はそのレベル値。そうでない場合は null を返却</returns>
    private int? extractOutlineLevel(Paragraph paragraph, Dictionary<string, Style> styles)
    {
        // 段落プロパティの参照
        var paraProp = paragraph.ParagraphProperties;
        if (paraProp == null) return null;

        // 直接アウトライン設定されていれば利用
        var lv = paraProp.GetFirstChild<OutlineLevel>()?.Val?.Value;
        if (lv != null) return lv.Value;

        // 段落スタイルIDを取得
        var styleId = paraProp.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        // 段落スタイルを参照
        var style = styles.GetValueOrDefault(styleId);
        if (style == null) return null;

        // スタイルのアウトラインがあれば利用
        var styleLv = style.StyleParagraphProperties?.GetFirstChild<OutlineLevel>()?.Val?.Value;
        return styleLv;
    }

    /// <summary>段落テキストを抽出する</summary>
    /// <param name="paragraph">段落オブジェクト</param>
    /// <param name="builder">書き込み先バッファ</param>
    /// <param name="footnoteIds">参照している脚注番号を格納するリスト</param>
    /// <param name="options">抽出オプション</param>
    private void extractParagraphText(Paragraph paragraph, StringBuilder builder, List<long> footnoteIds, WordTextExtractorOptions? options = null)
    {
        // 配下の要素を処理
        foreach (var element in paragraph.Descendants<OpenXmlElement>())
        {
            switch (element)
            {
            case Text text: // テキスト
                builder.Append(text.Text);
                break;

            case Break:     // 改行
                builder.AppendLine();
                break;

            case FootnoteReference fnRef when options?.WithFootnote == true && fnRef.Id?.Value is long refId:   // 脚注ID
                footnoteIds.Add(refId);
                builder.Append($"[{refId}]");
                break;

            default:
                break;
            }
        }

        // 段落末尾に
        builder.AppendLine();
    }

    /// <summary>テーブルテキストを抽出する</summary>
    /// <param name="table">テーブルオブジェク</param>
    /// <param name="builder">書き込み先バッファ</param>
    /// <param name="footnoteIds">参照している脚注番号を格納するリスト</param>
    /// <param name="options">抽出オプション</param>
    private void extractTableText(Table table, StringBuilder builder, List<long> footnoteIds, WordTextExtractorOptions? options = null)
    {
        // 行を処理
        foreach (var row in table.Elements<TableRow>())
        {
            // セルを処理
            foreach (var cell in row.Elements<TableCell>())
            {
                builder.Append('|');
                foreach (var element in cell.Descendants<OpenXmlElement>())
                {
                    switch (element)
                    {
                    case Text text: // テキスト
                        builder.Append(text.Text);
                        break;

                    case Break:     // 改行
                        builder.Append(' ');
                        break;

                    case FootnoteReference fnRef when options?.WithFootnote == true && fnRef.Id?.Value is long refId:   // 脚注ID
                        footnoteIds.Add(refId);
                        builder.Append($"[{refId}]");
                        break;

                    default:
                        break;
                    }
                }
            }
            builder.AppendLine("|");
        }
    }

    /// <summary>脚注をテキストを抽出する</summary>
    /// <param name="footnotes">ドキュメントの脚注</param>
    /// <param name="usedIds">参照されている脚注IDリスト</param>
    /// <param name="builder">書き込み先バッファ</param>
    private void extractFootnoteText(FootnotesPart? footnotes, List<long> usedIds, StringBuilder builder)
    {
        // 脚注出力する内容がない場合は
        if (footnotes?.Footnotes == null) return;
        if (usedIds.Count <= 0) return;

        // 脚注をIDで引けるように辞書化
        var footnoteDict = footnotes.Footnotes.Elements<Footnote>().Where(f => f.Id?.HasValue == true).ToDictionary(f => f.Id!.Value!);

        // 参照(使用)しているIDの脚注を出力
        foreach (var id in usedIds.Distinct())
        {
            // 脚注番号
            builder.Append($"[{id}]:");

            // 対応する脚注取得
            var footnote = footnoteDict.GetValueOrDefault(id);
            if (footnote == null)
            {
                // 脚注が見つからない場合はその旨を出力
                builder.Append(" ** Missing footnote. **");
            }
            else
            {
                // 脚注を出力
                foreach (var element in footnote.Descendants<OpenXmlElement>())
                {
                    switch (element)
                    {
                    case Text text:
                        builder.Append(text.Text);
                        break;

                    case Break:
                        builder.AppendLine();
                        break;

                    default:
                        break;
                    }
                }
            }

            // 脚注の末尾に改行が無ければ付与しておく
            if (0 < builder.Length && builder[^1] is not '\r' and not '\n')
            {
                builder.AppendLine();
            }
        }
    }
    #endregion

}
