using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Tivarloa;

/// <summary>アウトラインブロック</summary>
/// <param name="Level">アウトラインレベル。</param>
/// <param name="Caption">アウトラインキャプション</param>
/// <param name="Content">アウトラインブロック内容</param>
public record WordOutlineBlock(int Level, string Caption, string? Number, string Content);

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
    #region データ管理
    /// <summary></summary>
    /// <param name="NumID"></param>
    /// <param name="Level"></param>
    private record struct NumberingRef(int NumID, int Level);
    #endregion

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

        /// <summary>指定のナンバリングIDに対する採番コンテキストを取得する</summary>
        /// <param name="numID">ナンバリングID</param>
        /// <returns>採番コンテキスト</returns>
        public NumberingContext? TryGetNumbering(int numID)
        {
            if (!this.numberingContextCache.TryGetValue(numID, out var context))
            {
                this.numberingContextCache[numID] = context = makeNumberingContext(numID);
            }
            return context;
        }

        /// <summary>スタイル定義辞書キャッシュ</summary>
        private readonly Lazy<Dictionary<string, Style>> styleCache = new(() =>
        {
            var styleDict = new Dictionary<string, Style>();
            var styles = document?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>();
            if (styles != null)
            {
                foreach (var style in styles)
                {
                    if (style.StyleId?.Value is string styleId) styleDict.Add(styleId, style);
                }
            }
            return styleDict;
        });

        /// <summary>脚注定義辞書キャッシュ</summary>
        private readonly Lazy<Dictionary<long, Footnote>> footnotesCache = new(() =>
        {
            var fnDict = new Dictionary<long, Footnote>();
            var footnotes = document.MainDocumentPart?.FootnotesPart?.Footnotes?.Elements<Footnote>();
            if (footnotes != null)
            {
                foreach (var fn in footnotes)
                {
                    if (fn.Id?.HasValue == true) fnDict.Add(fn.Id.Value, fn);
                }
            }
            return fnDict;
        });

        /// <summary>ナンバリングコンテキストキャッシュ</summary>
        private readonly Dictionary<int, NumberingContext?> numberingContextCache = new();

        /// <summary>ナンバリングコンテキストを生成する</summary>
        /// <param name="numID">ナンバリングID</param>
        /// <returns>ナンバリングコンテキスト</returns>
        public NumberingContext? makeNumberingContext(int numID)
        {
            // ナンバリングインスタンスとナンバリング定義の要素シーケンスを取得
            var numInsts = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.Elements<NumberingInstance>();
            var absNums = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering?.Elements<AbstractNum>();
            if (numInsts == null || absNums == null) return null;

            // IDに対応するナンバリングインスタンスを取得
            var numInst = numInsts.FirstOrDefault(n => n.NumberID?.HasValue == true && n.NumberID.Value == numID);
            if (numInst == null) return null;
            if (numInst.AbstractNumId?.Val?.HasValue != true) return null;

            // ナンバリングインスタンスが参照するナンバリング定義を取得
            var absId = numInst.AbstractNumId.Val.Value;
            var absNum = absNums.FirstOrDefault(a => a.AbstractNumberId?.HasValue == true && a.AbstractNumberId.Value == absId);
            if (absNum == null) return null;

            return new NumberingContext(numInst, absNum);
        }
    }

    /// <summary>採番コンテキスト</summary>
    /// <remarks>
    /// このコンテキストはナンバリングID(ナンバリングインスタンス)と対応し、階層的なナンバリングの採番を行うための状態・処理となる。
    /// 文書中の各ナンバリング箇所で、1回づつ <see cref="Advance(int)"/> を呼び出すことで順次採番を行う。
    /// </remarks>
    private class NumberingContext
    {
        /// <summary>採番コンテキスト</summary>
        /// <param name="inst">ナンバリングインスタンス</param>
        /// <param name="abs">ナンバリング定義</param>
        public NumberingContext(NumberingInstance inst, AbstractNum abs)
        {
            this.levels = new();
            this.startNumbers = new();

            foreach (var lv in abs.Elements<Level>())
            {
                if (lv.LevelIndex?.HasValue != true) continue;
                if (lv.StartNumberingValue?.Val?.HasValue != true) continue;
                var level = lv.LevelIndex.Value;
                if (level < 0) continue;
                while (this.startNumbers.Count < (level + 1)) this.startNumbers.Add(null);
                this.startNumbers[level] = lv.StartNumberingValue.Val.Value;
            }
            foreach (var ovLv in inst.Elements<LevelOverride>())
            {
                if (ovLv.LevelIndex?.HasValue != true) continue;
                if (ovLv.StartOverrideNumberingValue?.Val?.HasValue != true) continue;
                var level = ovLv.LevelIndex.Value;
                if (level < 0) continue;
                while (this.startNumbers.Count < (level + 1)) this.startNumbers.Add(null);
                this.startNumbers[level] = ovLv.StartOverrideNumberingValue.Val.Value;
            }
        }

        /// <summary>指定したレベルのナンバリング採番を行う</summary>
        /// <param name="level">ナンバリングレベル</param>
        /// <returns>構築されたナンバリング文字列</returns>
        public string Advance(int level)
        {
            // 無効なレベル値の場合は空文字列にしておく
            if (level < 0) return "";

            // 指定のレベルの番号を生成
            if (this.levels.Count < (level + 1))
            {
                // 指定レベルの番号が無い場合、新たに生成。
                while (this.levels.Count < (level + 1))
                {
                    // 開始番号の決定
                    var startVal = 1;
                    var refLevel = this.levels.Count;
                    if (refLevel <= this.startNumbers.Count)
                    {
                        // レベルに対する開始値を持っている場合、最初にそのレベルの採番を行う場合に利用。一度利用したらクリア。
                        var val = this.startNumbers[refLevel];
                        if (val.HasValue)
                        {
                            startVal = val.Value;
                            this.startNumbers[refLevel] = null;
                        }
                    }
                    // 開始番号
                    this.levels.Add(startVal);
                }
            }
            else
            {
                // 指定レベルより深い番号付けをしている状態の場合、指定レベルより詳細な番号付けレベルを削除
                while ((level + 1) < this.levels.Count) this.levels.RemoveAt(this.levels.Count - 1);
                // 指定レベルをインクリメント
                this.levels[level] = this.levels[level] + 1;
            }

            // 階層的な各レベル値から文字列構築
            var builder = new StringBuilder();
            for (var i = 0; i < this.levels.Count; i++)
            {
                if (i != 0) builder.Append(".");
                builder.Append(this.levels[i]);
            }
            return builder.ToString();
        }

        private readonly List<int> levels;
        private readonly List<int?> startNumbers;
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
                if (extractNumberingString(context, paragraph) is string numText)
                {
                    writer.Write(numText);
                    writer.Write(" ");
                }
                extractParagraphText(context, paragraph, writer, options);
                writer.WriteLine();
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
        var number = default(string);
        var caption = new StringBuilderFacade(new());
        var content = new StringBuilderFacade(new());

        // 現在の収集情報を確定するローカルメソッド
        void commitOutlineBlock()
        {
            if (0 < content.Builder.Length || 0 < caption.Builder.Length)
            {
                outline.Add(new(level, caption.Builder.ToString(), number, content.Builder.ToString()));
                level = -1;
                number = null;
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
                // アウトライン段落であるかを判定
                if (extractOutlineLevel(context, paragraph) is int lv)
                {
                    // アウトライン段落の場合、すでに収集中のアウトラインブロックがあれば確定する
                    commitOutlineBlock();
                    // 新しいアウトラインブロックを開始する
                    level = lv;
                    number = extractNumberingString(context, paragraph);
                    extractParagraphText(context, paragraph, caption, options);
                }
                else
                {
                    // 段落にナンバリングが設定されている場合、先頭に付与する。
                    if (extractNumberingString(context, paragraph) is string numText)
                    {
                        content.Write(numText);
                        content.Write(" ");
                    }
                    // アウトラインブロックコンテンツとして、段落テキストを追加
                    extractParagraphText(context, paragraph, content, options);
                    content.WriteLine();
                }
                break;

            case Table table:
                // アウトラインブロックコンテンツとして、テーブルテキストを追加
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
        // スタイル参照辞書
        var styleDict = context.Styles;

        // 段落プロパティの参照
        var paraProp = paragraph.ParagraphProperties;
        if (paraProp == null) return null;

        // 直接アウトライン設定されていれば利用
        var lv = paraProp.OutlineLevel?.Val;
        if (lv?.HasValue == true) return lv.Value;

        // 段落スタイルIDを取得
        var styleId = paraProp.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        // スタイルから情報参照
        while (true)
        {
            // 段落スタイルを参照
            var style = styleDict.GetValueOrDefault(styleId);
            if (style == null) break;

            // スタイルのアウトラインがあれば利用
            var styleLv = style.StyleParagraphProperties?.OutlineLevel?.Val;
            if (styleLv?.HasValue == true) return styleLv.Value;

            // 見つかっていない場合、継承元のスタイルを参照
            styleId = style.BasedOn?.Val?.HasValue == true ? style.BasedOn.Val.Value : default;
            if (styleId == null) break;
        }

        return null;
    }

    /// <summary>段落のナンバリング情報の取得を試みる</summary>
    /// <param name="context">抽出コンテキスト</param>
    /// <param name="paragraph">段落</param>
    /// <returns>段落がナンバリング情報を持つ場合はその値。そうでない場合は null を返却</returns>
    private string? extractNumberingString(ExtractContext context, Paragraph paragraph)
    {
        // ナンバリング参照情報を取得
        var num = extractNumbering(context, paragraph);
        if (num == null) return null;

        // ナンバリングコンテキスト取得
        var numContext = context.TryGetNumbering(num.Value.NumID);
        if (numContext == null) return null;

        // ナンバリング文字列構築
        return numContext.Advance(num.Value.Level);
    }

    /// <summary>段落のナンバリング情報の取得を試みる</summary>
    /// <param name="context">抽出コンテキスト</param>
    /// <param name="paragraph">段落</param>
    /// <returns>段落がナンバリング情報を持つ場合はその値。そうでない場合は null を返却</returns>
    private NumberingRef? extractNumbering(ExtractContext context, Paragraph paragraph)
    {
        // スタイル参照辞書
        var styleDict = context.Styles;

        // ナンバリング情報
        var numId = default(int?);
        var level = default(int?);

        // 段落プロパティの参照
        var styleId = default(string);
        var paraProp = paragraph.ParagraphProperties;
        if (paraProp != null)
        {
            // 直接ナンバリング設定されていれば利用
            var num = paraProp.NumberingProperties;
            if (num != null)
            {
                numId ??= num.NumberingId?.Val?.HasValue == true ? num.NumberingId.Val.Value : null;
                level ??= num.NumberingLevelReference?.Val?.HasValue == true ? num.NumberingLevelReference.Val.Value : null;
            }

            // スタイル取得
            styleId = paraProp.ParagraphStyleId?.Val?.Value;
        }

        // スタイルから
        while (true)
        {
            // 情報が揃っていれば返却
            if (numId.HasValue && level.HasValue)
            {
                return new(numId.Value, level.Value);
            }

            // 参照するスタイルが無ければ探査終わり
            if (styleId == null) break;

            // 段落スタイルを参照
            var style = context.Styles.GetValueOrDefault(styleId);
            if (style == null) break;

            // スタイルのアウトラインがあれば利用
            var styleNum = style.StyleParagraphProperties?.NumberingProperties;
            if (styleNum != null)
            {
                numId ??= styleNum.NumberingId?.Val?.HasValue == true ? styleNum.NumberingId.Val.Value : null;
                level ??= styleNum.NumberingLevelReference?.Val?.HasValue == true ? styleNum.NumberingLevelReference.Val.Value : null;
            }

            // 次に継承元のスタイルを参照
            styleId = style.BasedOn?.Val?.HasValue == true ? style.BasedOn.Val.Value : default;
        }

        // ナンバリング参照が見つかっていれば、レベルはデフォルト値ゼロとみなす。
        if (numId.HasValue)
        {
            return new(numId.Value, level ?? 0);
        }

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
