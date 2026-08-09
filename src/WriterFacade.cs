using System.Runtime.CompilerServices;
using System.Text;

namespace Tivarloa;

/// <summary>テキスト書き込み処理抽象化インタフェース</summary>
internal interface IWriterFacade
{
    /// <summary>テキストを書き込む</summary>
    /// <param name="text">テキスト</param>
    public void Write(ReadOnlySpan<char> text);

    /// <summary>テキストを書き込む</summary>
    /// <param name="handler">文字列補間ハンドラ</param>
    public void Write([InterpolatedStringHandlerArgument("")] ref FacadeInterpolatedStringHandler handler) { }

    /// <summary>テキスト行を書き込む</summary>
    /// <param name="text">テキスト</param>
    public void WriteLine(ReadOnlySpan<char> text);

    /// <summary>テキスト行を書き込む</summary>
    /// <param name="handler">文字列補間ハンドラ</param>
    public void WriteLine([InterpolatedStringHandlerArgument("")] ref FacadeInterpolatedStringHandler handler) { this.WriteLine(); }

    /// <summary>改行を書き込む</summary>
    public void WriteLine();

    /// <summary>値を書き込む</summary>
    /// <typeparam name="T">値の型</typeparam>
    /// <param name="value">値</param>
    internal void Write<T>(T value);

    /// <summary>文字列補間ハンドラ</summary>
    [InterpolatedStringHandler]
    public struct FacadeInterpolatedStringHandler
    {
        /// <summary>抽象化インスタンスと紐づけるコンストラクタ</summary>
        /// <param name="literalLength">利用しない</param>
        /// <param name="formattedCount">利用しない</param>
        /// <param name="facade">書き込み抽象化インスタンス</param>
        public FacadeInterpolatedStringHandler(int literalLength, int formattedCount, IWriterFacade facade)
        {
            this.facade = facade;
        }

        /// <inheritdoc />
        public void AppendLiteral(string literal) => this.facade.Write(literal);

        /// <inheritdoc />
        public void AppendFormatted<T>(T value) => this.facade.Write(value);

        /// <summary>書き込み抽象化インスタンス</summary>
        private IWriterFacade facade;
    }
}

/// <summary>StringBuilder をバックエンドにするテキスト書き込み処理</summary>
/// <param name="builder">バックエンド StringBuilder</param>
internal class StringBuilderFacade(StringBuilder builder) : IWriterFacade
{
    /// <summary>バックエンド StringBuilder</summary>
    public StringBuilder Builder => builder;

    /// <inheritdoc />
    public void Write(ReadOnlySpan<char> text)
    {
        builder.Append(text);
    }

    /// <inheritdoc />
    public void WriteLine(ReadOnlySpan<char> text)
    {
        builder.Append(text);
        builder.AppendLine();
    }

    /// <inheritdoc />
    public void WriteLine()
    {
        builder.AppendLine();
    }

    /// <inheritdoc />
    void IWriterFacade.Write<T>(T value)
    {
        builder.Append(value);
    }
}

/// <summary>TextWriter をバックエンドにするテキスト書き込み処理</summary>
/// <param name="writer">バックエンド TextWriter</param>
internal class TextWriterFacade(TextWriter writer) : IWriterFacade
{
    /// <summary>バックエンド TextWriter</summary>
    public TextWriter Writer => writer;

    /// <inheritdoc />
    public void Write(ReadOnlySpan<char> text)
    {
        writer.Write(text);
    }

    /// <inheritdoc />
    public void WriteLine(ReadOnlySpan<char> text)
    {
        writer.WriteLine(text);
    }

    /// <inheritdoc />
    public void WriteLine()
    {
        writer.WriteLine();
    }

    /// <inheritdoc />
    void IWriterFacade.Write<T>(T value)
    {
        writer.Write(value);
    }
}