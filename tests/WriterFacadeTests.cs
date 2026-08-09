using System.Text;
using AwesomeAssertions;

namespace Tivarloa.Tests;

[TestClass]
public class WriterFacadeTests
{
    [TestMethod]
    public void StringBuilderFacade()
    {
        var now = DateTime.Now;

        var builder = new StringBuilder();
        var facade = new StringBuilderFacade(builder);
        facade.Write("ABC");
        facade.Write($"DEF:{now}");
        facade.WriteLine();
        facade.WriteLine("GHI");
        facade.WriteLine($"JKL:{now}");

        builder.ToString().Should().Be($"ABCDEF:{now}{Environment.NewLine}GHI{Environment.NewLine}JKL:{now}{Environment.NewLine}");
    }

    [TestMethod]
    public void TextWriterFacade()
    {
        var now = DateTime.Now;

        var writer = new StringWriter();
        var facade = new TextWriterFacade(writer);
        facade.Write("ABC");
        facade.Write($"DEF:{now}");
        facade.WriteLine();
        facade.WriteLine("GHI");
        facade.WriteLine($"JKL:{now}");

        writer.GetStringBuilder().ToString().Should().Be($"ABCDEF:{now}{Environment.NewLine}GHI{Environment.NewLine}JKL:{now}{Environment.NewLine}");
    }
}
