using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Tivarloa.Tests._Helper;

public static class WordDocHelper
{
    public static FileInfo CreateTestDocument(FileInfo file)
    {
        using var document = WordprocessingDocument.Create(file.FullName, WordprocessingDocumentType.Document);

        // メインパート
        var mainBody = new Body();
        var mainDoc = new Document(mainBody);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = mainDoc;

        // ナンバリング定義パート
        var numPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numPart.Numbering = new Numbering(
            // ナンバリング定義
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%%[%1]%%" }
                )
                { LevelIndex = 0 },
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "<%1 - %2>" }
                )
                { LevelIndex = 1 }
            )
            { AbstractNumberId = 0 },
            // ナンバリングインスタンス
            new NumberingInstance(
                new AbstractNumId { Val = 0 }
            )
            { NumberID = 1 }
        );

        // スタイル定義
        var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles = new Styles(
            // 見出しレベル１ - Heading1
            new Style(
                new StyleName { Val = "heading 1" },
                new StyleParagraphProperties(
                    new NumberingProperties(
                        new NumberingId { Val = 1 },
                        new NumberingLevelReference { Val = 0 }
                    )
                )
            )
            { Type = StyleValues.Paragraph, StyleId = "Heading1" },
            // 見出しレベル２ - Heading2
            new Style(
                new StyleName { Val = "heading 2" },
                new BasedOn { Val = "Heading1" }, // Heading1 を継承
                new StyleParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 1 } // ilvl を上書き
                    )
                )
            )
            { Type = StyleValues.Paragraph, StyleId = "Heading2" }
        );

        // 本文 
        mainBody.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }, new OutlineLevel { Val = 0, }),
            new Run(new Text("Chapter A"))
        ));
        mainBody.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }, new OutlineLevel { Val = 1, }),
            new Run(new Text("Section X"))
        ));
        mainBody.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }, new OutlineLevel { Val = 1, }),
            new Run(new Text("Section Y"))
        ));
        mainBody.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }, new OutlineLevel { Val = 0, }),
            new Run(new Text("Chapter B"))
        ));

        // 保存
        mainPart.Document.Save();

        return file;
    }

}
