# Tivarloa

A library for extracting plain text from documents.  
This targets Word, PDF, and HTML.  
This is primarily intended for my personal use, so the text conversion isn't particularly precise.  


```csharp
var extractor = new PdfTextExtractor();
var allText = extractor.Extract(file, new(UseActualText: true));
```

```csharp
var extractor = new HtmlTextExtractor();
var allText = await extractor.ExtractAsync(url, new(Trimming: true));
```

```csharp
var extractor = new WordTextExtractor();
var allText = extractor.Extract(file, new(WithFootnote: true));
```

```csharp
var extractor = new WordTextExtractor();
var outline = extractor.ExtractOutline(file);
foreach (var section in outline)
{
    Console.WriteLine($"{section.Number} {section.Caption}");
    Console.WriteLine($"{section.Content}");
}
```