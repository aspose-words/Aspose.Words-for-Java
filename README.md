![GitHub all releases](https://img.shields.io/github/downloads/aspose-words/Aspose.Words-for-Java/total) ![GitHub](https://img.shields.io/github/license/aspose-words/Aspose.Words-for-java)
# Java API for Various Document Formats

[Aspose.Words for Java](https://products.aspose.com/words/java) is an advanced Java Word processing API that enables you to perform a great range of document processing tasks directly within your Java applications. Aspose.Words for Java API supports processing word (DOC, DOCX, OOXML, RTF) HTML, OpenDocument, PDF, EPUB, XPS, SWF and all image formats. With Aspose.Words you can generate, modify, and convert documents without using Microsoft Word.

Directory | Description
--------- | -----------
[Examples](Examples) | A collection of Java examples that help you learn the product features.
[Plugins](Plugins) | Plugins that will demonstrate one or more features of Aspose.Words for Java.

<p align="center">
  <a title="Download Examples ZIP" href="https://github.com/aspose-words/Aspose.words-for-Java/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

## Word API Features

### Rendering and Printing

- Layout document into pages with high fidelity (exactly like Microsoft Word® would do that) to all the formats below.
- Render individual pages or complete documents to `PDF`, `XPS`, or `SWF`.
- Render document pages to raster images (Multipage `TIFF`, `PNG`, `JPEG`, `BMP`).
- Render pages to a Java Graphics object to a specific size.
- Print document pages using the Java printing infrastructure.
- Update TOC, page numbers, and other fields before rendering or printing.
- 3D Effects Rendering through the `OpenGL`.

### Document Content Features

- Access, create, and modify various document elements.
- Access and modify all document elements using `XmlDocument` -like classes and methods.
- Copy and move document elements between documents.
- Join and split documents.
- Specify document protection, open protected, and encrypted documents.
- Find and replace text, enumerate over document content.
- Preserve or extract OLE objects and ActiveX controls from the document.
- Preserve or remove VBA macros from the document. Preserve VBA macros digital signature.

### Reporting Features

- Support of C# syntax and LINQ extension methods directly in templates (even for `ADO.NET` data sources).
- Support of repeatable and conditional document blocks (loops and conditions) for tables, lists, and common content.
- Support of dynamically generated charts and images.
- Support of insertion of outer documents and `HTML` blocks into a document.
- Support of multiple data sources (including of different types) for the generation of a single document.
- Built-in support of data relations (master-detail).
- Comprehensive support of various data manipulations such as grouping, sorting, filtering, and others directly in templates.

For a more comprehensive list of features, please visit [Feature Overview](https://docs.aspose.com/words/java/feature-overview/).

## Read & Write Document Formats

**Microsoft Word:** DOC, DOCX, RTF, DOT, DOTX, DOTM, DOCM FlatOPC, FlatOpcMacroEnabled, FlatOpcTemplate, FlatOpcTemplateMacroEnabled\
**OpenOffice:** ODT, OTT\
**WordprocessingML:** WordML\
**Web:** HTML, MHTML\
**Fixed Layout:** PDF\
**Text:** TXT
**Other:** MD

## Save Word Files As

**Fixed Layout:** XPS, OpenXPS, PostScript (PS)\
**Images:** TIFF, JPEG, PNG, BMP, SVG, EMF, GIF\
**Web:** HtmlFixed\
**Others:** PCL, EPUB, XamlFixed, XamlFlow, XamlFlowPack

## Read File Formats

**MS Office:** DocPreWord60
**eBook:** MOBI

## Supported Environments

- **Microsoft Windows:** Windows Desktop & Server (x86, x64)
- **macOS:** Mac OS X
- **Linux:** Ubuntu, OpenSUSE, CentOS, and others
- **Java Versions:** `J2SE 7.0 (1.7)`, `J2SE 8.0 (1.8)` or above.

## Get Started with Aspose.Words for Java

Aspose hosts all Java APIs at the [Aspose Repository](https://repository.aspose.com/webapp/#/artifacts/browse/tree/General/repo/com/aspose/aspose-words). You can easily use Aspose.Words for Java API directly in your Maven projects with simple configurations. For the detailed instructions please visit [Installing Aspose.Words for Java from Maven Repository](https://docs.aspose.com/words/java/installation/) documentation page.

## Printing Multiple Pages on One Sheet using Java

```java
// Open the document.
Document doc = new Document(dataDir + "TestFile.doc");

// Create a print job to print our document with.
PrinterJob pj = PrinterJob.getPrinterJob();

// Initialize an attribute set with the number of pages in the document.
PrintRequestAttributeSet attributes = new HashPrintRequestAttributeSet();
attributes.add(new PageRanges(1, doc.getPageCount()));

// Pass the printer settings along with the other parameters to the print document.
MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true, attributes);

// Pass the document to be printed using the print job.
pj.setPrintable(awPrintDoc);

pj.print();
```

[Product Page](https://products.aspose.com/words/java) | [Docs](https://docs.aspose.com/words/java/) | [Demos](https://products.aspose.app/words/family) | [API Reference](https://apireference.aspose.com/words/java) | [Examples](https://github.com/aspose-words/Aspose.Words-for-Java/tree/master/Examples) | [Blog](https://blog.aspose.com/category/words/) | [Search](https://search.aspose.com/) | [Free Support](https://forum.aspose.com/c/words) | [Temporary License](https://purchase.aspose.com/temporary-license)
