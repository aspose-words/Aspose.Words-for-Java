package DocsExamples.File_Formats_and_Conversions.Load_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.TxtLoadOptions;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.TxtLeadingSpacesOptions;
import com.aspose.words.TxtTrailingSpacesOptions;
import com.aspose.words.DocumentDirection;
import com.aspose.words.Paragraph;
import com.aspose.ms.System.msConsole;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.TxtSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.TxtExportHeadersFootersMode;


public class WorkingWithTxtLoadOptions extends DocsExamplesBase
{
    @Test
    public void detectNumberingWithWhitespaces() throws Exception
    {
        //ExStart:DetectNumberingWithWhitespaces
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        // Create a plaintext document in the form of a string with parts that may be interpreted as lists.
        // Upon loading, the first three lists will always be detected by Aspose.Words,
        // and List objects will be created for them after loading.
        final String TEXT_DOC = "Full stop delimiters:\n" +
                               "1. First list item 1\n" +
                               "2. First list item 2\n" +
                               "3. First list item 3\n\n" +
                               "Right bracket delimiters:\n" +
                               "1) Second list item 1\n" +
                               "2) Second list item 2\n" +
                               "3) Second list item 3\n\n" +
                               "Bullet delimiters:\n" +
                               "• Third list item 1\n" +
                               "• Third list item 2\n" +
                               "• Third list item 3\n\n" +
                               "Whitespace delimiters:\n" +
                               "1 Fourth list item 1\n" +
                               "2 Fourth list item 2\n" +
                               "3 Fourth list item 3";

        // The fourth list, with whitespace inbetween the list number and list item contents,
        // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
        // to avoid paragraphs that start with numbers being mistakenly detected as lists.
        TxtLoadOptions loadOptions = new TxtLoadOptions(); { loadOptions.setDetectNumberingWithWhitespaces(true); }

        // Load the document while applying LoadOptions as a parameter and verify the result.
        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(TEXT_DOC)), loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithTxtLoadOptions.DetectNumberingWithWhitespaces.docx");
        //ExEnd:DetectNumberingWithWhitespaces
    }

    @Test
    public void handleSpacesOptions() throws Exception
    {
        //ExStart:HandleSpacesOptions
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        final String TEXT_DOC = "      Line 1 \n" +
                               "    Line 2   \n" +
                               " Line 3       ";

        TxtLoadOptions loadOptions = new TxtLoadOptions();
        {
            loadOptions.setLeadingSpacesOptions(TxtLeadingSpacesOptions.TRIM);
            loadOptions.setTrailingSpacesOptions(TxtTrailingSpacesOptions.TRIM);
        }

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(TEXT_DOC)), loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithTxtLoadOptions.HandleSpacesOptions.docx");
        //ExEnd:HandleSpacesOptions
    }

    @Test
    public void documentTextDirection() throws Exception
    {
        //ExStart:DocumentTextDirection
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        TxtLoadOptions loadOptions = new TxtLoadOptions(); { loadOptions.setDocumentDirection(DocumentDirection.AUTO); }

        Document doc = new Document(getMyDir() + "Hebrew text.txt", loadOptions);

        Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();
        msConsole.writeLine(paragraph.getParagraphFormat().getBidi());

        doc.save(getArtifactsDir() + "WorkingWithTxtLoadOptions.DocumentTextDirection.docx");
        //ExEnd:DocumentTextDirection
    }

    @Test
    public void exportHeadersFootersMode() throws Exception
    {
        //ExStart:ExportHeadersFootersMode
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        Document doc = new Document();

        // Insert even and primary headers/footers into the document.
        // The primary header/footers will override the even headers/footers.
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.HEADER_EVEN));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN).appendParagraph("Even header");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.FOOTER_EVEN));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN).appendParagraph("Even footer");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.HEADER_PRIMARY));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).appendParagraph("Primary header");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.FOOTER_PRIMARY));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).appendParagraph("Primary footer");

        // Insert pages to display these headers and footers.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Page 3");

        TxtSaveOptions options = new TxtSaveOptions();
        options.setSaveFormat(SaveFormat.TEXT);

        // All headers and footers are placed at the very end of the output document.
        options.setExportHeadersFootersMode(TxtExportHeadersFootersMode.ALL_AT_END);
        doc.save(getArtifactsDir() + "WorkingWithTxtLoadOptions.HeadersFootersMode.AllAtEnd.txt", options);

        // Only primary headers and footers are exported at the beginning and end of each section.
        options.setExportHeadersFootersMode(TxtExportHeadersFootersMode.PRIMARY_ONLY);
        doc.save(getArtifactsDir() + "WorkingWithTxtLoadOptions.HeadersFootersMode.PrimaryOnly.txt", options);

        // No headers and footers are exported.
        options.setExportHeadersFootersMode(TxtExportHeadersFootersMode.NONE);
        doc.save(getArtifactsDir() + "WorkingWithTxtLoadOptions.HeadersFootersMode.None.txt", options);
        //ExEnd:ExportHeadersFootersMode
    }
}

