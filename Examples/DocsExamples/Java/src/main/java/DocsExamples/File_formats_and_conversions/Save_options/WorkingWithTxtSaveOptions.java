package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class WorkingWithTxtSaveOptions extends DocsExamplesBase {
    @Test
    public void addBidiMarks() throws Exception {
        //ExStart:AddBidiMarks
        //GistId:c92d84644de8ee6e7148950debea90d6
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.getParagraphFormat().setBidi(true);
        builder.writeln("שלום עולם!");
        builder.writeln("مرحبا بالعالم!");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.setAddBidiMarks(true);

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
        //ExEnd:AddBidiMarks
    }

    @Test
    public void useTabForListIndentation() throws Exception {
        //ExStart:UseTabForListIndentation
        //GistId:c92d84644de8ee6e7148950debea90d6
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list with three levels of indentation.
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent();
        builder.write("Item 3");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.getListIndentation().setCount(1);
        saveOptions.getListIndentation().setCharacter('\t');

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.UseTabForListIndentation.txt", saveOptions);
        //ExEnd:UseTabForListIndentation
    }

    @Test
    public void useSpaceForListIndentation() throws Exception {
        //ExStart:UseSpaceForListIndentation
        //GistId:c92d84644de8ee6e7148950debea90d6
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list with three levels of indentation.
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent();
        builder.write("Item 3");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.getListIndentation().setCount(3);
        saveOptions.getListIndentation().setCharacter(' ');

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.UseSpaceForListIndentation.txt", saveOptions);
        //ExEnd:UseSpaceForListIndentation
    }

    @Test
    public void exportHeadersFootersMode() throws Exception {
        //ExStart:ExportHeadersFootersMode
        //GistId:c92d84644de8ee6e7148950debea90d6
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
