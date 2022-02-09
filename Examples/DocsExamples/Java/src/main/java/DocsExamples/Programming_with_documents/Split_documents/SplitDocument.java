package DocsExamples.Programming_with_documents.Split_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.testng.annotations.Test;

import java.io.File;
import java.text.MessageFormat;
import java.util.Collection;

@Test
public class SplitDocument extends DocsExamplesBase
{
    @Test
    public void byHeadingsHtml() throws Exception
    {
        //ExStart:SplitDocumentByHeadingsHtml
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            // Split a document into smaller parts, in this instance split by heading.
            options.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);
        }
        

        doc.save(getArtifactsDir() + "SplitDocument.ByHeadingsHtml.html", options);
        //ExEnd:SplitDocumentByHeadingsHtml
    }

    @Test
    public void bySectionsHtml() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
 
        //ExStart:SplitDocumentBySectionsHtml
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setDocumentSplitCriteria(DocumentSplitCriteria.SECTION_BREAK); }
        //ExEnd:SplitDocumentBySectionsHtml
        
        doc.save(getArtifactsDir() + "SplitDocument.BySectionsHtml.html", options);
    }

    @Test
    public void bySections() throws Exception
    {
        //ExStart:SplitDocumentBySections
        Document doc = new Document(getMyDir() + "Big document.docx");

        for (int i = 0; i < doc.getSections().getCount(); i++)
        {
            // Split a document into smaller parts, in this instance, split by section.
            Section section = doc.getSections().get(i).deepClone();

            Document newDoc = new Document();
            newDoc.getSections().clear();

            Section newSection = (Section) newDoc.importNode(section, true);
            newDoc.getSections().add(newSection);

            // Save each section as a separate document.
            newDoc.save(getArtifactsDir() + MessageFormat.format("SplitDocument.BySections_{0}.docx", i));
        }
        //ExEnd:SplitDocumentBySections
    }

    @Test
    public void pageByPage() throws Exception
    {
        //ExStart:SplitDocumentPageByPage
        Document doc = new Document(getMyDir() + "Big document.docx");

        int pageCount = doc.getPageCount();

        for (int page = 0; page < pageCount; page++)
        {
            // Save each page as a separate document.
            Document extractedPage = doc.extractPages(page, 1);
            extractedPage.save(getArtifactsDir() + MessageFormat.format("SplitDocument.PageByPage_{0}.docx", page + 1));
        }
        //ExEnd:SplitDocumentPageByPage

        mergeDocuments();
    }

    //ExStart:MergeSplitDocuments
    private void mergeDocuments() throws Exception
    {
        // Find documents using for merge.
        File directory = new File(getArtifactsDir());
        Collection<File> documentPaths = FileUtils.listFiles(directory, new WildcardFileFilter("SplitDocument.PageByPage_*.docx"), null);

        String sourceDocumentPath =
                FileUtils.getFile(getArtifactsDir(), "SplitDocument.PageByPage_1.docx").getPath();

        // Open the first part of the resulting document.
        Document sourceDoc = new Document(sourceDocumentPath);

        // Create a new resulting document.
        Document mergedDoc = new Document();
        DocumentBuilder mergedDocBuilder = new DocumentBuilder(mergedDoc);

        // Merge document parts one by one.
        for (File documentPath : documentPaths)
        {
            if (documentPath.getName().equals(sourceDocumentPath))
                continue;

            mergedDocBuilder.moveToDocumentEnd();
            mergedDocBuilder.insertDocument(sourceDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            sourceDoc = new Document(documentPath.getPath());
        }

        mergedDoc.save(getArtifactsDir() + "SplitDocument.MergeDocuments.docx");
    }
    //ExEnd:MergeSplitDocuments

    @Test
    public void byPageRange() throws Exception
    {
        //ExStart:SplitDocumentByPageRange
        Document doc = new Document(getMyDir() + "Big document.docx");
        
        // Get part of the document.
        Document extractedPages = doc.extractPages(3, 6);
        extractedPages.save(getArtifactsDir() + "SplitDocument.ByPageRange.docx");
        //ExEnd:SplitDocumentByPageRange
    }
}