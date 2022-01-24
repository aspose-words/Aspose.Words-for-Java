package DocsExamples.Programming_with_Documents.Split_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.Section;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.IO.SearchOption;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.msString;
import com.aspose.words.ImportFormatMode;


class SplitDocument extends DocsExamplesBase
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
            newDoc.save(getArtifactsDir() + $"SplitDocument.BySections_{i}.docx");
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
            extractedPage.save(getArtifactsDir() + $"SplitDocument.PageByPage_{page + 1}.docx");
        }
        //ExEnd:SplitDocumentPageByPage

        mergeDocuments();
    }

    //ExStart:MergeSplitDocuments
    private void mergeDocuments() throws Exception
    {
        // Find documents using for merge.
        FileSystemInfo[] documentPaths = new DirectoryInfo(getArtifactsDir())
            .GetFileSystemInfos("SplitDocument.PageByPage_*.docx").OrderBy(f => f.CreationTime).ToArray();
        String sourceDocumentPath =
            Directory.getFiles(getArtifactsDir(), "SplitDocument.PageByPage_1.docx", SearchOption.TOP_DIRECTORY_ONLY)[0];

        // Open the first part of the resulting document.
        Document sourceDoc = new Document(sourceDocumentPath);

        // Create a new resulting document.
        Document mergedDoc = new Document();
        DocumentBuilder mergedDocBuilder = new DocumentBuilder(mergedDoc);

        // Merge document parts one by one.
        for (FileSystemInfo documentPath : documentPaths)
        {
            if (msString.equals(documentPath.getFullName(), sourceDocumentPath))
                continue;

            mergedDocBuilder.moveToDocumentEnd();
            mergedDocBuilder.insertDocument(sourceDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            sourceDoc = new Document(documentPath.getFullName());
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


