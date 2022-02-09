package DocsExamples.Programming_with_documents.Split_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import com.aspose.words.ref.Ref;
import org.testng.annotations.Test;

import java.io.File;
import java.util.ArrayList;

@Test
public class SplitIntoHtmlPages extends DocsExamplesBase
{
    @Test
    public void htmlPages() throws Exception
    {
        String srcFileName = getMyDir() + "Footnotes and endnotes.docx";
        String tocTemplate = getMyDir() + "Table of content template.docx";

        String outDir = getArtifactsDir() + "HtmlPages";
        new File(outDir).mkdir();

        WordToHtmlConverter w = new WordToHtmlConverter();
        w.execute(srcFileName, tocTemplate, outDir);
    }
}

class WordToHtmlConverter
{
    /// <summary>
    /// Performs the Word to HTML conversion.
    /// </summary>
    /// <param name="srcFileName">The MS Word file to convert.</param>
    /// <param name="tocTemplate">An MS Word file that is used as a template to build a table of contents.
    /// This file needs to have a mail merge region called "TOC" defined and one mail merge field called "TocEntry".</param>
    /// <param name="dstDir">The output directory where to write HTML files.</param>
    void execute(String srcFileName, String tocTemplate, String dstDir) throws Exception
    {
        mDoc = new Document(srcFileName);
        mTocTemplate = tocTemplate;
        mDstDir = dstDir;

        ArrayList<Paragraph> topicStartParas = selectTopicStarts();
        insertSectionBreaks(topicStartParas);
        ArrayList<Topic> topics = saveHtmlTopics();
        saveTableOfContents(topics);
    }

    /// <summary>
    /// Selects heading paragraphs that must become topic starts.
    /// We can't modify them in this loop, so we need to remember them in an array first.
    /// </summary>
    private ArrayList<Paragraph> selectTopicStarts()
    {
        NodeCollection paras = mDoc.getChildNodes(NodeType.PARAGRAPH, true);
        ArrayList<Paragraph> topicStartParas = new ArrayList<Paragraph>();

        for (Paragraph para : (Iterable<Paragraph>) paras)
        {
            /*StyleIdentifier*/int style = para.getParagraphFormat().getStyleIdentifier();
            if (style == StyleIdentifier.HEADING_1)
                topicStartParas.add(para);
        }

        return topicStartParas;
    }

    /// <summary>
    /// Insert section breaks before the specified paragraphs.
    /// </summary>
    private void insertSectionBreaks(ArrayList<Paragraph> topicStartParas)
    {
        DocumentBuilder builder = new DocumentBuilder(mDoc);
        for (Paragraph para : topicStartParas)
        {
            Section section = para.getParentSection();

            // Insert section break if the paragraph is not at the beginning of a section already.
            if (para != section.getBody().getFirstParagraph())
            {
                builder.moveTo(para.getFirstChild());
                builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

                // This is the paragraph that was inserted at the end of the now old section.
                // We don't really need the extra paragraph, we just needed the section.
                section.getBody().getLastParagraph().remove();
            }
        }
    }

    /// <summary>
    /// Splits the current document into one topic per section and saves each topic
    /// as an HTML file. Returns a collection of Topic objects.
    /// </summary>
    private ArrayList<Topic> saveHtmlTopics() throws Exception
    {
        ArrayList<Topic> topics = new ArrayList<Topic>();
        for (int sectionIdx = 0; sectionIdx < mDoc.getSections().getCount(); sectionIdx++)
        {
            Section section = mDoc.getSections().get(sectionIdx);

            String paraText = section.getBody().getFirstParagraph().getText();

            // Use the text of the heading paragraph to generate the HTML file name.
            String fileName = makeTopicFileName(paraText);
            if ("".equals(fileName))
                fileName = "UNTITLED SECTION " + sectionIdx;

            fileName = mDstDir + fileName + ".html";

            // Use the text of the heading paragraph to generate the title for the TOC.
            String title = makeTopicTitle(paraText);
            if ("".equals(title))
                title = "UNTITLED SECTION " + sectionIdx;

            Topic topic = new Topic(title, fileName);
            topics.add(topic);

            saveHtmlTopic(section, topic);
        }

        return topics;
    }

    /// <summary>
    /// Leaves alphanumeric characters, replaces white space with underscore
    /// And removes all other characters from a string.
    /// </summary>
    private String makeTopicFileName(String paraText)
    {
        StringBuilder b = new StringBuilder();
        for (int i = 0; i < paraText.length(); i++)
        {
        	char c = paraText.charAt(i);
            if (Character.isLetterOrDigit(c))
                b.append(c);
            else if (c == ' ')
                b.append('_');
        }

        return b.toString();
    }

    /// <summary>
    /// Removes the last character (which is a paragraph break character from the given string).
    /// </summary>
    private String makeTopicTitle(String paraText)
    {
        return paraText.substring((0), (0) + (paraText.length() - 1));
    }

    /// <summary>
    /// Saves one section of a document as an HTML file.
    /// Any embedded images are saved as separate files in the same folder as the HTML file.
    /// </summary>
    private void saveHtmlTopic(Section section, Topic topic) throws Exception
    {
        Document dummyDoc = new Document();
        dummyDoc.removeAllChildren();
        dummyDoc.appendChild(dummyDoc.importNode(section, true, ImportFormatMode.KEEP_SOURCE_FORMATTING));

        dummyDoc.getBuiltInDocumentProperties().setTitle(topic.getTitle());

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setPrettyFormat(true);
            saveOptions.setAllowNegativeIndent(true); // This is to allow headings to appear to the left of the main text.
            saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);
        }

        dummyDoc.save(topic.getFileName(), saveOptions);
    }

    /// <summary>
    /// Generates a table of contents for the topics and saves to contents .html.
    /// </summary>
    private void saveTableOfContents(ArrayList<Topic> topics) throws Exception
    {
        Document tocDoc = new Document(mTocTemplate);

        // We use a custom mail merge event handler defined below,
        // and a custom mail merge data source based on collecting the topics we created.
        tocDoc.getMailMerge().setFieldMergingCallback(new HandleTocMergeField());
        tocDoc.getMailMerge().executeWithRegions(new TocMailMergeDataSource(topics));

        tocDoc.save(mDstDir + "contents.html");
    }

    private static class HandleTocMergeField implements IFieldMergingCallback
    {
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e) throws Exception
        {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(e.getDocument());

            // Our custom data source returns topic objects.
            Topic topic = (Topic) e.getFieldValue();

            mBuilder.moveToMergeField(e.getFieldName());
            mBuilder.insertHyperlink(topic.getTitle(), topic.getFileName(), false);

            // Signal to the mail merge engine that it does not need to insert text into the field.
            e.setText("");
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }

        private DocumentBuilder mBuilder;
    }

    private Document mDoc;
    private String mTocTemplate;
    private String mDstDir;
}

class Topic
{
    Topic(String title, String fileName)
    {
        mTitle = title;
        mFileName = fileName;
    }

    String getTitle() { return mTitle; };

    private String mTitle;

    String getFileName() { return mFileName; };

    private String mFileName;
}

class TocMailMergeDataSource implements IMailMergeDataSource
{
    TocMailMergeDataSource(ArrayList<Topic> topics)
    {
        mTopics = topics;
        mIndex = -1;
    }

    @Override
    public String getTableName() throws Exception {
        return "TOC";
    }

    public boolean moveNext()
    {
        if (mIndex < mTopics.size() - 1)
        {
            mIndex++;
            return true;
        }

        return false;
    }

    public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
    {
        if ("TocEntry".equals(fieldName))
        {
            // The template document is supposed to have only one field called "TocEntry".
            fieldValue.set(mTopics.get(mIndex));
            return true;
        }

        fieldValue.set(null);
        return false;
    }

    public IMailMergeDataSource getChildDataSource(String tableName)
    {
        return null;
    }

    private ArrayList<Topic> mTopics;
    private int mIndex;
}
