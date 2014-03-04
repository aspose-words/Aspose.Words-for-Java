/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.splitintohtmlpages.java;

import com.aspose.words.*;

import java.io.File;
import java.util.ArrayList;


/**
 *This class takes a Microsoft Word document, splits it into topics at paragraphs formatted
 * with the Heading 1 style and saves every topic as an HTML file.
 *
 * Also generates contents.html file that provides links to all saved topics.
 */
class Worker
{
    /**
     * Performs the Word to HTML conversion.
     *
     * @param srcFileName The MS Word file to convert.
     * @param tocTemplate An MS Word file that is used as a template to build
     * a table of contents. This file needs to have a mail merge region called "TOC" defined
     * and one mail merge field called "TocEntry".
     * @param dstDir The output directory where to write HTML files. Must exist.
     */
    void execute(String srcFileName, String tocTemplate, String dstDir) throws Exception
    {
        mDoc = new Document(srcFileName);
        mTocTemplate = tocTemplate;
        mDstDir = dstDir;

        ArrayList topicStartParas = selectTopicStarts();
        insertSectionBreaks(topicStartParas);
        ArrayList topics = saveHtmlTopics();
        saveTableOfContents(topics);
    }

    /**
     * Selects heading paragraphs that must become topic starts.
     * We can't modify them in this loop, we have to remember them in an array first.
     */
    private ArrayList selectTopicStarts() throws Exception
    {
        NodeCollection paras = mDoc.getChildNodes(NodeType.PARAGRAPH, true, false);
        ArrayList topicStartParas = new ArrayList();

        for (Paragraph para : (Iterable<Paragraph>) paras)
        {
            int style = para.getParagraphFormat().getStyleIdentifier();
            if (style == StyleIdentifier.HEADING_1)
                topicStartParas.add(para);
        }

        return topicStartParas;
    }

    /**
     * Inserts section breaks before the specified paragraphs.
     */
    private void insertSectionBreaks(ArrayList topicStartParas) throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder(mDoc);
        for (Paragraph para : (Iterable<Paragraph>) topicStartParas)
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

    /**
     * Splits the current document into one topic per section and saves each topic
     * as an HTML file. Returns a collection of Topic objects.
     */
    private ArrayList saveHtmlTopics() throws Exception
    {
        ArrayList topics = new ArrayList();
        for (int sectionIdx = 0; sectionIdx < mDoc.getSections().getCount(); sectionIdx++)
        {
            Section section = mDoc.getSections().get(sectionIdx);

            String paraText = section.getBody().getFirstParagraph().getText();

            // The text of the heading paragaph is used to generate the HTML file name.
            String fileName = makeTopicFileName(paraText);
            if ("".equals(fileName))
                fileName = "UNTITLED SECTION " + sectionIdx;

            fileName = new File(mDstDir, fileName + ".html").getPath();

            // The text of the heading paragraph is also used to generate the title for the TOC.
            String title = makeTopicTitle(paraText);
            if ("".equals(title))
                title = "UNTITLED SECTION " + sectionIdx;

            Topic topic = new Topic(title, fileName);
            topics.add(topic);

            saveHtmlTopic(section, topic);
        }

        return topics;
    }

    /**
     * Leaves alphanumeric characters, replaces white space with underscore
     * and removes all other characters from a string.
     */
    private static String makeTopicFileName(String paraText) throws Exception
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

    /**
     * Removes the last character (which is a paragraph break character from the given string).
     */
    private static String makeTopicTitle(String paraText) throws Exception
    {
        return paraText.substring((0), (0) + (paraText.length() - 1));
    }

    /**
     * Saves one section of a document as an HTML file.
     * Any embedded images are saved as separate files in the same folder as the HTML file.
     */
    private static void saveHtmlTopic(Section section, Topic topic) throws Exception
    {
        Document dummyDoc = new Document();
        dummyDoc.removeAllChildren();
        dummyDoc.appendChild(dummyDoc.importNode(section, true, ImportFormatMode.KEEP_SOURCE_FORMATTING));

        dummyDoc.getBuiltInDocumentProperties().setTitle(topic.getTitle());

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setPrettyFormat(true);
        // This is to allow headings to appear to the left of main text.
        saveOptions.setAllowNegativeLeftIndent(true);
        saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);

        dummyDoc.save(topic.getFileName(), saveOptions);
    }

    /**
     * Generates a table of contents for the topics and saves to contents.html.
     */
    private void saveTableOfContents(ArrayList topics) throws Exception
    {
        Document tocDoc = new Document(mTocTemplate);

        // We use a custom mail merge even handler defined below.
        tocDoc.getMailMerge().setFieldMergingCallback(new HandleTocMergeField());
        // We use a custom mail merge data source based on the collection of the topics we created.
        tocDoc.getMailMerge().executeWithRegions(new TocMailMergeDataSource(topics));

        tocDoc.save(new File(mDstDir, "contents.html").getPath());
    }

    private class HandleTocMergeField implements IFieldMergingCallback
    {
        public void fieldMerging(FieldMergingArgs e) throws Exception
        {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(e.getDocument());

            // Our custom data source returns topic objects.
            Topic topic = (Topic)e.getFieldValue();

            // We use the document builder to move to the current merge field and insert a hyperlink.
            mBuilder.moveToMergeField(e.getFieldName());
            mBuilder.insertHyperlink(topic.getTitle(), topic.getFileName(), false);

            // Signal to the mail merge engine that it does not need to insert text into the field
            // as we've done it already.
            e.setText("");
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception
        {
            // Do nothing.
        }

        private DocumentBuilder mBuilder;
    }

    private Document mDoc;
    private String mTocTemplate;
    private String mDstDir;
}