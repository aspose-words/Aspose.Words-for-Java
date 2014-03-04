/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.word2help.java;

import com.aspose.words.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.MessageFormat;


/**
 * Represents a single topic that will be written as an HTML file.
 */
public class Topic
{
    /**
     * Creates a topic.
     */
    public Topic(Section section, String fixUrl) throws Exception
    {
        mTopicDoc = new Document();
        mTopicDoc.appendChild(mTopicDoc.importNode(section, true, ImportFormatMode.KEEP_SOURCE_FORMATTING));
        mTopicDoc.getFirstSection().remove();

        Paragraph headingPara = (Paragraph)mTopicDoc.getFirstSection().getBody().getFirstChild();
        if (headingPara == null)
            throwTopicException("The section does not start with a paragraph.", section);

        mHeadingLevel = headingPara.getParagraphFormat().getStyleIdentifier() - StyleIdentifier.HEADING_1;
        if ((mHeadingLevel < 0) || (mHeadingLevel > 8))
            throwTopicException("This topic does not start with a heading style paragraph.", section);

        mTitle = headingPara.getText().trim();
        if ("".equals(mTitle))
            throwTopicException("This topic heading does not have text.", section);

        // We actually remove the heading paragraph because <h1> will be output in the banner.
        headingPara.remove();

        mTopicDoc.getBuiltInDocumentProperties().setTitle(mTitle);

        fixHyperlinks(section.getDocument(), fixUrl);
    }

    private static void throwTopicException(String message, Section section) throws Exception
    {
        throw new Exception(message + " Section text: " + section.getBody().toString(SaveFormat.TEXT).substring(0, 50));
    }

    private void fixHyperlinks(DocumentBase originalDoc, String fixUrl) throws Exception
    {
        if (fixUrl.endsWith("/"))
            fixUrl = fixUrl.substring(0, fixUrl.length() - 1);

        NodeCollection fieldStarts = mTopicDoc.getChildNodes(NodeType.FIELD_START, true);
        for (FieldStart fieldStart : (Iterable<FieldStart>) fieldStarts)
        {
            if (fieldStart.getFieldType() != FieldType.FIELD_HYPERLINK)
                continue;

            Hyperlink hyperlink = new Hyperlink(fieldStart);
            if (hyperlink.isLocal())
            {
                // We use "Hyperlink to a place in this document" feature of Microsoft Word
                // to create local hyperlinks between topics within the same doc file.
                // It causes MS Word to auto generate the bookmark name.
                String bmkName = hyperlink.getTarget();

                // But we have to follow the bookmark to get the text of the topic heading paragraph
                // in order to be able to build the proper filename of the topic file.
                Bookmark bmk = originalDoc.getRange().getBookmarks().get(bmkName);

               // String test1 = MessageFormat.format("Found a link to a bookmark, but cannot locate the bookmark. Name:{0}.", bmkName);

                if (bmk == null)
                    throw new Exception(MessageFormat.format("Found a link to a bookmark, but cannot locate the bookmark. Name:{0}.", bmkName));

                Paragraph para = (Paragraph)bmk.getBookmarkStart().getParentNode();
                String topicName = para.getText().trim();

                hyperlink.setTarget(headingToFileName(topicName) + ".html");
                hyperlink.setLocal(false);
            }
            else
            {
                // We "fix" URL like this:
                // http://www.aspose.com/Products/Aspose.Words/Api/Aspose.Words.Body.html
                // by changing them into this:
                // Aspose.Words.Body.html
                if (hyperlink.getTarget().startsWith(fixUrl) &&
                        (hyperlink.getTarget().length() > (fixUrl.length() + 1)))
                {
                    hyperlink.setTarget(hyperlink.getTarget().substring(fixUrl.length() + 1));
                }
            }
        }
    }

    public void writeHtml(String htmlHeader, String htmlBanner, String htmlFooter, String outDir) throws Exception
    {
        String fileName = new File(outDir,  getFileName()).getAbsolutePath();

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setPrettyFormat(true);
        // This is to allow headings to appear to the left of main text.
        saveOptions.setAllowNegativeLeftIndent(true);
        // Disable headers and footers.
        saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);

        // Export the document to HTML.
        mTopicDoc.save(fileName, saveOptions);

        // We need to modify the HTML string, read HTML back.
        String html;
        FileInputStream reader = null;

        try{
            reader = new FileInputStream(fileName);
            byte[] fileBytes = new byte[reader.available()];
            reader.read(fileBytes);
            html = new String(fileBytes);
        }

        finally { if (reader != null) reader.close(); }

        // Builds the HTML <head> element.
        String header = htmlHeader.replaceFirst(RegularExpressions.getHtmlTitle().pattern(), mTitle);

        // Applies the new <head> element instead of the original one.
        html = html.replaceFirst(RegularExpressions.getHtmlHead().pattern(), header);
        html = html.replaceFirst(RegularExpressions.getHtmlBodyDivStart().pattern(), " id=\"nstext\"");

        String banner = htmlBanner.replace("###TOPIC_NAME###", mTitle);

        // Add the standard banner.
        html = html.replace("<body>", "<body>" + banner);

        // Add the standard footer.
        html = html.replace("</body>", htmlFooter + "</body>");

        FileOutputStream writer = null;

        try{
            writer = new FileOutputStream(fileName);
            writer.write(html.getBytes());
        }

        finally { if (writer != null) writer.close(); }
    }

    /**
     * Removes various characters from the header to form a file name that does not require escaping.
     */
    private static String headingToFileName(String heading) throws Exception
    {
        StringBuilder b = new StringBuilder();
        for (int i = 0; i < heading.length(); i++)
        {
            char c = heading.charAt(i);
            if (Character.isLetterOrDigit(c))
                b.append(c);
        }

        return b.toString();
    }

    public Document getDocument() throws Exception { return mTopicDoc; }

    /**
     * Gets the name of the topic html file without path.
     */
    public String getFileName() throws Exception { return headingToFileName(mTitle) + ".html"; }

    public String getTitle() throws Exception { return mTitle; }

    public int getHeadingLevel() throws Exception { return mHeadingLevel; }

    /**
     * Returns true if the topic has no text (the heading paragraph has already been removed from the topic).
     */
    public boolean isHeadingOnly() throws Exception
    {
        Body body = mTopicDoc.getFirstSection().getBody();
        return (body.getFirstParagraph() == null);
    }

    private final Document mTopicDoc;
    private final String mTitle;
    private final int mHeadingLevel;
}