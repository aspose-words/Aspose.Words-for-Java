/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package loadingandsaving.loadingandsavinghtml.word2help.java;


import java.io.*;
import java.util.ArrayList;
import com.aspose.words.*;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;


/**
 * This is the main class.
 * Loads Word document(s), splits them into topics, saves HTML files and builds content.xml.
 */
public class TopicCollection
{
    /**
     * Ctor.
     *
     * @param htmlTemplatesDir The directory that contains header.html, banner.html and footer.html files.
     *
     * @param fixUrl The url that will be removed from any hyperlinks that start with this url.
     * This allows turning some absolute URLS into relative ones.
     */
    public TopicCollection(String htmlTemplatesDir, String fixUrl) throws Exception
    {
        mTopics = new ArrayList();
        mFixUrl = fixUrl;
        mHtmlHeader = readFile(htmlTemplatesDir + "header.html");
        mHtmlBanner = readFile(htmlTemplatesDir + "banner.html");
        mHtmlFooter = readFile(htmlTemplatesDir + "footer.html");
    }

    /**
     * Processes all DOC files found in the specified directory.
     * Loads and splits each document into separate topics.
     */
    public void addFromDir(String dirName) throws Exception
    {
           FilenameFilter fileFilter = new FilenameFilter() {

             public boolean accept(File dir, String name) {
               return name.endsWith(".doc");
             }
           };

        for (File filename : new File(dirName).listFiles(fileFilter))
            addFromFile(filename.getAbsolutePath());
    }

    /**
     * Processes a specified DOC file. Loads and splits into topics.
     */
    public void addFromFile(String fileName) throws Exception
    {
        Document doc = new Document(fileName);
        insertTopicSections(doc);
        addTopics(doc);
    }

    /**
     * Saves all topics as HTML files.
     */
    public void writeHtml(String outDir) throws Exception
    {
        for (Topic topic : (Iterable<Topic>) mTopics)
        {
            if (!topic.isHeadingOnly())
                topic.writeHtml(mHtmlHeader, mHtmlBanner, mHtmlFooter, outDir);
        }
    }

    /**
     * Saves the content.xml file that describes the tree of topics.
     */
    public void writeContentXml(String outDir) throws Exception
    {
        DocumentBuilderFactory fact = DocumentBuilderFactory.newInstance();
        javax.xml.parsers.DocumentBuilder parser = fact.newDocumentBuilder();
        org.w3c.dom.Document doc = parser.newDocument();

        Element root = doc.createElement("content");
        root.setAttribute("dir", outDir);
        doc.appendChild(root);

        Element currentElement = root;

        for (int i = 0; i < mTopics.size(); i++)
        {
            Topic topic = (Topic)mTopics.get(i);

            int nextTopicIdx = i + 1;
            Topic nextTopic = (nextTopicIdx < mTopics.size()) ? (Topic)mTopics.get(i + 1) : null;

            int nextHeadingLevel = (nextTopic != null) ? nextTopic.getHeadingLevel() : 0;

            if (nextHeadingLevel > topic.getHeadingLevel())
            {
                // Next topic is nested, therefore we have to start a book.
                // We only allow increase level at a time.
                if (nextHeadingLevel != topic.getHeadingLevel() + 1)
                    throw new Exception("Topic is nested for more than one level at a time. Title: " + topic.getTitle());

                currentElement = writeBookStart(currentElement, topic);
            }
            else if (nextHeadingLevel < topic.getHeadingLevel())
            {
                // Next topic is one or more levels higher in the outline.
                // Write out the current topic.
                writeItem(currentElement, topic.getTitle(), topic.getFileName());

                // End one or more nested topics could have ended at this point.
                int levelsToClose = topic.getHeadingLevel() - nextHeadingLevel;
                while (levelsToClose > 0)
                {
                    currentElement = (Element)currentElement.getParentNode();
                    levelsToClose--;
                }
            }
            else
            {
                // A topic at the current level and it has no children.
                writeItem(currentElement, topic.getTitle(), topic.getFileName());
            }
        }

        // Prepare the DOM document for writing
        Source source = new DOMSource(doc);

        // Prepare the output file
        File file = new File(outDir, "content.xml");
        FileOutputStream outputStream = new FileOutputStream(file.getAbsolutePath());
        StreamResult result = new StreamResult(new OutputStreamWriter(outputStream,"UTF-8")); // UTF-8 encoding must be specified in order for the output to have proper indentation.

        // Write the DOM document to disk.
        TransformerFactory tf = TransformerFactory.newInstance();
        tf.setAttribute("indent-number", 2); // Set the indentation for child elements.

        // Export as XML.
        Transformer transformer = tf.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.transform(source, result);
    }

    /**
     * Inserts section breaks that delimit the topics.
     *
     * @param doc The document where to insert the section breaks.
     */
    private static void insertTopicSections(Document doc) throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        NodeCollection paras = doc.getChildNodes(NodeType.PARAGRAPH, true, false);
        ArrayList topicStartParas = new ArrayList();

        for (Paragraph para : (Iterable<Paragraph>) paras)
        {
            int style = para.getParagraphFormat().getStyleIdentifier();
            if ((style >= StyleIdentifier.HEADING_1) && (style <= MAX_TOPIC_HEADING) &&
                (para.hasChildNodes()))
            {
                // Select heading paragraphs that must become topic starts.
                // We can't modify them in this loop, we have to remember them in an array first.
                topicStartParas.add(para);
            }
            else if ((style > MAX_TOPIC_HEADING) && (style <= StyleIdentifier.HEADING_9))
            {
                // Pull up headings. For example: if Heading 1-4 become topics, then I want Headings 5+
                // to become Headings 4+. Maybe I want to pull up even higher?
                para.getParagraphFormat().setStyleIdentifier((/*StyleIdentifier*/int)((int)style - 1));
            }
        }

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
     * Goes through the sections in the document and adds them as topics to the collection.
     */
    private void addTopics(Document doc) throws Exception
    {
        for (Section section : doc.getSections())
        {
            try
            {
                Topic topic = new Topic(section, mFixUrl);
                mTopics.add(topic);
            }
            catch (Exception e)
            {
                // If one topic fails, we continue with others.
                System.out.println(e.getMessage());
            }
        }
    }

    private static Element writeBookStart(Element root, Topic topic) throws Exception
    {
        Element book = root.getOwnerDocument().createElement("book");
        root.appendChild(book);

        book.setAttribute("name", topic.getTitle());

        if (!topic.isHeadingOnly())
            book.setAttribute("href", topic.getFileName());

        return book;
    }

    private static void writeItem(Element root, String name, String href) throws Exception
    {
        Element item = root.getOwnerDocument().createElement("item");
        root.appendChild(item);

        item.setAttribute("name", name);
        item.setAttribute("href", href);
    }

    private static String readFile(String fileName) throws Exception
    {
        FileInputStream reader = null;
        try
        {
            reader = new FileInputStream(fileName);
            byte[] fileBytes = new byte[reader.available()];

            reader.read(fileBytes);

            return new String(fileBytes);
        }

        finally {
            if (reader != null)
                reader.close();
        }
    }

    private final ArrayList mTopics;
    private final String mFixUrl;
    private final String mHtmlHeader;
    private final String mHtmlBanner;
    private final String mHtmlFooter;

    /**
     * Specifies the maximum Heading X number.
     * All of the headings above or equal to this will be put into their own topics.
     */
    private static final int MAX_TOPIC_HEADING = StyleIdentifier.HEADING_4;
}