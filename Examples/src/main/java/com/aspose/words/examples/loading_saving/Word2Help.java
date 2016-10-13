package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * This project converts documentation stored inside a DOC format to a series of HTML documents. This output is in
 * a form that can then be easily compiled together into a single compiled help file (CHM) by using
 * the Microsoft HTML Help Workshop application.
 */
public class Word2Help
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Word2Help.class);

        // Specifies the destination directory where the HTML files are output.
        File outPath = new File(dataDir, "Out");

        // Remove any existing output and recreate the Out folder.
        if(outPath.exists())
        {
            for(File file : outPath.listFiles())
            {
                file.delete();
            }
        }

        outPath.mkdirs();
        String outDir = outPath.getAbsolutePath();

        // Specifies the part of the URLs to remove. If there are any hyperlinks that start
        // with the above URL, this URL is removed. This allows the document designer to include
        // links to the HTML API and they will be "corrected" so they work both in the online
        // HTML and also in the compiled CHM.
        String fixUrl = "";

        // *** LICENSING ***
        // An Aspose.Words license is required to use this project fully.
        // Without a license Aspose.Words will work in evaluation mode and truncate documents
        // and output watermarks.
        //
        // You can download a free 30-day trial license from the Aspose site. The easiest way is to set the license is to
        // include the license in the executing directory and uncomment the following code.
        //
        // Aspose.Words.License license = new Aspose.Words.License();
        // license.setLicense("Aspose.Words.lic");
        System.out.println(MessageFormat.format("Extracting topics from {0}.", dataDir));

        TopicCollection topics = new TopicCollection(dataDir, fixUrl);
        topics.addFromDir(dataDir);
        topics.writeHtml(outDir);
        topics.writeContentXml(outDir);

        System.out.println("Conversion completed successfully.");
    }
}

/**
 * This "facade" class makes it easier to work with a hyperlink field in a Word document.
 *
 * A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words
 * consists of several nodes and it might be difficult to work with all those nodes directly.
 * This is a simple implementation and will work only if the hyperlink code and name
 * each consist of one Run only.
 *
 * [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
 *
 * The field code contains a string in one of these formats:
 * HYPERLINK "url"
 * HYPERLINK \l "bookmark name"
 *
 * The field result contains text that is displayed to the user.
 */
class Hyperlink
{
    public Hyperlink(FieldStart fieldStart) throws Exception
    {
        if (fieldStart == null)
            throw new IllegalArgumentException("fieldStart");
        if (fieldStart.getFieldType() != FieldType.FIELD_HYPERLINK)
            throw new IllegalArgumentException("Field start type must be FieldHyperlink.");

        mFieldStart = fieldStart;

        // Find field separator node.
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null)
            throw new Exception("Cannot find field separator.");

        // Find field end node. Normally field end will always be found, but in the example document
        // there happens to be a paragraph break included in the hyperlink and this puts the field end
        // in the next paragraph. It will be much more complicated to handle fields which span several
        // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);

        // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
        String fieldCode = getTextSameParent(mFieldStart.getNextSibling(), mFieldSeparator);

        Matcher match = G_REGEX.matcher(fieldCode.trim());

        if(match.find())
        {
            mIsLocal = match.group(1) != null;
            mTarget = match.group(2);
        }
    }

    /*
     * Gets or sets the display name of the hyperlink.
     */
    public String getName() throws Exception
    {
        return getTextSameParent(mFieldSeparator, mFieldEnd);
    }

    public void setName(String value) throws Exception
    {
        // Hyperlink display name is stored in the field result which is a Run
        // node between field separator and field end.
        Run fieldResult = (Run)mFieldSeparator.getNextSibling();
        fieldResult.setText(value);

        // But sometimes the field result can consist of more than one run, delete these runs.
        removeSameParent(fieldResult.getNextSibling(), mFieldEnd);
    }

    /*
     * Gets or sets the target url or bookmark name of the hyperlink.
     */
    public String getTarget() throws Exception
    {
        return mTarget;
    }

    public void setTarget(String value) throws Exception
    {
        mTarget = value;
        updateFieldCode();
    }

    /*
     * True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
     */
    public boolean isLocal() throws Exception
    {
        return mIsLocal;
    }

    public void setLocal(boolean value) throws Exception
    {
        mIsLocal = value;
        updateFieldCode();
    }

    /**
     * Updates the field code.
     */
    private void updateFieldCode() throws Exception
    {
        // Field code is stored in a Run node between field start and field separator.
        Run fieldCode = (Run)mFieldStart.getNextSibling();
        fieldCode.setText(java.text.MessageFormat.format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget));

        // But sometimes the field code can consist of more than one run, delete these runs.
        removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
    }

    /**
     * Goes through siblings starting from the start node until it finds a node of the specified type or null.
     */
    private static Node findNextSibling(Node start, int nodeType) throws Exception
    {
        for (Node node = start; node != null; node = node.getNextSibling())
        {
            if (node.getNodeType() == nodeType)
                return node;
        }
        return null;
    }

    /*
     * Retrieves text from start up to but not including the end node.
     */
    private static String getTextSameParent(Node start, Node end) throws Exception
    {
        if ((end != null) && (start.getParentNode() != end.getParentNode()))
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        StringBuilder builder = new StringBuilder();
        for (Node child = start; child != end; child = child.getNextSibling())
            builder.append(child.getText());
        return builder.toString();
    }

    /*
     * Removes nodes from start up to but not including the end node.
     * Start and end are assumed to have the same parent.
     */
    private static void removeSameParent(Node start, Node end) throws Exception
    {
        if ((end != null) && (start.getParentNode() != end.getParentNode()))
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        Node curChild = start;
        while (curChild != end)
        {
            Node nextChild = curChild.getNextSibling();
            curChild.remove();
            curChild = nextChild;
        }
    }

    private final Node mFieldStart;
    private final Node mFieldSeparator;
    private final Node mFieldEnd;
    private String mTarget;
    private boolean mIsLocal;

    private static final Pattern G_REGEX = Pattern.compile(
            "\\S+" +            // One or more non spaces HYPERLINK or other word in other languages
                    "\\s+" +            // One or more spaces
                    "(?:\"\"\\s+)?" +   // Non capturing optional "" and one or more spaces, found in one of the customers files.
                    "(\\\\l\\s+)?" +    // Optional \l flag followed by one or more spaces
                    "\"" +              // One apostrophe
                    "([^\"]+)" +        // One or more chars except apostrophe (hyperlink target)
                    "\""                // One closing apostrophe
    );
}

/**
 * Central storage for regular expressions used in the project.
 */
class RegularExpressions
{
    // This class is static. No instance creation is allowed.
    private RegularExpressions() throws Exception {}

    /**
     * Regular expression specifying html title (framing tags excluded).
     */
    public static Pattern getHtmlTitle() throws Exception
    {
        if (gHtmlTitle == null)
        {
            gHtmlTitle = Pattern.compile(HTML_TITLE_PATTERN,
                    Pattern.CASE_INSENSITIVE);
        }
        return gHtmlTitle;
    }

    /**
     * Regular expression specifying html head.
     */
    public static Pattern getHtmlHead() throws Exception
    {
        if (gHtmlHead == null)
        {
            gHtmlHead = Pattern.compile(HTML_HEAD_PATTERN,
                    Pattern.CASE_INSENSITIVE);
        }
        return gHtmlHead;
    }

    /**
     * Regular expression specifying space right after div keyword in the first div declaration of html body.
     */
    public static Pattern getHtmlBodyDivStart() throws Exception
    {
        if (gHtmlBodyDivStart == null)
        {
            gHtmlBodyDivStart = Pattern.compile(HTML_BODY_DIV_START_PATTERN,
                    Pattern.CASE_INSENSITIVE);
        }
        return gHtmlBodyDivStart;
    }

    private static final String HTML_TITLE_PATTERN = "(?<=\\<title\\>).*?(?=\\</title\\>)";
    private static Pattern gHtmlTitle;

    private static final String HTML_HEAD_PATTERN = "\\<head\\>.*?\\</head\\>";
    private static Pattern gHtmlHead;

    private static final String HTML_BODY_DIV_START_PATTERN = "(?<=\\<body\\>\\s{0,200}\\<div)\\s";
    private static Pattern gHtmlBodyDivStart;
}

/**
 * Represents a single topic that will be written as an HTML file.
 */
class TopicWord2Help
{
    /**
     * Creates a topic.
     */
    public TopicWord2Help(Section section, String fixUrl) throws Exception
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

/**
 * This is the main class.
 * Loads Word document(s), splits them into topics, saves HTML files and builds content.xml.
 */
class TopicCollection
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
        for (TopicWord2Help topic : (Iterable<TopicWord2Help>) mTopics)
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
            TopicWord2Help topic = (TopicWord2Help)mTopics.get(i);

            int nextTopicIdx = i + 1;
            TopicWord2Help nextTopic = (nextTopicIdx < mTopics.size()) ? (TopicWord2Help)mTopics.get(i + 1) : null;

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
                para.getParagraphFormat().setStyleIdentifier(style - 1);
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
                TopicWord2Help topic = new TopicWord2Help(section, mFixUrl);
                mTopics.add(topic);
            }
            catch (Exception e)
            {
                // If one topic fails, we continue with others.
                System.out.println(e.getMessage());
            }
        }
    }

    private static Element writeBookStart(Element root, TopicWord2Help topic) throws Exception
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

