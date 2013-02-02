//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// 3/1/08 by Roman Korchagin
package ImportFootnotesFromHtml;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.File;
import java.net.URI;

import com.aspose.words.Document;
import com.aspose.words.FootnoteType;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.Node;
import com.aspose.words.Run;
import com.aspose.words.Paragraph;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Footnote;


/**
 * This is a sample code for http://www.aspose.com/Community/Forums/thread/107477.aspx
 *
 * The scenario is as follows:
 *
 * 1. The customer has a DOC file with footnotes.
 *
 * 2. The customer uses Aspose.Words to convert DOC to HTML. Aspose.Words converts
 * footnotes and endnotes into hyperlinks. There are two hyperlinks per footnote actually.
 * One link is "forward" from the main text to the text of the footnote.
 * Another is "backward" from the text of the footnote to the main text.
 *
 * 3. The customer uses Aspose.Words to convert HTML back to DOC.
 * In the current version of Aspose.Words the hyperlinks do not become footnotes,
 * they just stay as hyperlink fields in the document. The customer wants
 * original footnotes to become footnotes during DOC->HTML->DOC roundtrip.
 *
 * This code is a workaround that detects hyperlinks related to footnotes and converts
 * them into proper footnotes. At some point in the future, this code will not be needed
 * when Aspose.Words will guarantee footnotes roundtripping.
 *
 * This code demonstrates some useful techniques, such as enumerating over nodes,
 * getting field code, removing fields etc.
 */
class Program
{
    public static void main(String[] args) throws Exception
    {
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = new File(exeDir.resolve("../../Data")) + File.separator;

        // Load DOC with footnotes into a document object.
        Document srcDoc = new Document(dataDir + "FootnoteSample.doc");

        // Save to HTML file. Footnotes get converted to hyperlinks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setPrettyFormat(true);
        String htmlFile = dataDir + "FootnoteSample Out.html";
        srcDoc.save(htmlFile, saveOptions);

        // Load HTML back into a document object.
        // In the current version of Aspose.Words hyperlinks do not become footnotes again,
        // they become just regular hyperlinks.
        Document dstDoc = new Document(htmlFile);

        // You can open this document in MS Word and see there are no footnotes, just hyperlinks.
        dstDoc.save(dataDir + "FootnoteSample Out1.doc");

        // This is the workaround method I'm suggesting. It will recognize hyperlinks that
        // should become footnotes and convert them into footnotes.
        convertHyperlinksToFootnotes(dstDoc);

        // You can open this document in MS Word and see the footnotes are as expected.
        dstDoc.save(dataDir + "FootnoteSample Out2.doc");
    }

    /*
     * A "workaround" method that you can use after DOC->HTML->DOC conversion of a document
     * with footnotes. Will make sure that original DOC footnotes will still be footnotes in
     * the final DOC file.
     */
    static void convertHyperlinksToFootnotes(Document doc) throws Exception
    {
        // When processing HYPERLINK fields we will remove them (convert to footnotes).
        // Since it is not a good thing to delete nodes while iterating over a collection,
        // we will collect the nodes during the first pass and delete them during the second.
        //
        // These collections contain HYPERLINK field starts of footnotes and endnotes in the main document.
        HashMap ftnFieldStarts = new HashMap();
        HashMap ednFieldStarts = new HashMap();
        // These collections contain HYPERLINK field starts of footnotes and endnotes themselves.
        HashMap ftnRefFieldStarts = new HashMap();
        HashMap ednRefFieldStarts = new HashMap();

        // Collect all the nodes into arrays before we start deleting them.
        collectFieldStarts(doc, ftnFieldStarts, ednFieldStarts, ftnRefFieldStarts, ednRefFieldStarts);

        // Remove the HR shapes that separate footnotes and endnotes from the main text.
        removeHorizontalLine(ftnRefFieldStarts);
        removeHorizontalLine(ednRefFieldStarts);

        // Convert the HYPERLINK fields into proper footnotes and endnotes.
        convertFieldsToNotes(ftnFieldStarts, ftnRefFieldStarts, FootnoteType.FOOTNOTE);
        convertFieldsToNotes(ednFieldStarts, ednRefFieldStarts, FootnoteType.ENDNOTE);
    }

    /**
     * Collects field start nodes of HYPERLINK fields related to footnotes and endnotes.
     *
     * @param doc The document to process.
     * @param ftnFieldStarts Starts of HYPERLINK fields that represent footnotes will be returned here.
     * @param ednFieldStarts Start of HYPERLINK fields that represent endnotes will be returned here.
     * @param ftnRefFieldStarts Starts of HYPERLINK fields that are back-links to footnotes will be returned here.
     * @param ednRefFieldStarts Starts of HYPERLINK fields that are back-links to endnotes will be returned here.
     */
    private static void collectFieldStarts(
        Document doc,
        HashMap ftnFieldStarts,
        HashMap ednFieldStarts,
        HashMap ftnRefFieldStarts,
        HashMap ednRefFieldStarts) throws Exception
    {
        // This regex parses the "command" which we use to determine the footnote/endnote type
        // and the id.
        Pattern pattern = Pattern.compile("HYPERLINK \\\\l \"(_ftn|_edn|_ftnref|_ednref)([0-9]+)\"");

        // We need to process all HYPERLINK fields. Therefore select all field starts.
        NodeCollection fieldStarts = doc.getChildNodes(NodeType.FIELD_START, true);
        for (FieldStart fieldStart : (Iterable<FieldStart>) fieldStarts)
        {
            if (fieldStart.getFieldType() == FieldType.FIELD_HYPERLINK)
            {
                // The field is a hyperlink, lets analyze the field code.
                String fieldCode = getFieldCode(fieldStart);

                Matcher matcher = pattern.matcher(fieldCode);

                if(matcher.find())
                {
                    String cmd = matcher.group(1);
                    String id = matcher.group(2);

                    if(cmd.equals("_ftn"))
                    {
                        // Field is HYPERLINK \l "_ftn1". It is a footnote in the main document.
                        ftnFieldStarts.put(Integer.parseInt(id), fieldStart);
                    }
                    else if(cmd.equals("_edn"))
                    {
                        ednFieldStarts.put(Integer.parseInt(id), fieldStart);
                    }
                    else if(cmd.equals("_ftnref"))
                    {
                        // Field is HYPERLINK \l "_ftnref1". It is a back-link to the footnote in
                        // the main document. The parent paragraph contains the text of the footnote.
                        ftnRefFieldStarts.put(Integer.parseInt(id), fieldStart);
                    }
                    else if(cmd.equals("_ednref"))
                    {
                        ednRefFieldStarts.put(Integer.parseInt(id), fieldStart);
                    }
                }
            }
        }
    }

    /**
     * A simplistic method to get the field code as a string.
     * Goes trough all Run nodes after the field start and concatenates their text.
     */
    private static String getFieldCode(FieldStart fieldStart) throws Exception
    {
        StringBuilder fieldCode = new StringBuilder();
        Node curNode = fieldStart.getNextSibling();
        while (curNode instanceof Run)
        {
            fieldCode.append(curNode.getText());
            curNode = curNode.getNextSibling();
        }
        return fieldCode.toString();
    }

    /**
     * Performs the actual conversion of HYPERLINK fields into footnotes/endnote.
     *
     * @param noteFieldStarts The starts of hyperlink fields in the main document.
     * @param refNoteFieldStarts The starts of back-link hyperlink fields in the footnotes.
     * @param noteType Specifies whether we are processing footnotes or endnotes.
     */
    private static void convertFieldsToNotes(
        HashMap noteFieldStarts,
        HashMap refNoteFieldStarts,
        int noteType) throws Exception
    {
        for (Map.Entry entry : (Iterable<Map.Entry>) noteFieldStarts.entrySet())
        {
            // Footnote/endnote id is stored in the key.
            int id = (Integer)entry.getKey();
            FieldStart noteFieldStart = (FieldStart)entry.getValue();
            // Using the id we can retrieve the field start of the back-link field.
            FieldStart refNoteFieldStart = (FieldStart)refNoteFieldStarts.get(id);

            convertFieldToNote(noteFieldStart, refNoteFieldStart, noteType);
        }
    }

    /**
     * Performs the actual task of converting one HYPERLINK into footnote or endnote.
     *
     * @param noteFieldStart The start of the hyperlink field in the main document.
     * @param refNoteFieldStart The start of the back-link hyperlink field in the footnote.
     * @param noteType Specifies whether we are processing footnotes or endnotes.
     */
    private static void convertFieldToNote(
        FieldStart noteFieldStart,
        FieldStart refNoteFieldStart,
        int noteType) throws Exception
    {
        // This is the paragraph that contains the text of the footnote.
        Paragraph oldNotePara = refNoteFieldStart.getParentParagraph();

        // Delete the hyperlink field from the text of the footnote because we don't need it anymore.
        deleteField(refNoteFieldStart);

        // Use document builder to move to the place in the main document where the footnote
        // should be and insert a proper footnote.
        DocumentBuilder builder = new DocumentBuilder((Document)noteFieldStart.getDocument());
        builder.moveTo(noteFieldStart);
        Footnote note = builder.insertFootnote(noteType, "");

        // Move all content from the old footnote paragraphs into the new.
        Paragraph newNotePara = note.getFirstParagraph();
        Node curNode = oldNotePara.getFirstChild();
        while (curNode != null)
        {
            Node nextNode = curNode.getNextSibling();
            newNotePara.appendChild(curNode);
            curNode = nextNode;
        }

        // Delete the old paragraph that represented the footnote.
        oldNotePara.remove();

        // Remove the hyperlink field from the main text to the footnote.
        deleteField(noteFieldStart);
    }

    /**
     * A simplistic method to delete all nodes of a field given a field start node.
     */
    private static void deleteField(FieldStart fieldStart) throws Exception
    {
        Node curNode = fieldStart;
        while (curNode.getNodeType() != NodeType.FIELD_END)
        {
            Node nextNode = curNode.getNextSibling();
            curNode.remove();
            curNode = nextNode;
        }
        curNode.remove();
    }

    /**
     * There is an HR (horizontal rule) shape in a separate paragraph just before
     * the first footnote and first endnote in a document imported from HTML.
     * This method deletes the paragraph and the HR shape.
     */
    private static void removeHorizontalLine(HashMap noteRefFieldStarts) throws Exception
    {
        // Footnote and endnote ids start from 1. Therefore we can get the first note.
        FieldStart noteFieldStart = (FieldStart)noteRefFieldStarts.get(1);
        // This is the paragraph that contains the first footnote.
        Paragraph notePara = noteFieldStart.getParentParagraph();
        // This is the previous paragraph that contains the HR shape. Delete the paragraph.
        Paragraph hrPara = (Paragraph)notePara.getPreviousSibling();
        hrPara.remove();
    }
}

