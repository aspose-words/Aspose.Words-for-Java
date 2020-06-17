// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.ProtectionType;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.words.Section;
import com.aspose.words.SectionStart;
import com.aspose.words.PaperSize;
import com.aspose.words.Body;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Run;
import java.awt.Color;
import com.aspose.words.ControlChar;
import com.aspose.words.NodeType;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Node;
import com.aspose.ms.System.msConsole;
import com.aspose.words.HeaderFooter;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.ms.System.Globalization.msCultureInfo;


@Test
public class ExSection extends ApiExampleBase
{
    @Test
    public void protect() throws Exception
    {
        //ExStart
        //ExFor:Document.Protect(ProtectionType)
        //ExFor:ProtectionType
        //ExFor:Section.ProtectedForForms
        //ExSummary:Shows how to protect a section so only editing in form fields is possible.
        Document doc = new Document();

        // Insert two sections with some text
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Section 1. Unprotected.");
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.writeln("Section 2. Protected.");

        // Section protection only works when document protection is turned and only editing in form fields is allowed
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS);

        // By default, all sections are protected, but we can selectively turn protection off
        doc.getSections().get(0).setProtectedForForms(false);

        doc.save(getArtifactsDir() + "Section.Protect.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Section.Protect.docx");

        Assert.assertFalse(doc.getSections().get(0).getProtectedForForms());
        Assert.assertTrue(doc.getSections().get(1).getProtectedForForms());
    }

    @Test
    public void addRemove() throws Exception
    {
        //ExStart
        //ExFor:Document.Sections
        //ExFor:Section.Clone
        //ExFor:SectionCollection
        //ExFor:NodeCollection.RemoveAt(Int32)
        //ExSummary:Shows how to add/remove sections in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");

        // This shows what is in the document originally. The document has two sections
        Assert.assertEquals("Section 1\fSection 2", msString.trim(doc.getText()));

        // Delete the first section from the document
        doc.getSections().removeAt(0);

        // Duplicate the last section and append the copy to the end of the document
        int lastSectionIdx = doc.getSections().getCount() - 1;
        Section newSection = doc.getSections().get(lastSectionIdx).deepClone();
        doc.getSections().add(newSection);

        // Check what the document contains after we changed it
        Assert.assertEquals("Section 2\fSection 2", msString.trim(doc.getText()));
        //ExEnd
    }

    @Test
    public void createFromScratch() throws Exception
    {
        //ExStart
        //ExFor:Node.GetText
        //ExFor:CompositeNode.RemoveAllChildren
        //ExFor:CompositeNode.AppendChild
        //ExFor:Section
        //ExFor:Section.#ctor
        //ExFor:Section.PageSetup
        //ExFor:PageSetup.SectionStart
        //ExFor:PageSetup.PaperSize
        //ExFor:SectionStart
        //ExFor:PaperSize
        //ExFor:Body
        //ExFor:Body.#ctor
        //ExFor:Paragraph
        //ExFor:Paragraph.#ctor
        //ExFor:Paragraph.ParagraphFormat
        //ExFor:ParagraphFormat
        //ExFor:ParagraphFormat.StyleName
        //ExFor:ParagraphFormat.Alignment
        //ExFor:ParagraphAlignment
        //ExFor:Run
        //ExFor:Run.#ctor(DocumentBase)
        //ExFor:Run.Text
        //ExFor:Inline.Font
        //ExSummary:Shows how to construct an Aspose Words document node by node.
        Document doc = new Document();

        // A newly created blank document still comes one section, one body and one paragraph
        // Calling this method will remove all those nodes to completely empty the document
        doc.removeAllChildren();

        // This document now has no composite nodes that content can be added to
        // If we wish to edit it, we will need to repopulate its node collection,
        // which we will start to do with by creating a new Section node
        Section section = new Section(doc);

        // Append the section to the document
        doc.appendChild(section);

        // Lets set some properties for the section
        section.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        section.getPageSetup().setPaperSize(PaperSize.LETTER);

        // The section that we created is empty, lets populate it. The section needs at least the Body node
        Body body = new Body(doc);
        section.appendChild(body);

        // The body needs to have at least one paragraph
        // Note that the paragraph has not yet been added to the document, 
        // but we have to specify the parent document
        // The parent document is needed so the paragraph can correctly work
        // with styles and other document-wide information
        Paragraph para = new Paragraph(doc);
        body.appendChild(para);

        // We can set some formatting for the paragraph
        para.getParagraphFormat().setStyleName("Heading 1");
        para.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // So far we have one empty paragraph in the document
        // The document is valid and can be saved, but lets add some text before saving
        // Create a new run of text and add it to our paragraph
        Run run = new Run(doc);
        run.setText("Hello World!");
        run.getFont().setColor(Color.RED);
        para.appendChild(run);

        Assert.assertEquals("Hello World!" + ControlChar.SECTION_BREAK_CHAR, doc.getText());

        doc.save(getArtifactsDir() + "Section.CreateFromScratch.docx");
        //ExEnd
    }

    @Test
    public void ensureSectionMinimum() throws Exception
    {
        //ExStart
        //ExFor:NodeCollection.Add
        //ExFor:Section.EnsureMinimum
        //ExFor:SectionCollection.Item(Int32)
        //ExSummary:Shows how to prepare a new section node for editing.
        Document doc = new Document();
        
        // A blank document comes with a section, which has a body, which in turn has a paragraph,
        // so we can edit the document by adding children to the paragraph like shapes or runs of text
        Assert.assertEquals(2, doc.getSections().get(0).getChildNodes(NodeType.ANY, true).getCount());

        // If we add a new section like this, it will not have a body or a paragraph that we can edit
        doc.getSections().add(new Section(doc));

        Assert.assertEquals(0, doc.getSections().get(1).getChildNodes(NodeType.ANY, true).getCount());

        // Makes sure that the section contains a body with at least one paragraph
        doc.getLastSection().ensureMinimum();

        // Now we can add content to this section
        Assert.assertEquals(2, doc.getSections().get(1).getChildNodes(NodeType.ANY, true).getCount());
        //ExEnd
    }

    @Test
    public void bodyEnsureMinimum() throws Exception
    {
        //ExStart
        //ExFor:Section.Body
        //ExFor:Body.EnsureMinimum
        //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
        // Open a document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");

        // This shows what is in the document originally
        // The document has two sections
        Assert.assertEquals($"Section 1{ControlChar.SectionBreak}Section 2{ControlChar.SectionBreak}", doc.getText());

        // Loop through all sections in the document
        for (Section section : doc.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            // Each section has a Body node that contains main story (main text) of the section
            Body body = section.getBody();

            // This clears all nodes from the body
            body.removeAllChildren();

            // Technically speaking, for the main story of a section to be valid, it needs to have
            // at least one empty paragraph. That's what the EnsureMinimum method does
            body.ensureMinimum();
        }

        // Check how the content of the document looks now
        Assert.assertEquals($"{ControlChar.SectionBreak}{ControlChar.SectionBreak}", doc.getText());
        //ExEnd
    }

    @Test
    public void bodyNodeType() throws Exception
    {
        //ExStart
        //ExFor:Body.NodeType
        //ExFor:HeaderFooter.NodeType
        //ExFor:Document.FirstSection
        //ExSummary:Shows how you can enumerate through children of a composite node and detect types of the children nodes.
        // Open a document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Primary header");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Primary footer");

        // Get the first section in the document
        Section section = doc.getFirstSection();

        // A Section is a composite node and therefore can contain child nodes
        // Section can contain only Body and HeaderFooter nodes
        for (Node node : (Iterable<Node>) section)
        {
            // Every node has the NodeType property
            switch (node.getNodeType())
            {
                case NodeType.BODY:
                {
                    // If the node type is Body, we can cast the node to the Body class
                    Body body = (Body) node;

                    // Write the content of the main story of the section to the console
                    System.out.println("*** Body ***");
                    System.out.println(body.getText());
                    break;
                }
                case NodeType.HEADER_FOOTER:
                {
                    // If the node type is HeaderFooter, we can cast the node to the HeaderFooter class
                    HeaderFooter headerFooter = (HeaderFooter) node;

                    // Write the content of the header footer to the console
                    System.out.println("*** HeaderFooter ***");
                    msConsole.writeLine(headerFooter.getHeaderFooterType());
                    System.out.println(headerFooter.getText());
                    break;
                }
                default:
                {
                    // Other types of nodes never occur inside a Section node
                    throw new Exception("Unexpected node type in a section.");
                }
            }
        }
        //ExEnd
    }

    @Test
    public void sectionsDeleteAllSections() throws Exception
    {
        //ExStart
        //ExFor:NodeCollection.Clear
        //ExSummary:Shows how to remove all sections from a document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // All of the document's content is stored in the child nodes of sections like this one
        Assert.assertEquals("Hello World!", msString.trim(doc.getText()));
        Assert.assertEquals(5, doc.getSections().get(0).getChildNodes(NodeType.ANY, true).getCount());

        doc.getSections().clear();
        
        // Clearing the section collection effectively empties the document
        Assert.assertEquals("", doc.getText());
        Assert.assertEquals(0, doc.getSections().getCount());
        //ExEnd
    }

    @Test
    public void sectionsAppendSectionContent() throws Exception
    {
        //ExStart
        //ExFor:Section.AppendContent
        //ExFor:Section.PrependContent
        //ExSummary:Shows how to append content of an existing section. The number of sections in the document remains the same.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 3");

        // This is the section that we will append and prepend to
        Section section = doc.getSections().get(2);

        // This copies content of the 1st section and inserts it at the beginning of the specified section
        Section sectionToPrepend = doc.getSections().get(0);
        section.prependContent(sectionToPrepend);

        // This copies content of the 2nd section and inserts it at the end of the specified section
        Section sectionToAppend = doc.getSections().get(1);
        section.appendContent(sectionToAppend);

        Assert.assertEquals("Section 1" + ControlChar.SECTION_BREAK +
                        "Section 2" + ControlChar.SECTION_BREAK +
                        "Section 1" + ControlChar.PARAGRAPH_BREAK +
                        "Section 3" + ControlChar.PARAGRAPH_BREAK +
                        "Section 2" + ControlChar.SECTION_BREAK, doc.getText());
        //ExEnd
    }

    @Test
    public void sectionsDeleteSectionContent() throws Exception
    {
        //ExStart
        //ExFor:Section.ClearContent
        //ExSummary:Shows how to clear the content of a section.
        Document doc = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals("Hello World!", msString.trim(doc.getText()));

        doc.getFirstSection().clearContent();

        Assert.assertEquals("", msString.trim(doc.getText()));
        //ExEnd
    }

    @Test
    public void sectionsDeleteHeaderFooter() throws Exception
    {
        //ExStart
        //ExFor:Section.ClearHeadersFooters
        //ExSummary:Clears content of all headers and footers in a section.
        Document doc = new Document(getMyDir() + "Header and footer types.docx");

        Section section = doc.getSections().get(0);

        Assert.assertEquals(6, section.getHeadersFooters().getCount());
        Assert.assertEquals("First header", msString.trim(section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST).getText()));

        section.clearHeadersFooters();

        Assert.assertEquals(6, section.getHeadersFooters().getCount());
        Assert.assertEquals("", section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST).getText());
        //ExEnd
    }

    @Test
    public void sectionDeleteHeaderFooterShapes() throws Exception
    {
        //ExStart
        //ExFor:Section.DeleteHeaderFooterShapes
        //ExSummary:Removes all images and shapes from all headers footers in a section.
        Document doc = new Document();
        Section section = doc.getSections().get(0);
        HeaderFooter firstHeader = new HeaderFooter(doc, HeaderFooterType.HEADER_FIRST);

        section.getHeadersFooters().add(firstHeader);

        firstHeader.appendParagraph("This paragraph contains a shape: ");

        Shape shape = new Shape(doc, ShapeType.ARROW);
        firstHeader.getFirstParagraph().appendChild(shape);

        Assert.assertEquals(1, firstHeader.getChildNodes(NodeType.SHAPE, true).getCount());

        section.deleteHeaderFooterShapes();

        Assert.assertEquals(0, firstHeader.getChildNodes(NodeType.SHAPE, true).getCount());
        //ExEnd
    }

    @Test
    public void sectionsCloneSection() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");
        Section cloneSection = doc.getSections().get(0).deepClone();
    }

    @Test
    public void sectionsImportSection() throws Exception
    {
        Document srcDoc = new Document(getMyDir() + "Document.docx");
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section) dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);
    }

    @Test
    public void migrateFrom2XImportSection() throws Exception
    {
        Document srcDoc = new Document();
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section) dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);
    }

    @Test
    public void modifyPageSetupInAllSections() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");

        // It is important to understand that a document can contain many sections and each
        // section has its own page setup. In this case we want to modify them all
        for (Section section : doc.<Section>OfType() !!Autoporter error: Undefined expression type )
            section.getPageSetup().setPaperSize(PaperSize.LETTER);

        doc.save(getArtifactsDir() + "Section.ModifyPageSetupInAllSections.doc");
    }

    @Test
    public void cultureInfoPageSetupDefaults() throws Exception
    {
        CurrentThread.setCurrentCulture(new msCultureInfo("en-us"));

        Document docEn = new Document();

        // Assert that page defaults comply current culture info
        Section sectionEn = docEn.getSections().get(0);
        Assert.assertEquals(72.0, sectionEn.getPageSetup().getLeftMargin()); // 2.54 cm         
        Assert.assertEquals(72.0, sectionEn.getPageSetup().getRightMargin()); // 2.54 cm
        Assert.assertEquals(72.0, sectionEn.getPageSetup().getTopMargin()); // 2.54 cm
        Assert.assertEquals(72.0, sectionEn.getPageSetup().getBottomMargin()); // 2.54 cm
        Assert.assertEquals(36.0, sectionEn.getPageSetup().getHeaderDistance()); // 1.27 cm
        Assert.assertEquals(36.0, sectionEn.getPageSetup().getFooterDistance()); // 1.27 cm
        Assert.assertEquals(36.0, sectionEn.getPageSetup().getTextColumns().getSpacing()); // 1.27 cm

        // Change culture and assert that the page defaults are changed
        CurrentThread.setCurrentCulture(new msCultureInfo("de-de"));

        Document docDe = new Document();

        Section sectionDe = docDe.getSections().get(0);
        Assert.assertEquals(70.85, sectionDe.getPageSetup().getLeftMargin()); // 2.5 cm         
        Assert.assertEquals(70.85, sectionDe.getPageSetup().getRightMargin()); // 2.5 cm
        Assert.assertEquals(70.85, sectionDe.getPageSetup().getTopMargin()); // 2.5 cm
        Assert.assertEquals(56.7, sectionDe.getPageSetup().getBottomMargin()); // 2 cm
        Assert.assertEquals(35.4, sectionDe.getPageSetup().getHeaderDistance()); // 1.25 cm
        Assert.assertEquals(35.4, sectionDe.getPageSetup().getFooterDistance()); // 1.25 cm
        Assert.assertEquals(35.4, sectionDe.getPageSetup().getTextColumns().getSpacing()); // 1.25 cm

        // Change page defaults
        sectionDe.getPageSetup().setLeftMargin(90.0); // 3.17 cm
        sectionDe.getPageSetup().setRightMargin(90.0); // 3.17 cm
        sectionDe.getPageSetup().setTopMargin(72.0); // 2.54 cm
        sectionDe.getPageSetup().setBottomMargin(72.0); // 2.54 cm
        sectionDe.getPageSetup().setHeaderDistance(35.4); // 1.25 cm
        sectionDe.getPageSetup().setFooterDistance(35.4); // 1.25 cm
        sectionDe.getPageSetup().getTextColumns().setSpacing(35.4); // 1.25 cm

        docDe = DocumentHelper.saveOpen(docDe);

        Section sectionDeAfter = docDe.getSections().get(0);
        Assert.assertEquals(90.0, sectionDeAfter.getPageSetup().getLeftMargin()); // 3.17 cm         
        Assert.assertEquals(90.0, sectionDeAfter.getPageSetup().getRightMargin()); // 3.17 cm
        Assert.assertEquals(72.0, sectionDeAfter.getPageSetup().getTopMargin()); // 2.54 cm
        Assert.assertEquals(72.0, sectionDeAfter.getPageSetup().getBottomMargin()); // 2.54 cm
        Assert.assertEquals(35.4, sectionDeAfter.getPageSetup().getHeaderDistance()); // 1.25 cm
        Assert.assertEquals(35.4, sectionDeAfter.getPageSetup().getFooterDistance()); // 1.25 cm
        Assert.assertEquals(35.4, sectionDeAfter.getPageSetup().getTextColumns().getSpacing()); // 1.25 cm
    }
}
