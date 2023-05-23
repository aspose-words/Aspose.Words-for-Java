package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Section;
import com.aspose.words.BreakType;
import com.aspose.words.PaperSize;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Body;
import com.aspose.ms.System.msConsole;
import com.aspose.words.HeaderFooter;
import com.aspose.words.Run;


class WorkingWithSection extends DocsExamplesBase
{
    @Test
    public void addSection() throws Exception
    {
        //ExStart:AddSection
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        builder.writeln("Hello2");

        Section sectionToAdd = new Section(doc);
        doc.getSections().add(sectionToAdd);
        //ExEnd:AddSection
    }

    @Test
    public void deleteSection() throws Exception
    {
        //ExStart:DeleteSection
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello2");
        doc.appendChild(new Section(doc));

        doc.getSections().removeAt(0);
        //ExEnd:DeleteSection
    }

    @Test
    public void deleteAllSections() throws Exception
    {
        //ExStart:DeleteAllSections
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello2");
        doc.appendChild(new Section(doc));

        doc.getSections().clear();
        //ExEnd:DeleteAllSections
    }

    @Test
    public void appendSectionContent() throws Exception
    {
        //ExStart:AppendSectionContent
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 3");

        Section section = doc.getSections().get(2);

        // Insert the contents of the first section to the beginning of the third section.
        Section sectionToPrepend = doc.getSections().get(0);
        section.prependContent(sectionToPrepend);

        // Insert the contents of the second section to the end of the third section.
        Section sectionToAppend = doc.getSections().get(1);
        section.appendContent(sectionToAppend);
        //ExEnd:AppendSectionContent
    }

    @Test
    public void cloneSection() throws Exception
    {
        //ExStart:CloneSection
        Document doc = new Document(getMyDir() + "Document.docx");
        Section cloneSection = doc.getSections().get(0).deepClone();
        //ExEnd:CloneSection
    }

    @Test
    public void copySection() throws Exception
    {
        //ExStart:CopySection
        Document srcDoc = new Document(getMyDir() + "Document.docx");
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section)dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);

        dstDoc.save(getArtifactsDir() + "WorkingWithSection.CopySection.docx");
        //ExEnd:CopySection
    }

    @Test
    public void deleteHeaderFooterContent() throws Exception
    {
        //ExStart:DeleteHeaderFooterContent
        Document doc = new Document(getMyDir() + "Document.docx");

        Section section = doc.getSections().get(0);
        section.clearHeadersFooters();
        //ExEnd:DeleteHeaderFooterContent
    }

    @Test
    public void deleteSectionContent() throws Exception
    {
        //ExStart:DeleteSectionContent
        Document doc = new Document(getMyDir() + "Document.docx");

        Section section = doc.getSections().get(0);
        section.clearContent();
        //ExEnd:DeleteSectionContent
    }

    @Test
    public void modifyPageSetupInAllSections() throws Exception
    {
        //ExStart:ModifyPageSetupInAllSections
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Section 1");
        doc.appendChild(new Section(doc));
        builder.writeln("Section 2");
        doc.appendChild(new Section(doc));
        builder.writeln("Section 3");
        doc.appendChild(new Section(doc));
        builder.writeln("Section 4");

        // It is important to understand that a document can contain many sections,
        // and each section has its page setup. In this case, we want to modify them all.
        for (Section section : (Iterable<Section>) doc)
            section.getPageSetup().setPaperSize(PaperSize.LETTER);

        doc.save(getArtifactsDir() + "WorkingWithSection.ModifyPageSetupInAllSections.doc");
        //ExEnd:ModifyPageSetupInAllSections
    }

    @Test
    public void sectionsAccessByIndex() throws Exception
    {
        //ExStart:SectionsAccessByIndex
        Document doc = new Document(getMyDir() + "Document.docx");

        Section section = doc.getSections().get(0);
        section.getPageSetup().setLeftMargin(90.0); // 3.17 cm
        section.getPageSetup().setRightMargin(90.0); // 3.17 cm
        section.getPageSetup().setTopMargin(72.0); // 2.54 cm
        section.getPageSetup().setBottomMargin(72.0); // 2.54 cm
        section.getPageSetup().setHeaderDistance(35.4); // 1.25 cm
        section.getPageSetup().setFooterDistance(35.4); // 1.25 cm
        section.getPageSetup().getTextColumns().setSpacing(35.4); // 1.25 cm
        //ExEnd:SectionsAccessByIndex
    }

    @Test
    public void sectionChildNodes() throws Exception
    {
        //ExStart:SectionChildNodes
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Primary header");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Primary footer");

        Section section = doc.getFirstSection();

        // A Section is a composite node and can contain child nodes,
        // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
        for (Node node : (Iterable<Node>) section)
        {
            switch (node.getNodeType())
            {
                case NodeType.BODY:
                    {
                        Body body = (Body)node;

                        System.out.println("Body:");
                        System.out.println("\t\"{body.GetText().Trim()}\"");
                        break;
                    }
                case NodeType.HEADER_FOOTER:
                    {
                        HeaderFooter headerFooter = (HeaderFooter)node;

                        System.out.println("HeaderFooter type: {headerFooter.HeaderFooterType}:");
                        System.out.println("\t\"{headerFooter.GetText().Trim()}\"");
                        break;
                    }
                default:
                    {
                        throw new Exception("Unexpected node type in a section.");
                    }
            }
        }
        //ExEnd:SectionChildNodes
    }

    @Test
    public void ensureMinimum() throws Exception
    {
        //ExStart:EnsureMinimum
        Document doc = new Document();

        // If we add a new section like this, it will not have a body, or any other child nodes.
        doc.getSections().add(new Section(doc));
        // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
        doc.getLastSection().ensureMinimum();
        
        doc.getSections().get(0).getBody().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd:EnsureMinimum
    }
}
