__author__ = 'fahadadeel'
import jpype

class HelloWorld:

    def __init__(self, dataDir):
        self.dataDir = dataDir

    def main(self):
        """
            : The path to the documents directory. :
        """

        Document = jpype.JClass("com.aspose.words.Document")

        DocumentBuilder = jpype.JClass("com.aspose.words.DocumentBuilder")

        doc = Document()
        builder = DocumentBuilder(doc)

        builder.writeln('Hello World!')
        doc.save(self.dataDir +'HelloWorld.docx')


class AppendDocuments:

    def main(self, outputDocFile,dstDocFile,srcDocFile):

        self.outputDocFile = outputDocFile
        self.dstDocFile = dstDocFile
        self.srcDocFile = srcDocFile

        Document = jpype.JClass("com.aspose.words.Document")
        ImportFormatMode = jpype.JClass("com.aspose.words.ImportFormatMode")

        dstDoc = Document(self.dstDocFile)
        srcDoc = Document(self.srcDocFile)

        dstDoc.appendDocument(srcDoc,ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dstDoc.save(self.outputDocFile)

class Doc2Pdf:

     def main(self,srcDocFile,dstPdfFile):

         self.srcDocFile = srcDocFile
         self.dstPdfFile = dstPdfFile

         Document =jpype.JClass("com.aspose.words.Document")

         doc = Document(self.srcDocFile)
         doc.save(self.dstPdfFile)

class FindAndReplace:

    def main(self,srcDocFile,outputDocFile,searchString,replaceString):

        Document =jpype.JClass("com.aspose.words.Document")
        FindReplaceDirection =jpype.JClass("com.aspose.words.FindReplaceDirection")
        FindReplaceOptions =jpype.JClass("com.aspose.words.FindReplaceOptions")

        doc = Document(srcDocFile)

        print "Original document text: " + doc.getRange().getText()

        doc.getRange().replace(searchString, replaceString, FindReplaceOptions(FindReplaceDirection.FORWARD))
        # Check the replacement was made.

        print "Document text after replace: " + doc.getRange().getText()

        # Save the modified document.
        doc.save(outputDocFile)

class LoadAndSaveToDisk:

    def main(self,srcDocFile,outputDocFile):

        Document =jpype.JClass("com.aspose.words.Document")

        # Load the document from the absolute path on disk.
        doc = Document(srcDocFile)
        # Save the document as DOCX document.
        doc.save(outputDocFile)

class LoadAndSaveToStream:

    def main(self,srcDocFile,outputRtfFile):

        Document =jpype.JClass("com.aspose.words.Document")
        FileInputStream = jpype.JClass("java.io.FileInputStream")
        FileOutputStream = jpype.JClass("java.io.FileOutputStream")

        ByteArrayOutputStream = jpype.JClass("java.io.ByteArrayOutputStream")
        SaveFormat = jpype.JClass("com.aspose.words.SaveFormat")

        # Open the stream. Read only access is enough for Aspose.Words to load a document.
        stream = FileInputStream(srcDocFile)

        # Load the entire document into memory.
        doc = Document(stream)

        # You can close the stream now, it is no longer needed because the document is in memory.
        stream.close()

        # ... do something with the document
        # Convert the document to a different format and save to stream.
        dstStream = ByteArrayOutputStream()
        doc.save(dstStream, SaveFormat.RTF)
        output = FileOutputStream(outputRtfFile)
        output.write(dstStream.toByteArray())
        output.close()

class SimpleMailMerge:

    def main(self,srcDocFile,outputFile):

        Document =jpype.JClass("com.aspose.words.Document")
        doc = Document(srcDocFile)
        # Fill the fields in the document with user data.
        doc.getMailMerge().execute(
                ["FullName", "Company", "Address", "Address2", "City"],
                ["James Bond", "MI5 Headquarters", "Milbank", "", "London"])
        # Saves the document to disk.
        doc.save(outputFile)

class UpdateFields:

    def main(self,outputFile):
        
        Document = jpype.JClass("com.aspose.words.Document")
        DocumentBuilder = jpype.JClass("com.aspose.words.DocumentBuilder")
        BreakType = jpype.JClass("com.aspose.words.BreakType")
        StyleIdentifier = jpype.JClass("com.aspose.words.StyleIdentifier")

        # Demonstrates how to insert fields and update them using Aspose.Words.
        # First create a blank document.
        doc = Document()
        # Use the document builder to insert some content and fields.
        builder = DocumentBuilder(doc)
        # Insert a table of contents at the beginning of the document.
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u")
        builder.writeln()
        # Insert some other fields.
        builder.write("Page: ")
        builder.insertField("PAGE")
        builder.write(" of ")
        builder.insertField("NUMPAGES")
        builder.writeln()
        builder.write("Date: ")
        builder.insertField("DATE")
        # Start the actual document content on the second page.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE)
        # Build a document with complex structure by applying different heading styles thus creating TOC entries.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1)
        builder.writeln("Heading 1")
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2)
        builder.writeln("Heading 1.1")
        builder.writeln("Heading 1.2")
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1)
        builder.writeln("Heading 2")
        builder.writeln("Heading 3")
        # Move to the next page.
        builder.insertBreak(BreakType.PAGE_BREAK)
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2)
        builder.writeln("Heading 3.1")
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3)
        builder.writeln("Heading 3.1.1")
        builder.writeln("Heading 3.1.2")
        builder.writeln("Heading 3.1.3")
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2)
        builder.writeln("Heading 3.2")
        builder.writeln("Heading 3.3")

        print "Updating all fields in the document."

        # Call the method below to update the TOC.
        doc.updateFields()
        doc.save(outputFile)

class WorkingWithNodes:

    def main(self):
        
        Document = jpype.JClass("com.aspose.words.Document")
        Paragraph = jpype.JClass("com.aspose.words.Paragraph")
        Node = jpype.JClass("com.aspose.words.Node")

        # Create a new document.
        doc = Document()

        # Creates and adds a paragraph node to the document.
        para = Paragraph(doc)

        # Typed access to the last section of the document.
        section = doc.getLastSection()
        section.getBody().appendChild(para)

        # Next print the node type of one of the nodes in the document.
        nodeType = doc.getFirstSection().getBody().getNodeType()

        print "NodeType: " + Node.nodeTypeToString(nodeType)

class ApplyLicense:

    def main(self):

        License = jpype.JClass("com.aspose.words.License")

        # This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        # You can also use the additional overload to load a license from a stream, this is useful for instance when the
        # license is stored as an embedded resource
        try:
            license = License()
            license.setLicense("Aspose.Words.lic")
        except Exception as e:
            # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            print "There was an error setting the license: "
