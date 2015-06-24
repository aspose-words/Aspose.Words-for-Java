module Asposewordsjavaforruby
  module UpdateFields
    def initialize()
        update_fields()
    end
        
    def update_fields()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'

        # Demonstrates how to insert fields and update them using Aspose.Words.
        # First create a blank document.
        doc = Rjb::import('com.aspose.words.Document').new()
        # Use the document builder to insert some content and fields.
        builder = Rjb::import('com.aspose.words.DocumentBuilder').new(doc)
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
        break_type = Rjb::import("com.aspose.words.BreakType")
        builder.insertBreak(break_type.SECTION_BREAK_NEW_PAGE)

        # Build a document with complex structure by applying different heading styles thus creating TOC entries.
        style_identifier = Rjb::import("com.aspose.words.StyleIdentifier")
        builder.getParagraphFormat().setStyleIdentifier(style_identifier.HEADING_1)
        builder.writeln("Heading 1")
        builder.getParagraphFormat().setStyleIdentifier(style_identifier.HEADING_2)
        builder.writeln("Heading 1.1")
        builder.writeln("Heading 1.2")
        builder.getParagraphFormat().setStyleIdentifier(style_identifier.HEADING_1)
        builder.writeln("Heading 2")
        builder.writeln("Heading 3")

        # Move to the next page.
        builder.insertBreak(break_type.PAGE_BREAK)
        builder.getParagraphFormat().setStyleIdentifier(style_identifier.HEADING_2)
        builder.writeln("Heading 3.1")
        builder.getParagraphFormat().setStyleIdentifier(style_identifier.HEADING_3)
        builder.writeln("Heading 3.1.1")
        builder.writeln("Heading 3.1.2")
        builder.writeln("Heading 3.1.3")
        builder.getParagraphFormat().setStyleIdentifier(style_identifier.HEADING_2)
        builder.writeln("Heading 3.2")
        builder.writeln("Heading 3.3")
        puts "Updating all fields in the document."
        
        # Call the method below to update the TOC.
        doc.updateFields()
        doc.save(data_dir + "Document Field Update Out.docx")
    end

  end
end
