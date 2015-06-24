module Asposewordsjavaforruby
  module AppendDocument
    def initialize()
        # The path to the documents directory.
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/joiningandappending/'
    
        simple_append_document()
        keep_source_formatting()
        use_destination_styles()
        join_continuous()
        join_new_page()
        restart_page_numbering()
        link_headers_footers()
        unlink_headers_footers()
        remove_source_headers_footers()
        different_page_setup()
    end

=begin
    Shows how to append a document to the end of another document using no additional options.
=end
    def simple_append_document()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Append the source document to the destination document using no extra options.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)
        
        # Save the document.
        dst_doc.save(@data_dir + "TestFile.SimpleAppendDocument Out.docx")
    end

=begin
    Shows how to append a document to another document while keeping the original formatting.
=end
    def keep_source_formatting()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Append the source document to the destination document using no extra options.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)
        
        # Save the document.
        dst_doc.save(@data_dir + "TestFile.KeepSourceFormatting Out.docx")
    end

=begin
    Shows how to append a document to another document using the formatting of the destination document.
=end
    def use_destination_styles()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Append the source document to the destination document using no extra options.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.USE_DESTINATION_STYLES)
        
        # Save the document.
        dst_doc.save(@data_dir + "TestFile.UseDestinationStyles Out.docx")
    end

=begin
    Shows how to append a document to another document so the content flows continuously.
=end
    def join_continuous()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Make the document appear straight after the destination documents content.
        section_start = Rjb::import("com.aspose.words.SectionStart")
        src_doc.getFirstSection().getPageSetup().setSectionStart(section_start.CONTINUOUS)
        
        # Append the source document using the original styles found in the source document.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)
        
        # Save the document.
        dst_doc.save(@data_dir + "TestFile.JoinContinuous Out.docx") 
    end

=begin
    Shows how to append a document to another document so it starts on a new page.
=end
    def join_new_page()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Set the appended document to start on a new page.
        section_start = Rjb::import("com.aspose.words.SectionStart")
        src_doc.getFirstSection().getPageSetup().setSectionStart(section_start.NEW_PAGE)
        
        # Append the source document using the original styles found in the source document.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)
        
        # Save the document.
        dst_doc.save(@data_dir + "TestFile.JoinNewPage Out.docx")  
    end 

=begin
    Shows how to append a document to another document with page numbering restarted.
=end
    def restart_page_numbering()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Set the appended document to start on a new page.
        section_start = Rjb::import("com.aspose.words.SectionStart")
        src_doc.getFirstSection().getPageSetup().setSectionStart(section_start.NEW_PAGE)

        # Restart the page numbering for the document to be appended.
        src_doc.getFirstSection().getPageSetup().setRestartPageNumbering(true)
        
        # Append the source document using the original styles found in the source document.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)
        
        # Save the document.
        dst_doc.save(@data_dir + "TestFile.RestartPageNumbering Out.docx")
    end

=begin
    Shows how to append a document to another document and continue headers and footers from the destination document.
=end
    def link_headers_footers()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Set the appended document to start on a new page.
        section_start = Rjb::import("com.aspose.words.SectionStart")
        src_doc.getFirstSection().getPageSetup().setSectionStart(section_start.NEW_PAGE)
        
        # Link the headers and footers in the source document to the previous section.
        # This will override any headers or footers already found in the source document.
        src_doc.getFirstSection().getHeadersFooters().linkToPrevious(true)
        
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)

        # Save the document.
        dst_doc.save(@data_dir + "TestFile.LinkHeadersFooters Out.docx")
    end

=begin
    Shows how to append a document to another document so headers and footers do not continue from the destination document.
=end
    def unlink_headers_footers()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
        # Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        # from the destination document.
        src_doc.getFirstSection().getHeadersFooters().linkToPrevious(false)
        
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)

        # Save the document.
        dst_doc.save(@data_dir + "TestFile.UnlinkHeadersFooters Out.docx")
    end

=begin
    Shows how to append a document to another document so headers and footers do not continue from the destination document.
=end
    def remove_source_headers_footers()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Remove the headers and footers from each of the sections in the source document.
        sections = src_doc.getSections().toArray()
        sections.each do |section|
            section.clearHeadersFooters()    
        end
        
        # Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        # for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        # document. This should set to false to avoid this behaviour.
        src_doc.getFirstSection().getHeadersFooters().linkToPrevious(false)
        
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)

        # Save the document.
        dst_doc.save(@data_dir + "TestFile.RemoveSourceHeadersFooters Out.docx")        
    end

=begin
    Shows how to append a document to another document continuously which has different page settings.
=end
    def different_page_setup()
        # Load the documents to join.
        dst_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import("com.aspose.words.Document").new(@data_dir + "TestFile.Source.doc")
        
        # Set the source document to continue straight after the end of the destination document.
        # If some page setup settings are different then this may not work and the source document will appear
        # on a new page.
        section_start = Rjb::import("com.aspose.words.SectionStart")
        src_doc.getFirstSection().getPageSetup().setSectionStart(section_start.CONTINUOUS)

        # To ensure this does not happen when the source document has different page setup settings make sure the
        # settings are identical between the last section of the destination document.
        # If there are further continuous sections that follow on in the source document then this will need to be
        # repeated for those sections as well.
        src_doc.getFirstSection().getPageSetup().setPageWidth(dst_doc.getLastSection().getPageSetup().getPageWidth())
        src_doc.getFirstSection().getPageSetup().setPageHeight(dst_doc.getLastSection().getPageSetup().getPageHeight())
        src_doc.getFirstSection().getPageSetup().setOrientation(dst_doc.getLastSection().getPageSetup().getOrientation())
        
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        dst_doc.appendDocument(src_doc, import_format_mode.KEEP_SOURCE_FORMATTING)

        # Save the document.
        dst_doc.save(@data_dir + "TestFile.DifferentPageSetup Out.docx")        
    end 
        
  end
end
