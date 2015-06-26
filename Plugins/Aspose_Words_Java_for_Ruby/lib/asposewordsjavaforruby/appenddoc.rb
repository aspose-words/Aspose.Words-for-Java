module Asposewordsjavaforruby
  module AppendDoc
    def initialize()
        # Append document.
        append_documents()
    end
        
    def append_documents()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'
        
        # Load the destination and source documents from disk.
        dst_doc = Rjb::import('com.aspose.words.Document').new(data_dir + "TestFile.Destination.doc")
        src_doc = Rjb::import('com.aspose.words.Document').new(data_dir + "TestFile.Source.doc")

        importformatmode = Rjb::import('com.aspose.words.ImportFormatMode')
        source_formating = importformatmode.KEEP_SOURCE_FORMATTING

        # Append the source document to the destination document while keeping the original formatting of the source document.
        dst_doc.appendDocument(src_doc, source_formating)
        dst_doc.save(data_dir + "TestFile Out.docx")
    end

  end
end
