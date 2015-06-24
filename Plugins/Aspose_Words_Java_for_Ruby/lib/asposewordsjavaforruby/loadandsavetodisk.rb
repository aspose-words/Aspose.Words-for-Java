module Asposewordsjavaforruby
  module LoadAndSaveToDisk
    def initialize()
        # Load and save the document.
        save_to_disk()
    end
        
    def save_to_disk()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "Document.doc")

        # Save the document as DOCX document.
        doc.save(data_dir + "Document Out.docx")
    end 

  end
end
