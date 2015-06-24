module Asposewordsjavaforruby
  module FindAndReplace
    def initialize()
        # Find and replace text in document.
        replace_text()
    end
        
    def replace_text()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "ReplaceSimple.doc")

        # Check the text of the document.
        puts "Original document text: " + doc.getRange().getText()
        
        # Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", false, false)
        
        # Check the replacement was made.
        puts "Document text after replace: " + doc.getRange().getText()
        
        # Save the modified document.
        doc.save(data_dir + "ReplaceSimple Out.doc")
    end 

  end
end
