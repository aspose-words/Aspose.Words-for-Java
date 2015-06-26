module Asposewordsjavaforruby
  module Doc2Pdf

    def doc_to_pdf()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Open document.
        document = Rjb::import('com.aspose.words.Document').new(data_dir + "Template.doc")
        
        # Save the document in PDF format.
        document.save(data_dir + "Doc2PdfSave Out.pdf") 
    end

  end
end
