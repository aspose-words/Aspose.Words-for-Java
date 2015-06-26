module Asposewordsjavaforruby
  module LoadTxt
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "LoadTxt.txt")

        # Save as any Aspose.Words supported format, such as DOCX.
        doc.save(data_dir + "LoadTxt Out.doc")
    end
  end
end
