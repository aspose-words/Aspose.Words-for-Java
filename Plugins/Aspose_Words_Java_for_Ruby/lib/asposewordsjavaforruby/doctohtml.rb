module Asposewordsjavaforruby
  module DocToHTML
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "TestFile.doc")

        #HtmlSaveOptions options = new HtmlSaveOptions();
        options = Rjb::import('com.aspose.words.HtmlSaveOptions').new

        # HtmlSaveOptions.ExportRoundtripInformation property specifies
        # whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
        # Default value is true for HTML and false for MHTML and EPUB.
        options.setExportRoundtripInformation(true)
        doc.save(data_dir + "ExportRoundtripInformation Out.html", options)

        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "ExportRoundtripInformation Out.html")

        # Save the document Docx file format
        save_format = Rjb::import('com.aspose.words.SaveFormat')
        doc.save(data_dir + "Out.docx", save_format.DOCX)
    end
  end
end
