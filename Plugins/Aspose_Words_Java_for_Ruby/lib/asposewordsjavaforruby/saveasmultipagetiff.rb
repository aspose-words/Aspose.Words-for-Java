module Asposewordsjavaforruby
  module SaveAsMultipageTiff
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "TestFile.doc")

        # Save the document as multipage TIFF.
        doc.save(data_dir + "TestFile Out.doc")

        save_format = Rjb::import('com.aspose.words.SaveFormat')

        options = Rjb::import('com.aspose.words.ImageSaveOptions').new(save_format.TIFF)
        options.setPageIndex(0)
        options.setPageCount(2)
        
        tiff_compression = Rjb::import('com.aspose.words.TiffCompression')
        options.setTiffCompression(tiff_compression.CCITT_4)
        options.setResolution(160)

        doc.save(data_dir + "TestFileWithOptions Out.tiff", options)
    end
  end
end
