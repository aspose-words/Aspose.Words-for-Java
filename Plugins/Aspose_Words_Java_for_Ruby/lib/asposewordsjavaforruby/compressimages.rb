module Asposewordsjavaforruby
  module CompressImages
    def initialize()
        # The path to the documents directory.
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        srcFileName = @data_dir + "TestCompressImages.docx"

        doc = Rjb::import('com.aspose.words.Document').new(@data_dir + "TestCompressImages.docx")

        # Demonstrate autofitting a table to the window.
        compress_images(doc, srcFileName)
    end

    def compress_images(doc, srcFileName)
        messageFormat = Rjb::import("java.text.MessageFormat")
        file_size = get_file_size(srcFileName)
        
        # 220ppi Print - said to be excellent on most printers and screens.
        # 150ppi Screen - said to be good for web pages and projectors.
        # 96ppi Email - said to be good for minimal document size and sharing.
        desiredPpi = 150
        # In Java this seems to be a good compression / quality setting.
        jpegQuality = 90

        # Resample images to desired ppi and save.
        resampler = Rjb::import("com.aspose.words.Resampler").new
        count = resampler.resample(doc, desiredPpi, jpegQuality)
        puts MessageFormat.format("Resampled {0} images.", count)
        if (count != 1) then
             puts "We expected to have only 1 image resampled in this test document!"
        end    
        dstFileName = @data_dir + "TestCompressImages Out.docx"
        doc.save(dstFileName)
        puts messageFormat.format("Saving {0}. Size {1}.", dstFileName, get_file_size(dstFileName))

        # Verify that the first image was compressed by checking the new Ppi.
        dst_doc = Rjb::import("com.aspose.words.Document").new(dstFileName)
        nodeType = Rjb::import("com.aspose.words.NodeType")
        shape = dst_doc.getChild(nodeType.DRAWING_ML, 0, true)
        convertUtil = Rjb::import("com.aspose.words.ConvertUtil")
        imagePpi = shape.getImageData().getImageSize().getWidthPixels() / convertUtil.pointToInch(shape.getSize().getX())
        if (imagePpi < 150) then
            puts "Image was not resampled successfully."
        end
    end

    def get_file_size(file_name)
        file = Rjb::import("java.io.File").new(file_name)
        return file.length()
    end

  end
end
