module Asposewordsjavaforruby
  module ImageToPdf
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/imagetopdf/'

        convert_image_to_pdf(data_dir + "Test.jpg", data_dir + "TestJpg Out.pdf")
        convert_image_to_pdf(data_dir + "Test.png", data_dir + "TestPng Out.pdf")
        convert_image_to_pdf(data_dir + "Test.bmp", data_dir + "TestBmp Out.pdf")
        convert_image_to_pdf(data_dir + "Test.gif", data_dir + "TestGif Out.pdf")
    end

    def convert_image_to_pdf(input_filename, output_filename)
        # Create Aspose.Words.Document and DocumentBuilder.
        # The builder makes it simple to add content to the document.
        doc = Rjb::import('com.aspose.words.Document').new
        builder = Rjb::import('com.aspose.words.DocumentBuilder').new(doc)

        # Load images from the disk using the approriate reader.
        # The file formats that can be loaded depends on the image readers available on the machine.
        imageio = Rjb::import('javax.imageio.ImageIO')
        ImageInputStream = Rjb::import('javax.imageio.stream.ImageInputStream')
        reader = Rjb::import('javax.imageio.ImageReader')
        
        iis = imageio.createImageInputStream(Rjb::import('java.io.File').new(input_filename))
        reader = imageio.getImageReaders(iis).next()
        reader.setInput(iis, false)

        # Get the number of frames in the image.
        framesCount = reader.getNumImages(true)

        # Loop through all frames.
        for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
        {
            # Insert a section break before each new page, in case of a multi-frame image.
            if (frameIdx != 0) then
                break_type = Rjb::import('com.aspose.words.BreakType')
                builder.insertBreak(break_type.SECTION_BREAK_NEW_PAGE)
            end    

            # Select active frame.
            image = Rjb::import('java.awt.image.BufferedImage')
            image = reader.read(frameIdx)

            # We want the size of the page to be the same as the size of the image.
            # Convert pixels to points to size the page to the actual image size.
            ps = Rjb::import('com.aspose.words.PageSetup')
            ps = builder.getPageSetup()

            convert_util = Rjb::import('com.aspose.words.ConvertUtil')
            ps.setPageWidth(convert_util.pixelToPoint(image.getWidth()))
            ps.setPageHeight(convert_util.pixelToPoint(image.getHeight()))

            # Insert the image into the document and position it at the top left corner of the page.
            builder.insertImage(
                image,
                RelativeHorizontalPosition.PAGE,
                0,
                RelativeVerticalPosition.PAGE,
                0,
                ps.getPageWidth(),
                ps.getPageHeight(),
                WrapType.NONE)

            # Save the document.
            doc.save(output_filename)
        end
    end

  end
end
