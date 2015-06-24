module Asposewordsjavaforruby
  module LoadAndSaveToStream
    def initialize()
        # Load and save to stream.
        save_to_stream()
    end
        
    def save_to_stream()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'

        # Open the stream. Read only access is enough for Aspose.Words to load a document.
        stream = Rjb::import('java.io.FileInputStream').new(data_dir + "Document.doc")

        # Load the entire document into memory.
        doc = Rjb::import('com.aspose.words.Document').new(stream)

        # You can close the stream now, it is no longer needed because the document is in memory.
        stream.close()
        # ... do something with the document
        # Convert the document to a different format and save to stream.
        dst_stream = Rjb::import("java.io.ByteArrayOutputStream").new()
        save_format = Rjb::import("com.aspose.words.SaveFormat")
        doc.save(dst_stream, save_format.RTF)

        output = Rjb::import("java.io.FileOutputStream").new(data_dir + "Document Out.rtf")
        output.write(dst_stream.toByteArray())
        output.close()
    end 

  end
end
