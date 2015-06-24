module Asposewordsjavaforruby
  module HelloWorld
    def initialize()
        # Create document.
        print_hello_world()
    end
        
    def print_hello_world()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'
        
        #Create a blank document.
        document = Rjb::import('com.aspose.words.Document').new()
        
        #DocumentBuilder provides members to easily add content to a document.
        builder = Rjb::import('com.aspose.words.DocumentBuilder').new(document)
        
        #Write a new paragraph in the document with the text "Hello World!"
        builder.writeln("Hello World!")
        
        # Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
        # Aspose.Words supports saving any document in many more formats.
        document.save(data_dir + "HelloWorld.docx") 
    end

  end
end
