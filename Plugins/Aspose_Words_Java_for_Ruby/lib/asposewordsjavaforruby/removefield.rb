module Asposewordsjavaforruby
  module RemoveField
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "Field.RemoveField.doc")

        #ExStart
        #ExFor:Field.Remove
        #ExId:DocumentBuilder_RemoveField
        #ExSummary:Removes a field from the document.
        field = doc.getRange().getFields().get(0)
        # Calling this method completely removes the field from the document.
        field.remove()
        #ExEnd

        # Save the document.
        doc.save(data_dir + "Field.RemoveField Out.doc")
    end
  end
end
