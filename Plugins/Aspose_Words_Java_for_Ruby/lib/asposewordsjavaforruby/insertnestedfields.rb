module Asposewordsjavaforruby
  module InsertNestedFields
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Create new document.
        doc = Rjb::import('com.aspose.words.Document').new
        builder = Rjb::import("com.aspose.words.DocumentBuilder").new(doc)
        
        # Insert few page breaks (just for testing)
        breakType = Rjb::import("com.aspose.words.BreakType")
        
        for i in 0..4    
            builder.insertBreak(breakType.PAGE_BREAK)
        end
            
        # Move DocumentBuilder cursor into the primary footer.
        headerFooterType = Rjb::import("com.aspose.words.HeaderFooterType")
        builder.moveToHeaderFooter(headerFooterType.FOOTER_PRIMARY)
        
        # We want to insert a field like this:
        # { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        field = builder.insertField("IF ")
        builder.moveTo(field.getSeparator())
        builder.insertField("PAGE")
        builder.write(" <> ")
        builder.insertField("NUMPAGES")
        builder.write(" \"See Next Page\" \"Last Page\" ")
        
        # Finally update the outer field to recalcaluate the final value. Doing this will automatically update
        # the inner fields at the same time.
        field.update()
       
        # Save the document.
        doc.save(data_dir + "InsertNestedFields Out.doc")
    end
  end
end
