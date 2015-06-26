module Asposewordsjavaforruby
  module HandleMergeField
    def initialize()        
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/mailmerge/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "Template.doc")
        #$doc->getMailMerge()->setFieldMergingCallback(new HandleMergeField())

        fieldNames = Array["RecipientName","SenderName","FaxNumber","PhoneNumber","Subject","Body","Urgent","ForReview","PleaseComment"]
        fieldValues = Array["Josh","Jenny","123456789","","Hello","Test Pakistan 1", true, false, true]
        doc.getMailMerge().execute(fieldNames,fieldValues)

        # Save the document.
        doc.save(data_dir + "Template Out.doc")

        remove_empty_regions()        
    end

    def remove_empty_regions()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/mailmerge/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "TestFile.doc")

        # Create a dummy data source containing no data.
        data = Rjb::import('com.aspose.words.DataSet').new
        #DataSet data = new DataSet()

        # Set the appropriate mail merge clean up options to remove any unused regions from the document.
        mailmerge_cleanup_options = Rjb::import('com.aspose.words.MailMergeCleanupOptions')
        doc.getMailMerge().setCleanupOptions(mailmerge_cleanup_options.REMOVE_UNUSED_REGIONS)

        # Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
        # automatically as they are unused.
        doc.getMailMerge().executeWithRegions(data)

        # Save the output document to disk.
        doc.save(data_dir + "TestFile.RemoveEmptyRegions Out.doc")
    end

  end
end
