module Asposewordsjavaforruby
  module SimpleMailMerge
    def initialize()
        mail_merge()
    end
        
    def mail_merge()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "MailMerge.doc")
        # Fill the fields in the document with user data.
        doc.getMailMerge().execute(
            Array["FullName", "Company", "Address", "Address2", "City"],
            Array["James Bond", "MI5 Headquarters", "Milbank", "", "London"]
        )
        # Saves the document to disk.
        doc.save(data_dir + "MailMerge Out.docx")
    end 

  end
end
