from asposewords import Settings
from com.aspose.words import Document

class SimpleMailMerge:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        doc = Document(dataDir + "MailMerge.doc");
        
        # Fill the fields in the document with user data.
        doc.getMailMerge().execute(
                ["FullName", "Company", "Address", "Address2", "City"],
                ["James Bond", "MI5 Headquarters", "Milbank", "", "London"])
                
        # Saves the document to disk.
        doc.save(dataDir + "MailMerge Result Out.docx")

        print "Mail merge performed successfully."
        
if __name__ == '__main__':
    SimpleMailMerge()