from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import DocumentBuilder

class HelloWorld:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        doc = Document()
        
        builder = DocumentBuilder(doc)
        builder.writeln('Hello World!')
        
        doc.save(dataDir + 'HelloWorld.docx')
        
        print "Document saved."

if __name__ == '__main__':        
    HelloWorld()