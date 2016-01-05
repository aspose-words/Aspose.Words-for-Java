from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import SaveFormat

from java.io import ByteArrayOutputStream
from java.io import FileInputStream
from java.io import FileOutputStream

class LoadAndSaveToStream:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        # Open the stream. Read only access is enough for Aspose.Words to load a document.
        stream = FileInputStream(dataDir + 'Document.doc')
        
        # Load the entire document into memory.
        doc = Document(stream)
        
        # You can close the stream now, it is no longer needed because the document is in memory.
        stream.close()
        
        # ... do something with the document
        # Convert the document to a different format and save to stream.
        dstStream = ByteArrayOutputStream()
        doc.save(dstStream, SaveFormat.RTF)
        output = FileOutputStream(dataDir + "Document Out.rtf")
        output.write(dstStream.toByteArray())
        output.close()

        print "Document loaded from stream and then saved successfully."

if __name__ == '__main__':           
    LoadAndSaveToStream()