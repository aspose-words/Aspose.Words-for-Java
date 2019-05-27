package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.words.Document;
import com.aspose.ms.System.IO.Stream;


public class DocumentTestClass
{
    public Document getDocument() { return mDocument; }; public void setDocument(Document value) { mDocument = value; };

    private Document mDocument;
    public Stream getDocumentStream() { return mDocumentStream; }; public void setDocumentStream(Stream value) { mDocumentStream = value; };

    private Stream mDocumentStream;
    public byte[] getDocumentBytes() { return mDocumentBytes; }; public void setDocumentBytes(byte[] value) { mDocumentBytes = value; };

    private byte[] mDocumentBytes;
    public String getDocumentUri() { return mDocumentUri; }; public void setDocumentUri(String value) { mDocumentUri = value; };

    private String mDocumentUri;

    public DocumentTestClass(Document doc, Stream docStream, byte[] docBytes, String docUri)
    {
        setDocument(doc);
        setDocumentStream(docStream);
        setDocumentBytes(docBytes);
        setDocumentUri(docUri);
    }
}
