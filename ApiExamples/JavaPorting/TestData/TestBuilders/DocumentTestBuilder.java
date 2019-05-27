package ApiExamples.TestData.TestBuilders;

// ********* THIS FILE IS AUTO PORTED *********

import ApiExamples.ApiExampleBase;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.Stream;
import ApiExamples.TestData.TestClasses.DocumentTestClass;


public class DocumentTestBuilder extends ApiExampleBase
{
    private Document mDocument;
    private Stream mDocumentStream;
    private byte[] mDocumentBytes;
    private String mDocumentUri;

    public DocumentTestBuilder() throws Exception
    {
        mDocument = new Document();
        mDocumentStream = Stream.Null;
        mDocumentBytes = new byte[0];
        mDocumentUri = "";
    }

    public DocumentTestBuilder withDocument(Document doc)
    {
        mDocument = doc;
        return this;
    }

    public DocumentTestBuilder withDocumentStream(Stream stream)
    {
        mDocumentStream = stream;
        return this;
    }

    public DocumentTestBuilder withDocumentBytes(byte[] docBytes)
    {
        mDocumentBytes = docBytes;
        return this;
    }

    public DocumentTestBuilder withDocumentUri(String docUri)
    {
        mDocumentUri = docUri;
        return this;
    }

    public DocumentTestClass build()
    {
        return new DocumentTestClass(mDocument, mDocumentStream, mDocumentBytes, mDocumentUri);
    }
}
