// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.Document;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.HtmlControlType;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.FormField;


@Test
class ExHtmlLoadOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void supportVml() throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.SupportVml
        //ExSummary:Shows how to parse HTML document with conditional comments like "<!--[if gte vml 1]>" and "<![if !vml]>"
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        //If value is true, then we parse "<!--[if gte vml 1]>", else parse "<![if !vml]>"
        loadOptions.setSupportVml(true);
        //Wait for a response, when loading external resources
        loadOptions.setWebRequestTimeout(1000);

        Document doc = new Document(getMyDir() + "Shape.VmlAndDml.htm", loadOptions);
        doc.save(getArtifactsDir() + "Shape.VmlAndDml.docx");
        //ExEnd
    }

    @Test
    public void webRequestTimeoutDefaultValue()
    {
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        msAssert.areEqual(100000, loadOptions.getWebRequestTimeout());
    }

    @Test
    public void getSelectAsSdt() throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.PreferredControlType
        //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
        final String HTML = "\r\n                <html>\r\n                    <select name='ComboBox' size='1'>\r\n                        <option value='val1'>item1</option>\r\n                        <option value='val2'></option>                        \r\n                    </select>\r\n                </html>\r\n            ";

        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
        htmlLoadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG);

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(HTML)), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag tag = (StructuredDocumentTag) nodes.get(0);
        //ExEnd

        msAssert.areEqual(2, tag.getListItems().getCount());

        msAssert.areEqual("val1", tag.getListItems().get(0).getValue());
        msAssert.areEqual("val2", tag.getListItems().get(1).getValue());
    }

    @Test
    public void getInputAsFormField() throws Exception
    {
        final String HTML = "\r\n                <html>\r\n                    <input type='text' value='Input value text' />\r\n                </html>\r\n            ";

        // By default "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField"
        // So, we do not set this value
        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(HTML)), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.FORM_FIELD, true);

        msAssert.areEqual(1, nodes.getCount());

        FormField formField = (FormField) nodes.get(0);
        msAssert.areEqual("Input value text", formField.getResult());
    }
}
