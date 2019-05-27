package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.Document;
import org.testng.Assert;
import com.aspose.words.HtmlControlType;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.FormField;

import java.io.ByteArrayInputStream;

public class ExHtmlLoadOptions extends ApiExampleBase {
    @Test
    public void supportVml() throws Exception {
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
    public void webRequestTimeoutDefaultValue() {
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Assert.assertEquals(loadOptions.getWebRequestTimeout(), 100000);
    }

    @Test
    public void getSelectAsSdt() throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.PreferredControlType
        //ExSummary:Shows how to set preffered type of document nodes that will represent imported <input> and <select> elements.
        final String html = "\r\n<html>\r\n<select name='ComboBox' size='1'>\r\n"
                + "<option value='val1'>item1</option>\r\n<option value='val2'></option>\r\n</select>\r\n</html>\r\n";

        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
        htmlLoadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG);

        Document doc = new Document(new ByteArrayInputStream(html.getBytes("UTF-8")), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag tag = (StructuredDocumentTag) nodes.get(0);
        //ExEnd

        Assert.assertEquals(tag.getListItems().getCount(), 2);

        Assert.assertEquals(tag.getListItems().get(0).getValue(), "val1");
        Assert.assertEquals(tag.getListItems().get(1).getValue(), "val2");
    }

    @Test
    public void getInputAsFormField() throws Exception {
        final String html = "\r\n<html>\r\n<input type='text' value='Input value text' />\r\n</html>\r\n";

        // By default "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField"
        // So, we do not set this value
        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

        Document doc = new Document(new ByteArrayInputStream(html.getBytes("UTF-8")), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.FORM_FIELD, true);

        Assert.assertEquals(nodes.getCount(), 1);

        FormField formField = (FormField) nodes.get(0);
        Assert.assertEquals(formField.getResult(), "Input value text");
    }
}
