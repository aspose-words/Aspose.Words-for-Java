package DocsExamples.Mail_Merge_And_Reporting;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.net.System.Data.DataSet;
import org.testng.annotations.Test;

@Test
public class WorkingWithXmlData extends DocsExamplesBase
{
    @Test
    public void xmlMailMerge() throws Exception
    {
        //ExStart:XmlMailMerge
        DataSet customersDs = new DataSet();
        customersDs.readXml(getMyDir() + "Mail merge data - Customers.xml");

        Document doc = new Document(getMyDir() + "Mail merge destinations - Registration complete.docx");
        doc.getMailMerge().execute(customersDs.getTables().get("Customer"));

        doc.save(getArtifactsDir() + "WorkingWithXmlData.XmlMailMerge.docx");
        //ExEnd:XmlMailMerge
    }

    @Test
    public void nestedMailMerge() throws Exception
    {
        //ExStart:NestedMailMerge
        // The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
        DataSet pizzaDs = new DataSet();
        pizzaDs.readXml(getMyDir() + "Mail merge data - Orders.xml");
        
        Document doc = new Document(getMyDir() + "Mail merge destinations - Invoice.docx");

        // Trim trailing and leading whitespaces mail merge values.
        doc.getMailMerge().setTrimWhitespaces(false);

        doc.getMailMerge().executeWithRegions(pizzaDs);

        doc.save(getArtifactsDir() + "WorkingWithXmlData.NestedMailMerge.docx");
        //ExEnd:NestedMailMerge
    }

    @Test
    public void mustacheSyntaxUsingDataSet() throws Exception
    {
        //ExStart:MailMergeUsingMustacheSyntax
        DataSet ds = new DataSet();
        ds.readXml(getMyDir() + "Mail merge data - Vendors.xml");

        Document doc = new Document(getMyDir() + "Mail merge destinations - Vendor.docx");

        doc.getMailMerge().setUseNonMergeFields(true);

        doc.getMailMerge().executeWithRegions(ds);
        
        doc.save(getArtifactsDir() + "WorkingWithXmlData.MustacheSyntaxUsingDataSet.docx");
        //ExEnd:MailMergeUsingMustacheSyntax
    }
}
