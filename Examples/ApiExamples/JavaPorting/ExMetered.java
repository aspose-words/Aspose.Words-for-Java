// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import org.testng.Assert;
import com.aspose.words.Metered;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Document;
import java.lang.Thread;


@Test
public class ExMetered extends ApiExampleBase
{
    @Test
    public void testMeteredUsage()
    {
        Assert.<IllegalStateException>Throws(usage);
    }

    @Test (enabled = false)
    public void usage() throws Exception
    {
        //ExStart
        //ExFor:Metered
        //ExFor:Metered.#ctor
        //ExFor:Metered.GetConsumptionCredit
        //ExFor:Metered.GetConsumptionQuantity
        //ExFor:Metered.SetMeteredKey(String, String)
        //ExFor:Metered.IsMeteredLicensed
        //ExFor:Metered.GetProductName
        //ExSummary:Shows how to activate a Metered license and track credit/consumption.
        // Create a new Metered license, and then print its usage statistics.
        Metered metered = new Metered();
        metered.setMeteredKey("MyPublicKey", "MyPrivateKey");

        System.out.println("Is metered license accepted: {Metered.IsMeteredLicensed()}");
        System.out.println("Product name: {metered.GetProductName()}");
        System.out.println("Credit before operation: {Metered.GetConsumptionCredit()}");
        System.out.println("Consumption quantity before operation: {Metered.GetConsumptionQuantity()}");

        // Operate using Aspose.Words, and then print our metered stats again to see how much we spent.
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Metered.Usage.pdf");

        // Aspose Metered Licensing mechanism does not send the usage data to purchase server every time,
        // you need to use waiting.
        Thread.sleep(10000);

        System.out.println("Credit after operation: {Metered.GetConsumptionCredit()}");
        System.out.println("Consumption quantity after operation: {Metered.GetConsumptionQuantity()}");
        //ExEnd
    }
}
