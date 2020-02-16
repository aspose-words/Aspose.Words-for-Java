// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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


@Test
public class ExMetered extends ApiExampleBase
{
    @Test
    public void testMeteredUsage()
    {
        Assert.<IllegalStateException>Throws(meteredUsage);
    }

    @Test (enabled = false)
    public void meteredUsage() throws Exception
    {
        //ExStart
        //ExFor:Metered
        //ExFor:Metered.#ctor
        //ExFor:Metered.GetConsumptionCredit
        //ExFor:Metered.GetConsumptionQuantity
        //ExFor:Metered.SetMeteredKey(String, String)
        //ExSummary:Shows how to activate a Metered license and track credit/consumption.
        // Set a public and private key for a new Metered instance
        Metered metered = new Metered();
        metered.setMeteredKey("MyPublicKey", "MyPrivateKey");
        
        // Print credit/usage 
        System.out.println("Credit before operation: {Metered.GetConsumptionCredit()}");
        System.out.println("Consumption quantity before operation: {Metered.GetConsumptionQuantity()}");

        // Do something
        Document doc = new Document(getMyDir() + "Document.docx");

        // Print credit/usage to see how much was spent
        System.out.println("Credit after operation: {Metered.GetConsumptionCredit()}");
        System.out.println("Consumption quantity after operation: {Metered.GetConsumptionQuantity()}");
        //ExEnd
    }
}
