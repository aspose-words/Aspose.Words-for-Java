package Examples;

// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.Metered;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;

@Test
public class ExMetered extends ApiExampleBase {
    @Test
    public void testMeteredUsage() {
        Assert.assertThrows(IllegalStateException.class, () -> meteredUsage());
    }

    @Test(enabled = false)
    public void meteredUsage() throws Exception {
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
        System.out.println(MessageFormat.format("Credit before operation: {0}", Metered.getConsumptionCredit()));
        System.out.println(MessageFormat.format("Consumption quantity before operation: {0}", Metered.getConsumptionQuantity()));

        // Do something
        Document doc = new Document(getMyDir() + "Document.docx");

        // Print credit/usage to see how much was spent
        System.out.println(MessageFormat.format("Credit after operation: {0}", Metered.getConsumptionCredit()));
        System.out.println(MessageFormat.format("Consumption quantity after operation: {0}", Metered.getConsumptionQuantity()));
        //ExEnd
    }
}
