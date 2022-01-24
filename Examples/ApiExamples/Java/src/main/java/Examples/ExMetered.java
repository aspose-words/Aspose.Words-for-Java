package Examples;

// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        Assert.assertThrows(IllegalStateException.class, () -> usage());
    }

    @Test(enabled = false)
    public void usage() throws Exception {
        //ExStart
        //ExFor:Metered
        //ExFor:Metered.#ctor
        //ExFor:Metered.GetConsumptionCredit
        //ExFor:Metered.GetConsumptionQuantity
        //ExFor:Metered.SetMeteredKey(String, String)
        //ExSummary:Shows how to activate a Metered license and track credit/consumption.
        // Create a new Metered license, and then print its usage statistics.
        Metered metered = new Metered();
        metered.setMeteredKey("MyPublicKey", "MyPrivateKey");

        System.out.println("Credit before operation: {Metered.GetConsumptionCredit()}");
        System.out.println("Consumption quantity before operation: {Metered.GetConsumptionQuantity()}");

        // Operate using Aspose.Words, and then print our metered stats again to see how much we spent.
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Metered.Usage.pdf");

        System.out.println(MessageFormat.format("Credit after operation: {0}", Metered.getConsumptionCredit()));
        System.out.println(MessageFormat.format("Consumption quantity after operation: {0}", Metered.getConsumptionQuantity()));
        //ExEnd
    }
}
