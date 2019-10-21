package com.aspose.words.examples.quickstart;

import com.aspose.words.*;

public class ApplyLicense {
    public static void main(String[] args) throws Exception {
        //ExStart:ApplyLicense
        // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        // You can also use the additional overload to load a license from a stream, this is useful for instance when the
        // license is stored as an embedded resource
        try {
            License license = new License();
            license.setLicense("Aspose.Words.lic");
            System.out.println("License set successfully.");
        } catch (Exception e) {
            // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            System.out.println("There was an error setting the license: " + e.getMessage());
        }
        //ExEnd:ApplyLicense
    }

    public static void ApplyMeteredLicense() throws Exception {
        //ExStart:ApplyMeteredLicense
        String publicKey = "";
        String privateKey = "";

        Metered m = new Metered();
        m.setMeteredKey(publicKey, privateKey);

        // Optionally, the following two lines returns true if a valid license has been applied;
        // false if the component is running in evaluation mode.
        License lic = new License();

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello World!");

        System.out.println(doc.toString(SaveFormat.TEXT));
        //ExEnd:ApplyMeteredLicense
    }
}