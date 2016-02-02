package com.aspose.words.examples.linq;
import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import  java.util.*;
public class HelloWorld {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception
    {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(HelloWorld.class);

        String fileName = "HelloWorld.doc";
        // Load the template document.
        Document doc = new Document(dataDir + fileName);

        // Create an instance of sender class to set it's properties.
        Sender sender = new Sender();
        sender.setName("LINQ Reporting Engine");
        sender.setMessage("Hello World");

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Execute the build report.
        engine.buildReport(doc, sender, "sender");

        dataDir = dataDir + Utils.GetOutputFilePath(fileName);

        // Save the finished document to disk.
        doc.save(dataDir);

        System.out.println("\nTemplate document is populated with the data about the sender.\nFile saved at " + dataDir);

    }
}
