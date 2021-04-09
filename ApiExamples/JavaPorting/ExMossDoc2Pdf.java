// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import com.aspose.ms.System.IO.StreamWriter;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.Environment;
import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;


/// <summary>
/// DOC2PDF document converter for SharePoint.
/// Uses Aspose.Words to perform the conversion.
/// </summary>
public /*static*/ class ExMossDoc2Pdf
{
	/* Simulation of static class by using private constructor */
	private ExMossDoc2Pdf()
	{}

    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    @STAThread
    public static void mossDoc2Pdf(String[] args) throws Exception
    {
        // Although SharePoint passes "-log <filename>" to us and we are
        // supposed to log there, we will use our hardcoded path to the log file for the sake of simplicity.
        // 
        // Make sure there are permissions to write into this folder.
        // The document converter will be called under the document 
        // conversion account (not sure what name), so for testing purposes, 
        // I would give the Users group write permissions into this folder.
        gLog = new StreamWriter("C:\\Aspose2Pdf\\log.txt", true);

        try
        {
            gLog.writeLine(new Date().toString(msCultureInfo.getInvariantCulture()) + " Started");
            gLog.writeLine(Environment.getCommandLine());

            parseCommandLine(args);

            // Uncomment the code below when you have purchased a license for Aspose.Words.
            //
            // You need to deploy the license in the same folder as your 
            // executable, alternatively you can add the license file as an 
            // embedded resource to your project.
            //
            // Set license for Aspose.Words.
            // Aspose.Words.License wordsLicense = new Aspose.Words.License();
            // wordsLicense.SetLicense("Aspose.Total.lic");

            convertDoc2Pdf(gInFileName, gOutFileName);
        }
        catch (Exception e)
        {
            gLog.writeLine(e.getMessage());
            Environment.setExitCode(100);
        }
        finally
        {
            gLog.close();
        }
    }

    private static void parseCommandLine(String[] args)
    {
        int i = 0;
        while (i < args.length)
        {
            String s = args[i];
            switch (gStringSwitchMap.of(s.toLowerCase()))
            {
                case /*"-in"*/0:
                    i++;
                    gInFileName = args[i];
                    break;
                case /*"-out"*/1:
                    i++;
                    gOutFileName = args[i];
                    break;
                case /*"-config"*/2:
                    // Skip the name of the config file and do nothing.
                    i++;
                    break;
                case /*"-log"*/3:
                    // Skip the name of the log file and do nothing.
                    i++;
                    break;
                default:
                    throw new Exception("Unknown command line argument: " + s);
            }

            i++;
        }
    }

    private static void convertDoc2Pdf(String inFileName, String outFileName) throws Exception
    {
        // You can load not only DOC here, but any format supported by
        // Aspose.Words: DOC, DOCX, RTF, WordML, HTML, MHTML, ODT etc.
        Document doc = new Document(inFileName);

        doc.save(outFileName, new PdfSaveOptions());
    }

    private static String gInFileName;
    private static String gOutFileName;
    private static StreamWriter gLog;

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"-in",
		"-out",
		"-config",
		"-log"
	);

}
