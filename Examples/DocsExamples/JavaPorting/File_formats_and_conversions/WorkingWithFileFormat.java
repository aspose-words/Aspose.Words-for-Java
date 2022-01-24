package DocsExamples.File_Formats_and_Conversions;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.msConsole;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.LoadFormat;
import com.aspose.ms.System.IO.File;


public class WorkingWithFileFormat extends DocsExamplesBase
{
    @Test
    public void detectFileFormat() throws Exception
    {
        //ExStart:CheckFormatCompatibility
        String supportedDir = getArtifactsDir() + "Supported";
        String unknownDir = getArtifactsDir() + "Unknown";
        String encryptedDir = getArtifactsDir() + "Encrypted";
        String pre97Dir = getArtifactsDir() + "Pre97";

        // Create the directories if they do not already exist.
        if (Directory.exists(supportedDir) == false)
            Directory.createDirectory(supportedDir);
        if (Directory.exists(unknownDir) == false)
            Directory.createDirectory(unknownDir);
        if (Directory.exists(encryptedDir) == false)
            Directory.createDirectory(encryptedDir);
        if (Directory.exists(pre97Dir) == false)
            Directory.createDirectory(pre97Dir);

        //ExStart:GetListOfFilesInFolder
        Iterable<String> fileList = Directory.getFiles(getMyDir()).Where(name => !name.EndsWith("Corrupted document.docx"));
        //ExEnd:GetListOfFilesInFolder
        for (String fileName : fileList)
        {
            String nameOnly = Path.getFileName(fileName);
            
            msConsole.write(nameOnly);
            //ExStart:DetectFileFormat
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);

            // Display the document type
            switch (info.getLoadFormat())
            {
                case LoadFormat.DOC:
                    System.out.println("\tMicrosoft Word 97-2003 document.");
                    break;
                case LoadFormat.DOT:
                    System.out.println("\tMicrosoft Word 97-2003 template.");
                    break;
                case LoadFormat.DOCX:
                    System.out.println("\tOffice Open XML WordprocessingML Macro-Free Document.");
                    break;
                case LoadFormat.DOCM:
                    System.out.println("\tOffice Open XML WordprocessingML Macro-Enabled Document.");
                    break;
                case LoadFormat.DOTX:
                    System.out.println("\tOffice Open XML WordprocessingML Macro-Free Template.");
                    break;
                case LoadFormat.DOTM:
                    System.out.println("\tOffice Open XML WordprocessingML Macro-Enabled Template.");
                    break;
                case LoadFormat.FLAT_OPC:
                    System.out.println("\tFlat OPC document.");
                    break;
                case LoadFormat.RTF:
                    System.out.println("\tRTF format.");
                    break;
                case LoadFormat.WORD_ML:
                    System.out.println("\tMicrosoft Word 2003 WordprocessingML format.");
                    break;
                case LoadFormat.HTML:
                    System.out.println("\tHTML format.");
                    break;
                case LoadFormat.MHTML:
                    System.out.println("\tMHTML (Web archive) format.");
                    break;
                case LoadFormat.ODT:
                    System.out.println("\tOpenDocument Text.");
                    break;
                case LoadFormat.OTT:
                    System.out.println("\tOpenDocument Text Template.");
                    break;
                case LoadFormat.DOC_PRE_WORD_60:
                    System.out.println("\tMS Word 6 or Word 95 format.");
                    break;
                case LoadFormat.UNKNOWN:
                    System.out.println("\tUnknown format.");
                    break;
            }
            //ExEnd:DetectFileFormat

            if (info.isEncrypted())
            {
                System.out.println("\tAn encrypted document.");
                File.copy(fileName, Path.combine(encryptedDir, nameOnly), true);
            }
            else
            {
                switch (info.getLoadFormat())
                {
                    case LoadFormat.DOC_PRE_WORD_60:
                        File.copy(fileName, Path.combine(pre97Dir, nameOnly), true);
                        break;
                    case LoadFormat.UNKNOWN:
                        File.copy(fileName, Path.combine(unknownDir, nameOnly), true);
                        break;
                    default:
                        File.copy(fileName, Path.combine(supportedDir, nameOnly), true);
                        break;
                }
            }
        }
        //ExEnd:CheckFormatCompatibility
    }

    @Test
    public void detectDocumentSignatures() throws Exception
    {
        //ExStart:DetectDocumentSignatures
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Digitally signed.docx");

        if (info.hasDigitalSignature())
        {
            System.out.println("Document {Path.GetFileName(MyDir + ");
        }
        //ExEnd:DetectDocumentSignatures            
    }

    @Test
    public void verifyEncryptedDocument() throws Exception
    {
        //ExStart:VerifyEncryptedDocument
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Encrypted.docx");
        msConsole.writeLine(info.isEncrypted());
        //ExEnd:VerifyEncryptedDocument
    }
}
