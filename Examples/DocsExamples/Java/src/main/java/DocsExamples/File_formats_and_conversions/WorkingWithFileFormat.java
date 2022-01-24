package DocsExamples.File_formats_and_conversions;

import DocsExamples.DocsExamplesBase;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.LoadFormat;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Test
public class WorkingWithFileFormat extends DocsExamplesBase
{
    @Test
    public void detectFileFormat() throws Exception {
        //ExStart:CheckFormatCompatibility
        File supportedDir = new File(getArtifactsDir() + "Supported");
        File unknownDir = new File(getArtifactsDir() + "Unknown");
        File encryptedDir = new File(getArtifactsDir() + "Encrypted");
        File pre97Dir = new File(getArtifactsDir() + "Pre97");

        // Create the directories if they do not already exist.
        if (supportedDir.exists() == false)
            supportedDir.mkdir();
        if (unknownDir.exists() == false)
            unknownDir.mkdir();
        if (encryptedDir.exists() == false)
            encryptedDir.mkdir();
        if (pre97Dir.exists() == false)
            pre97Dir.mkdir();

        //ExStart:GetListOfFilesInFolder
        Set<String> listFiles = Stream.of(new File(getMyDir()).listFiles())
                .filter(file -> !file.getName().endsWith("Corrupted document.docx") && !Files.isDirectory(file.toPath()))
                .map(File::getPath)
                .collect(Collectors.toSet());
        //ExEnd:GetListOfFilesInFolder
        for (String fileName : listFiles) {
            String nameOnly = Paths.get(fileName).getFileName().toString();

            System.out.println(nameOnly);
            //ExStart:DetectFileFormat
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);

            // Display the document type
            switch (info.getLoadFormat()) {
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

            if (info.isEncrypted()) {
                System.out.println("\tAn encrypted document.");
                FileUtils.copyFile(new File(fileName), new File(encryptedDir, nameOnly));
            } else {
                switch (info.getLoadFormat()) {
                    case LoadFormat.DOC_PRE_WORD_60:
                        FileUtils.copyFile(new File(fileName), new File(pre97Dir, nameOnly));
                        break;
                    case LoadFormat.UNKNOWN:
                        FileUtils.copyFile(new File(fileName), new File(unknownDir, nameOnly));
                        break;
                    default:
                        FileUtils.copyFile(new File(fileName), new File(supportedDir, nameOnly));
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
        System.out.println(info.isEncrypted());
        //ExEnd:VerifyEncryptedDocument
    }
}
