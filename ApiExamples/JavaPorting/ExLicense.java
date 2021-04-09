// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.License;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.Stream;
import java.io.FileInputStream;


@Test
class ExLicense !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void licenseFromFileNoPath() throws Exception
    {
        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExSummary:Shows how initialize a license for Aspose.Words using a license file in the local file system.
        // Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
        String licenseFileName = Path.combine(getLicenseDir(), "Aspose.Words.NET.lic");

        License license = new License();
        license.setLicense(licenseFileName);

        // Create a copy of our license file in the binaries folder of our application.
        String licenseCopyFileName = Path.combine(getAssemblyDir(), "Aspose.Words.NET.lic");
        File.copy(licenseFileName, licenseCopyFileName);

        // If we pass a file's name without a path,
        // the SetLicense will search several local file system locations for this file.
        // One of those locations will be the "bin" folder, which contains a copy of our license file.
        license.setLicense("Aspose.Words.NET.lic");
        //ExEnd

        license.setLicense("");
        File.delete(licenseCopyFileName);
    }

    @Test
    public void licenseFromStream() throws Exception
    {
        //ExStart
        //ExFor:License.SetLicense(Stream)
        //ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
        // Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
        Stream myStream = new FileInputStream(Path.combine(getLicenseDir(), "Aspose.Words.NET.lic"));
        try /*JAVA: was using*/
        {
            License license = new License();
            license.setLicenseInternal(myStream);
        }
        finally { if (myStream != null) myStream.close(); }
        //ExEnd
    }
}

