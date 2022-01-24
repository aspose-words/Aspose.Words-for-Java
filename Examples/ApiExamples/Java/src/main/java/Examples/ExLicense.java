package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.License;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExLicense extends ApiExampleBase {
    @Test
    public void licenseFromFileNoPath() throws Exception {
        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExSummary:Shows how initialize a license for Aspose.Words using a license file in the local file system.
        // Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
        Path licenseFileName = Paths.get(getLicenseDir(), "Aspose.Words.Java.lic");

        License license = new License();
        license.setLicense(licenseFileName.toString());

        // Create a copy of our license file in the binaries folder of our application.
        Path licenseCopyFileName = Paths.get(System.getProperty("user.dir"), "Aspose.Words.Java.lic");
        FileUtils.copyFile(new File(licenseFileName.toString()), new File(licenseCopyFileName.toString()));

        // If we pass a file's name without a path,
        // the SetLicense will search several local file system locations for this file.
        // One of those locations will be the "bin" folder, which contains a copy of our license file.
        license.setLicense("Aspose.Words.Java.lic");
        //ExEnd

        license.setLicense("");
        Files.deleteIfExists(licenseCopyFileName);
    }

    @Test
    public void licenseFromStream() throws Exception {
        //ExStart
        //ExFor:License.SetLicense(Stream)
        //ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
        // Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
        try (FileInputStream myStream = new FileInputStream(getLicenseDir() + "Aspose.Words.Java.lic")) {
            License license = new License();
            license.setLicense(myStream);
        }
        //ExEnd
    }
}

