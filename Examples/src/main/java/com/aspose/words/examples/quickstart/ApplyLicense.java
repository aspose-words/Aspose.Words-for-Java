/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package com.aspose.words.examples.quickstart;

import com.aspose.words.License;

public class ApplyLicense
{
    public static void main(String[] args) throws Exception
    {
        // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        // You can also use the additional overload to load a license from a stream, this is useful for instance when the
        // license is stored as an embedded resource
        try
        {
            License license = new License();
            license.setLicense("Aspose.Words.lic");
            System.out.println("License set successfully.");
        }
        catch (Exception e)
        {
            // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            System.out.println("There was an error setting the license: " + e.getMessage());
        }
    }
}