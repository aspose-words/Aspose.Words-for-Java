//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package loadingandsaving.loadingandsavinghtml.savemhtmlandemail.java;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;
import com.aspose.network.*;


public class SaveMhtmlAndEmail
{
    public static void main(String[] args) throws Exception
    {
        // Find out the directory where the file we want to send is stored.
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = "src/loadingandsaving/loadingandsavinghtml/savemhtmlandemail/data/";

        //ExStart
        //ExId:SaveMhtmlAndEmail
        //ExSummary:Shows how to save any document from Aspose.Words as MHTML and create a Outlook MSG file from it using Aspose.Network.
        // Load the document into Aspose.Words.
        String srcFileName = dataDir + "DinnerInvitationDemo.doc";
        Document doc = new Document(srcFileName);

        // Save to an output stream in MHTML format.
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        doc.save(outputStream, SaveFormat.MHTML);

        // Load the MHTML stream back into an input stream for use with Aspose.Network.
        ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());

        // Create an Aspose.Network MIME email message from the stream.
        MailMessage message = MailMessage.load(inputStream, MessageFormat.getMht());
        message.setFrom(new MailAddress("your_from@email.com"));
        message.getTo().add("your_to@email.com");
        message.setSubject("Aspose.Words + Aspose.Network MHTML Test Message");

        // Save the message in Outlook msg format.
        message.save(dataDir + "Message Out.msg", MailMessageSaveType.getOutlookMessageFormat());
        //ExEnd
    }
}