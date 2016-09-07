package com.aspose.words.examples.loading_saving;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import com.aspose.email.MailAddress;
import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class ConvertADocumentToMHTMLAndEmail {

	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(ConvertADocumentToMHTMLAndEmail.class) + "LoadingSavingAndConverting/";
		
		// Load the document into Aspose.Words.
		String srcFileName = dataDir + "Document.doc";
		Document doc = new Document(srcFileName);

		// Save to an output stream in MHTML format.
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		doc.save(outputStream, SaveFormat.MHTML);

		// Load the MHTML stream back into an input stream for use with Aspose.Email.
		ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());

		// Create an Aspose.Email MIME email message from the stream.
		MailMessage message = MailMessage.load(inputStream);
		message.setFrom(new MailAddress("your_from@email.com"));
		message.getTo().add("your_to@email.com");
		message.setSubject("Aspose.Words + Aspose.Email MHTML Test Message");

		// Save the message in Outlook MSG format.
		message.save(dataDir + "Message Out.msg", SaveOptions.getDefaultMsg());
	}
}
