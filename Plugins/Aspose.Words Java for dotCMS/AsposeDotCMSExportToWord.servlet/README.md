# Export to Word Plugin for dotCMS

Export to Word Plugin for dotCMS allow users to export online content into Word Processing document using [Aspose.Words for Java](https://products.aspose.com/words/java). It dynamically exports the content of the Web page to a Word  Processing document and then automatically downloads the file to  the disk location selected by the user in just couple of seconds. The generated Word processing document can then be opened using any Word Processing Application such as Microsoft Word or Apache OpenOffice etc.

![Export to Word Image 1](http://i.imgur.com/aqlE55d.png|align=center,border=1!)

## Installing

Once downloaded, please follow these steps to install the plugin into your dotCMS website:

1. Log into your site as either admin or another super-user level account.
2. Go to the dotCMS Dynamic Plugin portlet under the System tab and click on the "Upload Plugin" button and then choose the AsposeDotCMSExportToWord JAR file.
![Export to Word Image 2](http://i.imgur.com/7dhmMs3.png|align=center,border=1!)  
OR  
Copy the AsposeDotCMSExportToWord JAR file inside the Felix OSGI container (dotCMS/felix/load).  
3. Please add the following 2 exported packages either by changing the file: dotCMS/WEB-INF/felix/osgi-extra.conf or using the dotCMS UI (System \-> Dynamic Plugins \-> Exported Packages).
  1. javax.xml.stream
  2. javax.xml.namespace

## Using

After you have installed the Export to Word OSGI plugin, it is really simple to start using it on your website. Please follow these simple steps to get started:

1. Make sure you are logged-in to dotCMS with a Host or Admin level account.
2. Navigate to the page whose content you want to export to a Word Processing document.
3. Add following HTML code in your page content.
```
<form action="/app/exporttoword" method="POST">
    <input type="hidden" name="page_url" value=$dotPageContent.url />
    <input type="submit" value="Export to Word" style="float: right;" />
</form>
```
This will add **Export to Word** button on the page and clicking a button will dynamically exports the content of the page into a Word Processing document.

## How to apply Aspose License?

This Plugin uses an evaluation version of Aspose.Words. Once you are happy with your evaluation, you can purchase a license at the [Aspose website](http://www.aspose.com/purchase/default.aspx).  
To remove evaluation message and feature limitations, product license should be applied. You will receive a license file after you have purchased the product. Please follow the steps below to apply the license

- Make sure the license file is named as **Aspose.Words.Java.lic.**
- Place **Aspose.Words.Java.lic** file in the folder that contains the Aspose.Words.jar
- Use following code for activating the license:

```
License license = new License();
license.setLicense("Aspose.Words.Java.lic");
```
Please check this [article](https://docs.aspose.com/display/wordsjava/Applying+a+License) for further details.

## Supported Platforms

In order to deploy Export to Word OSGI/Dynamic Plugin you need to have the following requirements met:
- dotCMS 2.2 +

Please feel free to contact us if you wish to install this plugin on older versions of dotCMS.

## Contact

Your feedback is very important to us. Please email us all your queries and feedback at marketplace@aspose.com.
