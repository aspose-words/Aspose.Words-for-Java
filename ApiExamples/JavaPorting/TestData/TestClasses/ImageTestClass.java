package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.IO.Stream;


public class ImageTestClass
{
    public SKBitmap getImage() { return mImage; }; public void setImage(SKBitmap value) { mImage = value; };

    private SKBitmap mImage;
    public Stream getImageStream() { return mImageStream; }; public void setImageStream(Stream value) { mImageStream = value; };

    private Stream mImageStream;
    public byte[] getImageBytes() { return mImageBytes; }; public void setImageBytes(byte[] value) { mImageBytes = value; };

    private byte[] mImageBytes;
    public String getImageUri() { return mImageUri; }; public void setImageUri(String value) { mImageUri = value; };

    private String mImageUri;

    public ImageTestClass(SKBitmap image, Stream imageStream, byte[] imageBytes, String imageUri)
    {
        this.setImage(image);
        this.setImageStream(imageStream);
        this.setImageBytes(imageBytes);
        this.setImageUri(imageUri);
    }        
}
