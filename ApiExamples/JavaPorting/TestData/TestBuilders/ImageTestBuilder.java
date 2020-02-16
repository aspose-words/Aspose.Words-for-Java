package ApiExamples.TestData.TestBuilders;

// ********* THIS FILE IS AUTO PORTED *********

import ApiExamples.ApiExampleBase;
import com.aspose.ms.System.IO.Stream;
import ApiExamples.TestData.TestClasses.ImageTestClass;


public class ImageTestBuilder extends ApiExampleBase
{
    private SKBitmap mImage;
    private Stream mImageStream;
    private byte[] mImageBytes;
    private String mImageUri;

    public ImageTestBuilder()
    {
        this.mImage = SKBitmap.Decode(getImageDir() + "Transparent background logo.png");
        mImageStream = Stream.Null;
        mImageBytes = new byte[0];
        mImageUri = "";
    }

    public ImageTestBuilder withImage(SKBitmap image)
    {
        this.mImage = image;
        return this;
    }

    public ImageTestBuilder withImageStream(Stream imageStream)
    {
        mImageStream = imageStream;
        return this;
    }

    public ImageTestBuilder withImageBytes(byte[] imageBytes)
    {
        mImageBytes = imageBytes;
        return this;
    }

    public ImageTestBuilder withImageUri(String imageUri)
    {
        mImageUri = imageUri;
        return this;
    }

    public ImageTestClass build()
    {
        return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageUri);
    }
}
