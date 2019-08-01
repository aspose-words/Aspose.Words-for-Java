package TestData.TestBuilders;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import Examples.ApiExampleBase;
import TestData.TestClasses.ImageTestClass;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import javax.imageio.ImageIO;

public class ImageTestBuilder extends ApiExampleBase {
    private BufferedImage mImage;
    private FileInputStream mImageStream;
    private byte[] mImageBytes;
    private String mImageUri;

    public ImageTestBuilder() throws Exception {

        mImage = ImageIO.read(new File(getImageDir() + "Watermark.png"));

        mImageStream = null;
        mImageBytes = new byte[0];
        mImageUri = "";
    }

    public ImageTestBuilder withImage(final BufferedImage image) {
        mImage = image;
        return this;
    }

    public ImageTestBuilder withImageStream(final FileInputStream imageStream) {
        mImageStream = imageStream;
        return this;
    }

    public ImageTestBuilder withImageBytes(final byte[] imageBytes) {
        mImageBytes = imageBytes;
        return this;
    }

    public ImageTestBuilder withImageUri(final String imageUri) {
        mImageUri = imageUri;
        return this;
    }

    public ImageTestClass build() {
        return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageUri);
    }
}
