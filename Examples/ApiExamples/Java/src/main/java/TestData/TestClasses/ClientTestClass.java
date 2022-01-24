package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

public class ClientTestClass {
    private String mName;
    private String mCountry;
    private String mLocalAddress;

    public ClientTestClass() {
    }

    public ClientTestClass(final String name, final String localAddress) {
        setName(name);
        setLocalAddress(localAddress);
    }

    public ClientTestClass(final String name, final String country, final String localAddress) {
        setName(name);
        setCountry(country);
        setLocalAddress(localAddress);
    }

    public void setName(final String value) {
        mName = value;
    }

    public void setCountry(final String value) {
        mCountry = value;
    }

    public void setLocalAddress(final String value) {
        mLocalAddress = value;
    }

    public String getName() {
        return mName;
    }

    public String getCountry() {
        return mCountry;
    }

    public String getLocalAddress() {
        return mLocalAddress;
    }
}
