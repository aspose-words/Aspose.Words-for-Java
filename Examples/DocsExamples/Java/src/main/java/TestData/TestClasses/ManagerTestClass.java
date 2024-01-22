package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import DocsExamples.DocsExamplesBase;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class ManagerTestClass extends DocsExamplesBase {
    private String mName;
    private int mAge;
    private byte[] mPhoto;
    private ArrayList<ContractTestClass> mContracts;

    public ManagerTestClass() {
    }

    public ManagerTestClass(final String name, final int age, final ArrayList<ContractTestClass> contracts) throws IOException {
        setName(name);
        setAge(age);
        setPhoto(FileUtils.readFileToByteArray(new File(getImagesDir() + "Logo.jpg")));
        setContracts(contracts);
    }

    public void setName(final String value) {
        mName = value;
    }

    public void setAge(final int value) {
        mAge = value;
    }

    public void setPhoto(byte[] value) { mPhoto = value; }

    public void setContracts(final ArrayList<ContractTestClass> value) {
        mContracts = value;
    }

    public String getName() {
        return mName;
    }

    public int getAge() {
        return mAge;
    }

    public byte[] getPhoto() { return mPhoto; }

    public ArrayList<ContractTestClass> getContracts() {
        return mContracts;
    }

    public int getContractsSum() {
        int contractsSum = 0;

        for (ContractTestClass contract : getContracts()) {
            contractsSum = (int) (contractsSum + contract.getPrice());
        }

        return contractsSum;
    }
}
