package ru.curs.xylophone;

import org.junit.Test;

public class TestWriterODS extends FullApprovalsTester {

    @Test
    public void WriterODS() throws XylophoneError {
        approvalTest("descriptor_testdata/overall/testsaxdescriptor3.yaml",
                "testdata.xml", "template.ods",
                OutputType.ODS, false);
    }

}
