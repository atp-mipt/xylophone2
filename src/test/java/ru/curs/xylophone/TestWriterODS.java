package ru.curs.xylophone;

import org.junit.Test;

public class TestWriterODS extends FullApprovalsTester {

    @Test
    public void WriterODS() throws XylophoneError {
        approvalTest("testsaxdescriptor3.json", "testdata.xml", "template.ods",
                OutputType.ODS, false);
    }

}
