package ru.curs.xylophone;

import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.Assert.assertTrue;

public class WriterODSTest {

    @Test
    public void WriterODS() throws XML2SpreadSheetError, Exception {
        InputStream descrStream = TestReader.class
                .getResourceAsStream("testdescriptor3.xml");
        InputStream dataStream = TestReader.class
                .getResourceAsStream("testdata.xml");
        InputStream templateStream = TestReader.class
                .getResourceAsStream("template.ods");

        new ODSReportWriter(templateStream, templateStream);

        assertTrue(true);
    }

}
