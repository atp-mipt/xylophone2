package ru.curs.xylophone;

import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.Assert.assertTrue;

public class WriterODSTest {

    @Test
    public void WriterODS() throws IOException, XML2SpreadSheetError {
        InputStream templateStream = TestReader.class
                .getResourceAsStream("template.ods");

        InputStream templateCopyStream = TestReader.class
                .getResourceAsStream("template.ods");

        new ODSReportWriter(templateStream, templateCopyStream);

        assertTrue(true);
    }

}
