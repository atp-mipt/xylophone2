package ru.curs.xylophone;

import org.junit.Test;

import java.io.*;
import java.util.stream.Collectors;

import static junit.framework.TestCase.fail;
import static org.junit.Assert.assertTrue;

public class TestWriterODS {

    @Test
    public void WriterODS() throws XylophoneError, FileNotFoundException {
        InputStream descrStream = TestReader.class
                .getResourceAsStream("testsaxdescriptor3.json");
        InputStream dataStream = TestReader.class
                .getResourceAsStream("testdata.xml");
        InputStream templateStream = TestReader.class
                .getResourceAsStream("template.xls");

        File resultFile = new File("result.ods");
        OutputStream resultStream = new FileOutputStream(resultFile);

        XML2Spreadsheet.process(dataStream, descrStream, templateStream,
                OutputType.ODS, true, false, resultStream);

        fail();
    }

}
