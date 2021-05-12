package ru.curs.xylophone;

import com.github.miachm.sods.SpreadSheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.approvaltests.Approvals;
import org.approvaltests.approvers.FileApprover;
import org.approvaltests.core.Options;
import org.approvaltests.writers.ApprovalBinaryFileWriter;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.lambda.functions.Function2;

import java.io.*;

import static bad.robot.excel.matchers.Matchers.sameWorkbook;

public class TestWriterODS {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    /**
     * Функция, сравнивающая файлы с содержимым-таблицами.
     */
    public Function2<File, File, Boolean> compareSpreadsheetFiles = (actualFile, expectedFile) -> {
        try {
            SpreadSheet actual = new SpreadSheet(actualFile);
            SpreadSheet expected = new SpreadSheet(expectedFile);
            return actual.equals(expected);
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    };

    @Test
    public void WriterODS() throws XylophoneError {
        InputStream descrStream = TestReader.class
                .getResourceAsStream("testsaxdescriptor3.json");
        InputStream dataStream = TestReader.class
                .getResourceAsStream("testdata.xml");
        InputStream templateStream = TestReader.class
                .getResourceAsStream("template.ods");

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        XML2Spreadsheet.process(dataStream, descrStream, templateStream,
                OutputType.ODS, false, false, bos);

        byte[] writtenData = bos.toByteArray();

        Options options = new Options();
        Approvals.verify(
                new FileApprover(
                        new ApprovalBinaryFileWriter(new ByteArrayInputStream(writtenData),
                                "ods"),
                        options.forFile().getNamer(),
                        compareSpreadsheetFiles
                ),
                options
        );
    }

}
