package ru.curs.xylophone;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class DescriptorOutput extends DescriptorOutputBase {
    private static final Pattern RANGE = Pattern
            .compile("(-?[0-9]+):(-?[0-9]+)");

    private final String worksheet;
    private final RangeAddress range;
    private final String sourceSheet;
    private final int startRepeatingColumn;
    private final int endRepeatingColumn;
    private final int startRepeatingRow;
    private final int endRepeatingRow;
    private final boolean pageBreak;

    @JsonCreator(mode = JsonCreator.Mode.PROPERTIES)
    DescriptorOutput(
            @JsonProperty("worksheet")      String worksheet,
            @JsonProperty("range")          String range,
            @JsonProperty("sourcesheet")    String sourceSheet,
            @JsonProperty("repeatingcols")  String repeatingCols,
            @JsonProperty("repeatingraws")  String repeatingRows,
            @JsonProperty("pagebreak")      Boolean pageBreak) throws XML2SpreadSheetError
    {
        this(
                worksheet,
                range == null? null: new RangeAddress(range),
                sourceSheet,
                repeatingCols,
                repeatingRows,
                pageBreak == null ? false : pageBreak);
    }

    DescriptorOutput(String worksheet, RangeAddress range,
                     String sourceSheet, String repeatingCols, String repeatingRows,
                     Boolean pageBreak) throws XML2SpreadSheetError {
        this.worksheet = worksheet;
        this.range = range;
        this.sourceSheet = sourceSheet;
        this.pageBreak = pageBreak;
        Matcher m1 = RANGE.matcher(repeatingCols == null ? "-1:-1"
                : repeatingCols);
        Matcher m2 = RANGE.matcher(repeatingRows == null ? "-1:-1"
                : repeatingRows);
        if (m1.matches() && m2.matches()) {
            this.startRepeatingColumn = Integer.parseInt(m1.group(1));
            this.endRepeatingColumn = Integer.parseInt(m1.group(2));
            this.startRepeatingRow = Integer.parseInt(m2.group(1));
            this.endRepeatingRow = Integer.parseInt(m2.group(2));
        } else {
            throw new XML2SpreadSheetError(String.format(
                    "Invalid col/row range %s %s", repeatingCols,
                    repeatingRows));
        }
    }

    String getWorksheet() {
        return worksheet;
    }

    String getSourceSheet() {
        return sourceSheet;
    }

    RangeAddress getRange() {
        return range;
    }

    public int getStartRepeatingColumn() {
        return startRepeatingColumn;
    }

    public int getEndRepeatingColumn() {
        return endRepeatingColumn;
    }

    public int getStartRepeatingRow() {
        return startRepeatingRow;
    }

    public int getEndRepeatingRow() {
        return endRepeatingRow;
    }

    public boolean getPageBreak() {
        return pageBreak;
    }
}
