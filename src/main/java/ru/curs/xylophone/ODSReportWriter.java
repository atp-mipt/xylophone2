/*
   (с) 2016 ООО "КУРС-ИТ"

   Этот файл — часть КУРС:Xylophone.

   КУРС:Xylophone — свободная программа: вы можете перераспространять ее и/или изменять
   ее на условиях Стандартной общественной лицензии ограниченного применения GNU (LGPL)
   в том виде, в каком она была опубликована Фондом свободного программного обеспечения; либо
   версии 3 лицензии, либо (по вашему выбору) любой более поздней версии.

   Эта программа распространяется в надежде, что она будет полезной,
   но БЕЗО ВСЯКИХ ГАРАНТИЙ; даже без неявной гарантии ТОВАРНОГО ВИДА
   или ПРИГОДНОСТИ ДЛЯ ОПРЕДЕЛЕННЫХ ЦЕЛЕЙ. Подробнее см. в Стандартной
   общественной лицензии GNU.

   Вы должны были получить копию Стандартной общественной лицензии  ограниченного
   применения GNU (LGPL) вместе с этой программой. Если это не так,
   см. http://www.gnu.org/licenses/.


   Copyright 2016, COURSE-IT Ltd.

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Lesser General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Lesser General Public License for more details.
   You should have received a copy of the GNU Lesser General Public License
   along with this program.  If not, see http://www.gnu.org/licenses/.

*/
package ru.curs.xylophone;

import com.github.miachm.sods.Range;
import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Реализация ReportWriter для вывода в формат OpenOffice (ODS).
 */
final class ODSReportWriter extends ReportWriter {

    private SpreadSheet template;
    private SpreadSheet result;
    private Sheet activeTemplateSheet;
    private Sheet activeResultSheet;
    private final MergeRegionContainer mergeRegionContainer = MergeRegionContainer.getContainer();

    ODSReportWriter(InputStream template, InputStream templateCopy) throws XylophoneError {
        try {
            this.template = new SpreadSheet(template);

            if (templateCopy == null) {
                this.result = new SpreadSheet();
            } else {
                this.result = new SpreadSheet(templateCopy);
            }

        } catch (IOException e) {
            throw new XylophoneError(e.getMessage());
        }

    }

    /**
     * Create new sheet "sheetName" from "sourceSheet" with styles. If sourceSheet null, create empty sheet
     *
     * @param sheetName            name of sheet
     * @param sourceSheet          source to copy
     * @param startRepeatingColumn ignore
     * @param endRepeatingColumn   ignore
     * @param startRepeatingRow    ignore
     * @param endRepeatingRow      ignore
     * @throws XylophoneError
     */
    @Override
    void newSheet(String sheetName, String sourceSheet,
                  int startRepeatingColumn, int endRepeatingColumn,
                  int startRepeatingRow, int endRepeatingRow) throws XylophoneError {
        System.out.printf("new sheet  %s<-%s%n", sheetName, sourceSheet);
        updateActiveTemplateSheet(sourceSheet);

        activeResultSheet = result.getSheet(sourceSheet);
        if (activeResultSheet != null) {
            return;
        }

        try {
            activeResultSheet = new Sheet(sheetName, activeTemplateSheet.getMaxRows(), activeTemplateSheet.getMaxColumns());
            Range copyFrom = activeTemplateSheet.getDataRange();
            copyFrom.copyTo(activeResultSheet.getDataRange());
        } catch (Exception e) {
            throw new XylophoneError(e.getMessage());
        }

        result.appendSheet(activeResultSheet);
    }

    private void updateActiveTemplateSheet(String sourceSheet) throws XylophoneError {
        System.out.printf("updateActiveTemplateSheet %s%n", sourceSheet);
        if (sourceSheet != null) {
            activeTemplateSheet = template.getSheet(sourceSheet);
        }
        if (activeTemplateSheet == null) {
            activeTemplateSheet = template.getSheet(0);
            System.out.println("not found, fall back to sheet 0");
        }
        if (activeTemplateSheet == null) {
            throw new XylophoneError(String.format(
                    "Sheet '%s' does not exist.", sourceSheet));
        }
    }

    private int getLastRowNum(Sheet activeSheet) {
        int k = 0;
        String data = String.valueOf(activeSheet.getRange(k, 1).getValue());
        while (!data.equals("null")) {
            k++;
            if (k >= activeSheet.getMaxRows()) {
                k = activeSheet.getMaxRows();
                break;
            }
            Range test = activeSheet.getRange(k, 1);
            data = String.valueOf(test.getValue());
        }
        return k;
    }

    private int getLastCellNum(Sheet activeSheet, int rowNum) {
        int k = 0;
        String data = String.valueOf(activeSheet.getRange(rowNum, k).getValue());
        while (!data.equals("null")) {
            k++;
            if (k >= activeSheet.getMaxColumns()) {
                k = activeSheet.getMaxColumns();
                break;
            }
            Range test = activeSheet.getRange(rowNum, k);
            data = String.valueOf(test.getValue());
        }
        return k;
    }

    @Override
    void putSection(XMLContext context, CellAddress growthPoint,
                    String sourceSheet, RangeAddress range) throws XylophoneError {

        System.out.printf("put section %s, %s, %s%n", growthPoint.getAddress(), sourceSheet, range.getAddress());
        updateActiveTemplateSheet(sourceSheet);
        if (activeResultSheet == null) {
            newSheet("Sheet1", sourceSheet, -1, -1, -1, -1);
        }

        int rowStart = range.top();
//        в реализации XLS тут максимум, но из-за особенностей библиотеки ODS пришлось
//        сделать просто проверку в цикле на выход за пределы страницы
//        int rowFinish = Math.max(range.bottom(), getLastRowNum(activeResultSheet));
        int rowFinish = range.bottom();
        System.out.println("first " + range.bottom() + " : " + getLastRowNum(activeResultSheet));
        for (int i = rowStart; i <= rowFinish; i++) {
            final int numColumns = getLastCellNum(activeTemplateSheet, i - 1);
            if (i >= activeTemplateSheet.getMaxRows()) {
                continue;
            }
            Range sourceRow = activeTemplateSheet.getRange(i - 1, 0, 1, numColumns);
            if (sourceRow.getValue() == null) {
                continue;
            }
            Range resultRow;
            if (growthPoint.getRow() + i - rowStart >= activeResultSheet.getMaxRows()) {
                activeResultSheet.appendRow();
            }
            resultRow = activeResultSheet.getRange(
                    growthPoint.getRow() + i - rowStart - 1, 0, 1, numColumns);

            // Копируем стиль ...
            resultRow.setStyle(sourceRow.getStyle());

            int colStart = range.left();
            int colFinish = Math.min(range.right(), numColumns);
            System.out.println(range.right() + ":" + numColumns);
            for (int j = colStart; j <= colFinish; j++) {
                Range sourceCell = sourceRow.getCell(0, j - 1);
                if (sourceCell.getValue() == null) {
                    continue;
                }
                if(resultRow.getLastColumn() < growthPoint.getCol() + j - colStart - 1){
                    activeResultSheet.appendColumns(j);
                    resultRow = activeResultSheet.getRange(
                            growthPoint.getRow() + i - rowStart - 1, 0, 1, numColumns+j);
                    resultRow.setStyle(sourceRow.getStyle());
                }
                Range resultCell = resultRow.getCell(0,
                        growthPoint.getCol() + j - colStart - 1);

                // Копируем стиль...
                resultCell.setStyle(sourceCell.getStyle());

                // Копируем значение...
                // ДЛЯ СТРОКОВЫХ ЯЧЕЕК ВЫЧИСЛЯЕМ ПОДСТАНОВКИ!!
                String val;
                String buf;
                if (sourceCell.getValue() == null) {
                    continue;
                }
                if (sourceCell.getValue().getClass().equals(String.class)) {
                    val = String.valueOf(sourceCell.getValue());
                    buf = context.calc(val);
                    DynamicCellWithStyle<Range> cellWithStyle = DynamicCellWithStyle.defineCellStyle(sourceCell, buf);
                    // Если ячейка содержит строковое представление числа и при
                    // этом содержит плейсхолдер --- меняем его на число.
                    System.out.print(cellWithStyle.isStylesPresent() + "\n");
                    if (!cellWithStyle.isStylesPresent()) {
                        writeTextOrNumber(resultCell, buf,
                                context.containsPlaceholder(val));
                    } else {
                        Map<String, String> properties = cellWithStyle.getProperties();
                        for (Map.Entry<String, String> entry : properties.entrySet()) {
                            switch (entry.getKey().toUpperCase()) {
                                case CellPropertyType.MERGE_LEFT_VALUE:
                                    mergeLeft(entry.getValue(), resultCell, cellWithStyle);
                                    break;
                                case CellPropertyType.MERGE_UP_VALUE:
                                    mergeUp(entry.getValue(), resultCell, cellWithStyle);
                                    break;
                                case CellPropertyType.MERGE_UP_LEFT_VALUE:
                                    mergeUp(entry.getValue(), resultCell, cellWithStyle);
                                    mergeLeft(entry.getValue(), resultCell, cellWithStyle);
                                    break;
                                case CellPropertyType.MERGE_LEFT_UP_VALUE:
                                    mergeLeft(entry.getValue(), resultCell, cellWithStyle);
                                    mergeUp(entry.getValue(), resultCell, cellWithStyle);
                                    break;
                                default:
                                    break;
                            }
                        }
                        writeTextOrNumber(resultCell, cellWithStyle.getValue(),
                                context.containsPlaceholder(val));
                    }
                } else if (sourceCell.getValue().getClass().equals(Float.class)) {
                    // Обрабатываем формулу
                    // ??? нет такого типа в SODS
                    // "The values could be String, Float, Integer, OfficeCurrency,
                    // OfficePercentage or a Date Empty cells returns a null object"
                    resultCell.setValue(sourceCell.getValue());
                } else {
                    resultCell.setValue(sourceCell.getValue());
                }
            }
        }

        // Разбираемся с merged-ячейками
//        arrangeMergedCells(growthPoint, range);

    }

    private void mergeUp(String attribute, Range resultCell, DynamicCellWithStyle cellWithStyle) {
        if (!CellPropertyType.MERGE_UP.contains(attribute.toLowerCase())) {
            String propertyValues = Arrays.stream(CellPropertyType.MERGE_UP.getValues())
                    .collect(Collectors.joining(", "));
            throw new IllegalArgumentException(
                    String.format("There are no such value: %s. Please choice one of %s",
                            attribute, propertyValues));
        }

        switch (attribute.toLowerCase()) {
            case CellPropertyType.MERGE_YES:
                mergeRegionContainer.mergeUp(
                        new CellAddress(resultCell.toString()));
                break;
            case CellPropertyType.MERGE_IFEQUALS:
                CellRangeAddress rangeAddress = new CellRangeAddress(
                        resultCell.getRow() - 1, resultCell.getRow(),
                        resultCell.getColumn(), resultCell.getColumn());

                if (ifEquals(rangeAddress, cellWithStyle)) {
                    mergeRegionContainer.mergeUp(
                            new CellAddress(resultCell.toString()));
                }
                break;
            case CellPropertyType.MERGE_NO:
            default:
                break;
        }
    }

    private void mergeLeft(String attribute, Range resultCell, DynamicCellWithStyle cellWithStyle) {
        if (!CellPropertyType.MERGE_LEFT.contains(attribute.toLowerCase())) {
            String propertyValues = Arrays.stream(CellPropertyType.MERGE_LEFT.getValues())
                    .collect(Collectors.joining(", "));
            throw new RuntimeException(
                    String.format("There are no such value: %s. Please choice one of %s",
                            attribute, propertyValues));
        }

        switch (attribute.toLowerCase()) {
            case CellPropertyType.MERGE_YES:
                mergeRegionContainer.mergeLeft(
                        new CellAddress(resultCell.toString()));
                break;
            case CellPropertyType.MERGE_IFEQUALS:
                CellRangeAddress rangeAddress = new CellRangeAddress(
                        resultCell.getRow(), resultCell.getRow(),
                        resultCell.getColumn() - 1, resultCell.getColumn());

                if (ifEquals(rangeAddress, cellWithStyle)) {
                    mergeRegionContainer.mergeLeft(
                            new CellAddress(resultCell.toString()));
                }
                break;
            case CellPropertyType.MERGE_NO:
                break;
        }
    }

    private boolean ifEquals(CellRangeAddress rangeAddress, DynamicCellWithStyle cellWithStyle) {
        try {
            CellRangeAddress mergedRegion =
                    mergeRegionContainer.findIntersectedRange(rangeAddress);

            Range cell = activeResultSheet.getDataRange().getCell(
                    mergedRegion.getFirstRow(), mergedRegion.getFirstColumn());

            return cell.getFormula().equalsIgnoreCase(cellWithStyle.getValue());
        } catch (IllegalArgumentException exc) {
            System.out.println(exc.getMessage());
        }
        return false;
    }

    private void writeTextOrNumber(Range resultCell, String buf, boolean decide) {
        System.out.printf("writeTextOrNumber %d:%d, '%s', %s%n", resultCell.getColumn(),
                resultCell.getRow(), buf, decide);

        final Pattern NUMBER = Pattern
                .compile("[+-]?\\d+(\\.\\d+)?([eE][+-]?\\d+)?");
        /**
         * Регексп для даты в ISO-формате.
         */
        final Pattern DATE = Pattern
                .compile("(\\d\\d\\d\\d)-(0[1-9]|1[0-2])-(0[1-9]|[12]\\d|3[01])");

        if (decide
                && !"@".equals(resultCell.getFormula())) {
            Matcher numberMatcher = NUMBER.matcher(buf.trim());
            Matcher dateMatcher = DATE.matcher(buf.trim());
            // может, число?
            if (numberMatcher.matches())
                resultCell.setValue(Double.parseDouble(buf));
                // может, дата?
            else if (dateMatcher.matches()) {
                Calendar c = Calendar.getInstance();
                c.clear();
                c.set(Integer.parseInt(dateMatcher.group(1)),
                        Integer.parseInt(dateMatcher.group(2)) - 1,
                        Integer.parseInt(dateMatcher.group(3)));
                resultCell.setValue(c.getTime());
            } else {
                resultCell.setValue(buf);
            }
        } else {
            resultCell.setValue(buf);
        }
    }

    @Override
    public void flush() {
        System.out.println("flush");
        try {
            this.result.save(getOutput());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    void mergeUp(CellAddress a1, CellAddress a2) {
        int countRowsToMerge = 1 + Math.abs(a1.getRow() - a2.getRow());
        int countColumnsToMerge = 1 + Math.abs(a1.getCol() - a2.getCol());

        Range toMerge = activeResultSheet.getRange(a1.getRow(), a1.getCol(),
                countRowsToMerge, countColumnsToMerge);

        toMerge.merge();
    }

    // пока не понял что эта функция делает
    @Override
    void addNamedRegion(String name, CellAddress a1, CellAddress a2) {
        // TODO Auto-generated method stub

    }

    // from javadoc apache poi
    // Sets a page break at the indicated row Breaks occur above
    // the specified row and left of the specified column inclusive.
    @Override
    void putRowBreak(int rowNumber) {
        // TODO Auto-generated method stub

    }

    // from javadoc apache poi
    // Sets a page break at the indicated column.
    @Override
    void putColBreak(int colNumber) {
        // TODO Auto-generated method stub

    }

    // import org.apache.poi.ss.util.CellRangeAddress;
    @Override
    void applyMergedRegions(Stream<CellRangeAddress> mergedRegions) {
//        mergedRegions.forEach(activeResultSheet::addMergedRegion);

    }

}
