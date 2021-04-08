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
import java.util.stream.Stream;

/**
 * Реализация ReportWriter для вывода в формат OpenOffice (ODS).
 */
final class ODSReportWriter extends ReportWriter {

    private SpreadSheet template;
    private SpreadSheet result;
    private Sheet activeTemplateSheet;
    private Sheet activeResultSheet;

    ODSReportWriter(InputStream template, InputStream templateCopy) throws XylophoneError {
        try {
            this.template = new SpreadSheet(template);

            if(templateCopy == null){
                this.result = new SpreadSheet();
            } else {
                this.result = new SpreadSheet(templateCopy);
            }

        } catch ( IOException e) {
            throw new XylophoneError(e.getMessage());
        }

    }

    @Override
    void newSheet(String sheetName, String sourceSheet,
            int startRepeatingColumn, int endRepeatingColumn,
            int startRepeatingRow, int endRepeatingRow) throws XylophoneError {

        if (sourceSheet != null) {
            activeTemplateSheet = template.getSheet(sourceSheet);
        }
        if (activeTemplateSheet == null) {
            activeTemplateSheet = template.getSheet(0);
        }
        if (activeTemplateSheet == null) {
            throw new XylophoneError(String.format(
                    "Sheet '%s' does not exist.", sourceSheet));
        }

        activeResultSheet = result.getSheet(sourceSheet);
        if (activeResultSheet != null) {
            return;
        }

        try {
            activeResultSheet = new Sheet(sheetName, activeTemplateSheet.getMaxRows(), activeTemplateSheet.getMaxColumns());
            Range copyFrom = activeTemplateSheet.getDataRange();
            copyFrom.copyTo(activeResultSheet.getDataRange());
        } catch (Exception e){
            throw new XylophoneError(e.getMessage());
        }

        result.appendSheet(activeResultSheet);
    }

    @Override
    void putSection(XMLContext context, CellAddress growthPoint2,
            String sourceSheet, RangeAddress range) {
        // TODO Auto-generated method stub

    }

    @Override
    public void flush() {
        try{
            this.result.save(getOutput());
        } catch (IOException e){
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
    void applyMergedRegions(Stream<CellRangeAddress> mergedRegions){
//        mergedRegions.forEach(activeResultSheet::addMergedRegion);

    }

}
