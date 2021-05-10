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

/**
 * Тип вывода: ODS, XLS или XLSX.
 */
public enum OutputType {
    /**
     * Формат OpenOffice.
     */
    ODS("ods"),

    /**
     * Формат MS Office 97-2003.
     */
    XLS("xls"),

    /**
     * Формат MS Office Open Document Format.
     */
    XLSX("xlsx");

    private final String extension;

    private OutputType(String extension) {
        this.extension = extension;
    }

    public String getExtension() {
        return extension;
    }

    public static OutputType fromExtension(String extension) throws XylophoneError {
        for (OutputType outputType : OutputType.values()) {
            if (outputType.extension.equalsIgnoreCase(extension)) {
                return outputType;
            }
        }
        throw new XylophoneError(
                "Cannot define output format, template has non-standard extention.");
    }
}
