/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package org.netbeans.modules.db.dataview.output.dataexport;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openide.util.Exceptions;

/**
 * Exports the given data to the target file in the provided Workbook format.
 *
 * @author Periklis Ntanasis <pntanasis@gmail.com>
 */
public interface POIDataExporter {

    String DATE_FORMAT = "yyyy-mm-dd";
    String TIME_FORMAT = "hh:mm:ss";
    String TIMESTAMP_FORMAT = "yyyy-mm-dd hh:mm:ss.000";

    default void exportData(String[] headers, Object[][] contents, File file, Workbook workbook) {

        CreationHelper createHelper = workbook.getCreationHelper();

        CellStyle DATE_CELL_STYLE = workbook.createCellStyle();
        DATE_CELL_STYLE.setDataFormat(createHelper.createDataFormat().getFormat(DATE_FORMAT));
        CellStyle TIME_CELL_STYLE = workbook.createCellStyle();
        TIME_CELL_STYLE.setDataFormat(createHelper.createDataFormat().getFormat(TIME_FORMAT));
        CellStyle TIMESTAMP_CELL_STYLE = workbook.createCellStyle();
        TIMESTAMP_CELL_STYLE.setDataFormat(createHelper.createDataFormat().getFormat(TIMESTAMP_FORMAT));

        int columns = headers.length;
        int rows = contents.length;

        try {
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow(0);
            for (int j = 0; j < columns; j++) {
                Cell cell = row.createCell(j, CellType.STRING);
                cell.setCellValue(headers[j]);
            }
            for (int i = 0; i < rows; i++) {
                row = sheet.createRow(i + 1);
                for (int j = 0; j < columns; j++) {
                    Object value = contents[i][j];
                    Cell cell = row.createCell(j);
                    if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else if (value instanceof Time) {
                        cell.setCellValue((Time) value);
                        cell.setCellStyle(TIME_CELL_STYLE);
                    } else if (value instanceof Timestamp) {
                        cell.setCellValue((Timestamp) value);
                        cell.setCellStyle(TIMESTAMP_CELL_STYLE);
                    } else if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                        cell.setCellStyle(DATE_CELL_STYLE);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    } else if (value != null) {
                        cell.setCellValue(value.toString());
                    }
                }

            }

            try ( FileOutputStream outputStream = new FileOutputStream(file)) {
                workbook.write(outputStream);
                outputStream.flush();
            } finally {
                workbook.close();
            }
        } catch (IOException ex) {
            Exceptions.printStackTrace(ex);
        }

    }

}
