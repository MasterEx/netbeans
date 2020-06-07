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
import java.io.FileInputStream;
import java.io.IOException;
import static junit.framework.TestCase.assertEquals;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 *
 * @author Periklis Ntanasis <pntanasis@gmail.com>
 */
public class XLSDataExporterTest extends AbstractDataExporterTestBase {

    public XLSDataExporterTest(String name) {
        super(name, XLSDataExporter.INSTANCE, "xls");
    }

    /**
     * Compare generated file to golden file by content. It does not perform
     * exact match check because it will fail due to different metadata (i.e.
     * owner).
     *
     * @throws IOException
     */
    public void testFileCreation() throws IOException {
        File file = new File(getWorkDir(), "test.xls");

        EXPORTER.exportData(headers, contents, file);

        try ( HSSFWorkbook wb1 = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)))) {
            try ( HSSFWorkbook wb2 = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(getGoldenFile())))) {
                String workbookA = new ExcelExtractor(wb1).getText();
                String workbookB = new ExcelExtractor(wb2).getText();
                assertEquals("XLS Content Missmatch", workbookB, workbookA);
            }
        }
    }

}
