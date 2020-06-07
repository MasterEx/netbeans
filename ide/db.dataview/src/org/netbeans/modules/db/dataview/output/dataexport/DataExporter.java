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
import javax.swing.filechooser.FileFilter;

/**
 * A DataExporter is used to export the contents of the JTable containing the
 * rows returned by an SQL command to a file. In case you want to add a new
 * Export method implement this class and add it to
 * DataViewTableDataExportFileChooser EXPORTERS.
 *
 * @author Periklis Ntanasis <pntanasis@gmail.com>
 * @see
 * org.netbeans.modules.db.dataview.output.dataexport.DataViewTableDataExportFileChooser#EXPORTERS
 */
public interface DataExporter {

    /**
     * Returns true if the given file's filename extension is handled by this
     * exporter.
     *
     * @param file
     * @return True if the file's filename extension is handled by this
     * exporter. Otherwise false.
     */
    public boolean handlesFileFormat(File file);

    /**
     * Returns the FileFilter matching the file types this exporter handles.
     *
     * @return FileFilter
     */
    public FileFilter getFileFilter();

    public void exportData(String[] headers, Object[][] contents, File file);

    /**
     * Returns the file extension which is appended to the files created by this
     * exporter.
     *
     * @return File extension as String. Example "csv".
     */
    public String getDefaultFileExtension();

}
