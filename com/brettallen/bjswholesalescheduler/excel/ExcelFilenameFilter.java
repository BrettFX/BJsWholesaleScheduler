/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package com.brettallen.bjswholesalescheduler.excel;

import java.io.File;
import java.io.FilenameFilter;

public class ExcelFilenameFilter implements FilenameFilter 
{
	 /**
     * Determine those files that will be returned by a call to the
     * listFiles() method. In this case, the name of the file must end with
     * either of the following two extension; '.xls' or '.xlsx'. For the
     * future, it is very possible to parameterize this and allow the
     * containing class to pass, for example, an array of Strings to this
     * class on instantiation. Each element in that array could encapsulate
     * a valid file extension - '.xls', '.xlsx', '.xlt', '.xlst', etc. These
     * could then be used to control which files were returned by the call
     * to the listFiles() method.
     *
     * @param file An instance of the File class that encapsulates a handle
     *             referring to the folder/directory that contains the file.
     * @param name An instance of the String class that encapsulates the
     *             name of the file.
     * @return A boolean value that indicates whether the file should be
     *         included in the array returned by the call to the listFiles()
     *         method. In this case true will be returned if the name of the
     *         file ends with either '.xls' or '.xlsx' and false will be
     *         returned in all other instances.
     */
	
    public boolean accept(File file, String name) {
        return(name.endsWith(".xls") || name.endsWith(".xlsx"));
    }
}
