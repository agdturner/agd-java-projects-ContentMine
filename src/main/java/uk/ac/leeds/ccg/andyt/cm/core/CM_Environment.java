/*
 * Copyright 2019 Centre for Computational Geography, University of Leeds.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package uk.ac.leeds.ccg.andyt.cm.core;

import java.io.IOException;
import uk.ac.leeds.ccg.andyt.generic.core.Generic_Environment;
import uk.ac.leeds.ccg.andyt.cm.io.CM_Files;

/**
 *
 * @author geoagdt
 */
public class CM_Environment {
    
    public transient CM_Files files;
    public transient Generic_Environment ge;
    
    public CM_Environment(Generic_Environment ge) throws IOException {
        this.ge = ge;
        files = new CM_Files(ge.files.getDir());
    }
    
}
