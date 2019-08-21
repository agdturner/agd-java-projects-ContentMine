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
package uk.ac.leeds.ccg.andyt.cm.process;

import java.io.File;
import java.io.FileFilter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import uk.ac.leeds.ccg.andyt.cm.core.CM_Environment;
import uk.ac.leeds.ccg.andyt.cm.core.CM_Object;
import uk.ac.leeds.ccg.andyt.cm.tools.CM_ParsePDF;
import uk.ac.leeds.ccg.andyt.cm.tools.CM_ParseWordDocument;
import uk.ac.leeds.ccg.andyt.generic.core.Generic_Environment;

/**
 * To be extended by Run methods.
 */
public class CM_Run extends CM_Object implements Runnable {

    public CM_Run(CM_Environment env) {
        super(env);
    }

    public static void main(String[] args) {
        new CM_Run(new CM_Environment(new Generic_Environment())).run();
    }

    @Override
    public void run() {
        String year = "2018-2019";
        String s_OnTime = "OnTime";
        String s_Late = "Late";
        String type = s_Late;
        File dir = new File(env.files.getInputDataDir(), year);
        dir = new File(dir, "turnitinuk_original_bulk_download" + type);
        File dirOut = env.files.getOutputDataDir();
        int startPage = 0;
        int endPage = 4;
        int numberOfParagraphs = 100;
        env.ge.log(dir.toString());
        File[] fs = dir.listFiles();
        try (PrintWriter pw = env.ge.io.getPrintWriter(new File(dirOut, year + type + ".txt"), false)) {
            for (File f : fs) {
                pw.write("<" + f.toString() + ">");
                String s = "";
                if (f.toString().toLowerCase().endsWith(".pdf")) {
                    s = CM_ParsePDF.getText(f, startPage, endPage);
                } else {
                    if (f.toString().toLowerCase().endsWith(".docx")) {
                        s = CM_ParseWordDocument.getText(f, numberOfParagraphs);
                    } else {
                        env.ge.log("Need a parser for file " + f);
                    }
                }
                env.ge.log(s);
                pw.write(s);
                pw.write("</" + f.toString() + ">");
            }
        }

    }

}
