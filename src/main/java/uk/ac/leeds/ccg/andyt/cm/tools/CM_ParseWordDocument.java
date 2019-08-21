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
package uk.ac.leeds.ccg.andyt.cm.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jdom2.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

/**
 *
 * @author geoagdt
 */
public class CM_ParseWordDocument {

    /**
     * For parsing and returning text from a word document.
     *
     * @param f
     * @param n Number of Paragraphs
     * @return
     */
    public static String getText(File f, int n) {
        String r = "";
        try (FileInputStream fis = new FileInputStream(f)) {
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));
            //XWPFWordExtractor ext = new XWPFWordExtractor(doc);
            //System.out.println(ext.getText());  
            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            boolean complete = false;
            int c = 1;
            int i = 0;
            while (!complete) {
                XWPFParagraph paragraph = paragraphs.get(i);
                //r += paragraph.getText() + "\r\n";
                String s = paragraph.getText();
                if (!s.isBlank()) {
                    if (s.trim().endsWith("Introduction")) {
                        return r;
                    }
                    c ++;
                    r += s + "\n";
                }
                i ++;
                if (c == n) {
                    complete = true;
                }
            }
        } catch (IOException e) {
            System.out.println(e);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(CM_ParseWordDocument.class.getName()).log(Level.SEVERE, null, ex);
        }
        return r;
    }

//    public static void splitIntoPages() {
//        XWPFDocument doc = null;
//
//        try {
//            //Input Word Document
//            File file = new File("C:/TestDoc.docx");
//            FileInputStream in = new FileInputStream(file);
//            doc = new XWPFDocument(in);
//
//            //Determine how many paragraphs per page
//            List<Integer> paragraphCountList = getParagraphCountPerPage(doc);
//
//            if (paragraphCountList != null && paragraphCountList.size() > 0) {
//                int docCount = 0;
//                int startIndex = 0;
//                int endIndex = paragraphCountList.get(0);
//
//                //Loop through the paragraph counts for each page
//                for (int i = 0; i < paragraphCountList.size(); i++) {
//                    XWPFDocument outputDocument = new XWPFDocument();
//
//                    List<XWPFParagraph> paragraphs = doc.getParagraphs();
//                    List<XWPFParagraph> pageParagraphs = new ArrayList<XWPFParagraph>();
//
//                    if (paragraphs != null && paragraphs.size() > 0) {
//                        //Get the paragraphs from the input Word document
//                        for (int j = startIndex; j < endIndex; j++) {
//                            if (paragraphs.get(j) != null) {
//                                pageParagraphs.add(paragraphs.get(j));
//                            }
//                        }
//
//                        //Set the start and end point for the next set of paragraphs
//                        startIndex = endIndex;
//
//                        if (i < paragraphCountList.size() - 2) {
//                            endIndex = endIndex + paragraphCountList.get(i + 1);
//                        } else {
//                            endIndex = paragraphs.size() - 1;
//                        }
//
//                        //Create a new Word Doc with the paragraph subset
//                        createPageInAnotherDocument(outputDocument, pageParagraphs);
//
//                        //Write the file
//                        String outputPath = "C:/TestDocOutput" + docCount + ".docx";
//                        FileOutputStream outputStream = new FileOutputStream(outputPath);
//                        outputDocument.write(outputStream);
//                        outputDocument.close();
//
//                        docCount++;
//                        pageParagraphs = new ArrayList<XWPFParagraph>();
//                    }
//                }
//            }
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        } finally {
//            try {
//                doc.close();
//            } catch (IOException ioe) {
//                ioe.printStackTrace();
//            }
//        }
//    }
//
//    private static List<Integer> getParagraphCountPerPage(XWPFDocument doc) throws Exception {
//        List<Integer> paragraphCountList = new ArrayList<>();
//        int paragraphCount = 0;
//
//        Document domDoc = convertStringToDOM(doc.getDocument().getBody().toString());
//        NodeList rootChildNodeList = domDoc.getChildNodes().item(0).getChildNodes();
//
//        for (int i = 0; i < rootChildNodeList.getLength(); i++) {
//            Node childNode = rootChildNodeList.item(i);
//
//            if (childNode.getNodeName().equals("w:p")) {
//                paragraphCount++;
//
//                if (childNode.getChildNodes() != null) {
//                    for (int k = 0; k < childNode.getChildNodes().getLength(); k++) {
//                        if (childNode.getChildNodes().item(k).getNodeName().equals("w:r")) {
//                            for (int m = 0; m < childNode.getChildNodes().item(k).getChildNodes().getLength(); m++) {
//                                if (childNode.getChildNodes().item(k).getChildNodes().item(m).getNodeName().equals("w:br")) {
//
//                                    paragraphCountList.add(paragraphCount);
//                                    paragraphCount = 0;
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//        }
//
//        paragraphCountList.add(paragraphCount + 1);
//
//        return paragraphCountList;
//    }
//
//    private static Document convertStringToDOM(String xmlData) throws Exception {
//        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
//        DocumentBuilder builder = factory.newDocumentBuilder();
//
//        Document document = builder.parse(new InputSource(new StringReader(xmlData)));
//
//        return document;
//    }
//
//    private static void createPageInAnotherDocument(XWPFDocument outputDocument, List<XWPFParagraph> pageParagraphs) throws IOException {
//        for (int i = 0; i < pageParagraphs.size(); i++) {
//            addParagraphToDocument(outputDocument, pageParagraphs.get(i).getText());
//        }
//    }
//
//    private static void addParagraphToDocument(XWPFDocument outputDocument, String text) throws IOException {
//        XWPFParagraph paragraph = outputDocument.createParagraph();
//        XWPFRun run = paragraph.createRun();
//        run.setText(text);
//    }
}
