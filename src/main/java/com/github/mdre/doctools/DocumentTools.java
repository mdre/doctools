/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.mdre.doctools;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author Marcelo D. Ré {@literal <marcelo.re@gmail.com>}
 */
public class DocumentTools {

    private final static Logger LOGGER = Logger.getLogger(DocumentTools.class.getName());

    static {
        if (LOGGER.getLevel() == null) {
            LOGGER.setLevel(Level.INFO);
        }
    }

    /**
     * Fill the documento replacing the pattern with text provided.
     * 
     * @param template 
     * @param out 
     * @param fillData
     * @throws IOException 
     */
    public static void fill(File template, File out, FillerCommand fillData) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get(template.toURI())))) {
            List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
            //Iterate over paragraph list and check for the replaceable text in each paragraph
            for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                    String docText = xwpfRun.getText(0);

                    for (Iterator<Command> iterator = fillData.iterator(); iterator.hasNext();) {
                        Command fc = iterator.next();

                        docText = docText.replace(fc.pattern, fc.replace);
                        xwpfRun.setText(docText, 0);
                    }

                }
            }

            // save the docs
            try (FileOutputStream fileOut = new FileOutputStream(out)) {
                doc.write(fileOut);
            }

        }

    }

    public static void addWatermark(File document, String watermark) throws IOException {
        addWatermark(document, watermark, "#d8d8d8", 315);
    }
    
    public static void addWatermark(File document, String watermark, String color, float rotation ) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get(document.toURI())))) {
            XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
            if (headerFooterPolicy == null) {
                headerFooterPolicy = doc.createHeaderFooterPolicy();
            }

            // create default Watermark - fill color black and not rotated
            headerFooterPolicy.createWatermark(watermark);

            // get the default header
            // Note: createWatermark also sets FIRST and EVEN headers 
            // but this code does not updating those other headers
            XWPFHeader header = headerFooterPolicy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFParagraph paragraph = header.getParagraphArray(0);

            // get com.microsoft.schemas.vml.CTShape where fill color and rotation is set
            org.apache.xmlbeans.XmlObject[] xmlobjects = paragraph.getCTP().getRArray(0).getPictArray(0).selectChildren(
                    new javax.xml.namespace.QName("urn:schemas-microsoft-com:vml", "shape"));

            if (xmlobjects.length > 0) {
                com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape) xmlobjects[0];
                // set fill color
                ctshape.setFillcolor(color);
                // set rotation
                ctshape.setStyle(ctshape.getStyle() + ";rotation:"+rotation);
                //System.out.println(ctshape);
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(document)) {
                doc.write(fileOut);
                doc.close();
            }

        }

    }
}
