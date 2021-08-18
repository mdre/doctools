/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.mdre.doctools;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

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

    XWPFDocument doc;
    File inFile;

    public DocumentTools(File inFile) throws IOException {
        this.inFile = inFile;
        this.doc = new XWPFDocument(Files.newInputStream(Paths.get(this.inFile.toURI())));
    }

    /**
     * Fill the documento replacing the pattern with text provided.
     *
     * @param template
     * @param out
     * @param fillData
     *
     * @throws IOException
     */
    public DocumentTools fill(FillerCommand fillData) throws IOException {

        List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
        //Iterate over paragraph list and check for the replaceable text in each paragraph
        for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                String docText = xwpfRun.getText(0);

                if (docText != null) {
                    for (Iterator<Command> iterator = fillData.iterator(); iterator.hasNext();) {
                        Command fc = iterator.next();

                        docText = docText.replace(fc.pattern, fc.replace);
                        xwpfRun.setText(docText, 0);
                    }
                }

            }
        }

        // reemplazar en cabeceras
        for (XWPFHeader xWPFHeader : doc.getHeaderList()) {
            for (XWPFParagraph xwpfParagraph : xWPFHeader.getParagraphs()) {
                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                    String docText = xwpfRun.getText(0);
                    if (docText != null) {
                        for (Iterator<Command> iterator = fillData.iterator(); iterator.hasNext();) {
                            Command fc = iterator.next();

                            docText = docText.replace(fc.pattern, fc.replace);
                            xwpfRun.setText(docText, 0);
                        }
                    }
                }
            }
        }

        // reemplazar en pié de página
        for (XWPFFooter xWPFFooter : doc.getFooterList()) {
            for (XWPFParagraph xwpfParagraph : xWPFFooter.getParagraphs()) {
                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                    String docText = xwpfRun.getText(0);
                    System.out.println("docText:" + docText);
                    if (docText != null) {
                        for (Iterator<Command> iterator = fillData.iterator(); iterator.hasNext();) {
                            Command fc = iterator.next();

                            docText = docText.replace(fc.pattern, fc.replace);
                            xwpfRun.setText(docText, 0);
                        }
                    }
                }
            }
        }

        return this;

    }

    public DocumentTools addWatermark(String watermark) throws IOException {
        addWatermark(watermark, "#d8d8d8", 315);
        return this;
    }

    public DocumentTools addWatermark(String watermark, String color, float rotation) throws IOException {

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
            ctshape.setStyle(ctshape.getStyle() + ";rotation:" + rotation);
            //System.out.println(ctshape);
        }

        return this;

    }

    public DocumentTools removeWatermark() throws IOException {

        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        if (headerFooterPolicy == null) {
            headerFooterPolicy = doc.createHeaderFooterPolicy();
        }

        // create default Watermark - fill color black and not rotated
        headerFooterPolicy.createWatermark("");
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
            ctshape.setFillcolor("#ffffffff");
            // set rotation
            ctshape.setStyle(ctshape.getStyle() + ";rotation:0");
            //System.out.println(ctshape);
        }
        return this;

    }

    public DocumentTools convertToPDF(File pdfPath) {
        try {
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(pdfPath);
            PdfConverter.getInstance().convert(this.doc, out, options);
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
        return this;
    }

    
    public DocumentTools addHeader(String header) {

        XWPFParagraph p = doc.createParagraph();

        CTP ctP = CTP.Factory.newInstance();
        CTText t = ctP.addNewR().addNewT();
        t.setStringValue(header);
        XWPFParagraph[] pars = new XWPFParagraph[1];
        p = new XWPFParagraph(ctP, doc);
        pars[0] = p;

        XWPFHeaderFooterPolicy hfPolicy = doc.createHeaderFooterPolicy();
        hfPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, pars);

        return this;
    }

    public DocumentTools addFooter(String footer) {

        XWPFParagraph p = doc.createParagraph();

        CTP ctP = CTP.Factory.newInstance();
        CTText t = ctP.addNewR().addNewT();
        t.setStringValue(footer);
        XWPFParagraph[] pars = new XWPFParagraph[1];
        p = new XWPFParagraph(ctP, doc);
        pars[0] = p;

        XWPFHeaderFooterPolicy hfPolicy = doc.createHeaderFooterPolicy();
        hfPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, pars);

        return this;
    }

    public DocumentTools save(File outFile) throws IOException {
        // save the docs
        try (FileOutputStream fileOut = new FileOutputStream(outFile)) {
            doc.write(fileOut);
        }
        return this;
    }

    /**
     * Close document.
     */
    public void close() throws IOException {
        this.doc.close();
    }
}
