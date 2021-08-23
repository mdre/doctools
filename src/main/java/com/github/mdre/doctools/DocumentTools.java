/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.mdre.doctools;

import com.microsoft.schemas.office.office.CTLock;
import com.microsoft.schemas.office.office.STConnectType;
import com.microsoft.schemas.vml.CTFormulas;
import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTH;
import com.microsoft.schemas.vml.CTHandles;
import com.microsoft.schemas.vml.CTPath;
import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTShapetype;
import com.microsoft.schemas.vml.CTTextPath;
import com.microsoft.schemas.vml.STExt;
import com.microsoft.schemas.vml.STTrueFalse;
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
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

/**
 *
 * @author Marcelo D. Ré {@literal <marcelo.re@gmail.com>}
 */
public class DocumentTools {

    private final static Logger LOGGER = Logger.getLogger(DocumentTools.class.getName());

    static {
        if (LOGGER.getLevel() == null) {
            LOGGER.setLevel(Level.FINEST);
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
        addWatermark(watermark, "#d8d8d8", 315, HeaderFooterType.DEFAULT);
        return this;
    }

    public DocumentTools addWatermark(String watermark, String color, float rotation, HeaderFooterType... hft) throws IOException {
        if (hft.length == 0) {
            hft = new HeaderFooterType[1];
            hft[0] = HeaderFooterType.DEFAULT;
        }
        
        for (HeaderFooterType headerFooterType : hft) {
            
            // get or create the default header
            XWPFHeader header = doc.createHeader(headerFooterType);
            // get or create first paragraph in first header
            XWPFParagraph paragraph = header.getParagraphArray(0);
            if (paragraph == null) {
                paragraph = header.createParagraph();
            }
            // set watermark to that paragraph
            setWatermarkInParagraph(paragraph, watermark, color, rotation);

        }
        return this;

    }

    public DocumentTools removeWatermark() throws IOException {
        removeWatermark(HeaderFooterType.DEFAULT,HeaderFooterType.EVEN,HeaderFooterType.FIRST);
        return this;
    }
    
    public DocumentTools removeWatermark(HeaderFooterType... hft) throws IOException {
        if (hft.length == 0) {
            hft = new HeaderFooterType[1];
            hft[0] = HeaderFooterType.DEFAULT;
        }
        
        for (HeaderFooterType headerFooterType : hft) {
            
            for (XWPFHeader hdr : doc.getHeaderList()) {
                for (XWPFParagraph p : hdr.getParagraphs()) {
                    for (CTR ctr : p.getCTP().getRList()) {
                        
                        for (int i = 0; i < ctr.getPictArray().length; i++) {
                            CTPicture pic = ctr.getPictArray(i);
                            org.apache.xmlbeans.XmlObject[] xmlobjects = pic.selectChildren(new javax.xml.namespace.QName("urn:schemas-microsoft-com:vml", "shape"));
                            for (XmlObject xmlo : xmlobjects) {
                                com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape) xmlo;
                                if (ctshape.getId().startsWith("PowerPlusWaterMarkObject")) {
                                    ctr.removePict(i);
                                }
                            }
                        }
                    }
                }
            }

        }
        
        //=====================================================
        // Borrar todo lo que sigue.
//        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
//        if (headerFooterPolicy == null) {
//            headerFooterPolicy = doc.createHeaderFooterPolicy();
//        }
//
//        // create default Watermark - fill color black and not rotated
//        headerFooterPolicy.createWatermark("");
//        // get the default header
//        // Note: createWatermark also sets FIRST and EVEN headers 
//        // but this code does not updating those other headers
//        XWPFHeader header = headerFooterPolicy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
//        XWPFParagraph paragraph = header.getParagraphArray(0);
//
//        // get com.microsoft.schemas.vml.CTShape where fill color and rotation is set
//        org.apache.xmlbeans.XmlObject[] xmlobjects = paragraph.getCTP().getRArray(0).getPictArray(0).selectChildren(
//                new javax.xml.namespace.QName("urn:schemas-microsoft-com:vml", "shape"));
//
//        if (xmlobjects.length > 0) {
//            com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape) xmlobjects[0];
//            // set fill color
//            ctshape.setFillcolor("#ffffffff");
//            // set rotation
//            ctshape.setStyle(ctshape.getStyle() + ";rotation:0");
//            //System.out.println(ctshape);
//        }
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

    //=====================================================
    //=====================================================
    private static void setWatermarkInParagraph(XWPFParagraph paragraph, String text, String color, float rotation) {
        //CTP p = CTP.Factory.newInstance();
        CTP p = paragraph.getCTP();
        XWPFDocument doc = paragraph.getDocument();
        CTBody ctBody = doc.getDocument().getBody();
        byte[] rsidr = null;
        byte[] rsidrdefault = null;
        if (ctBody.sizeOfPArray() == 0) {
            // TODO generate rsidr and rsidrdefault
        } else {
            CTP ctp = ctBody.getPArray(0);
            rsidr = ctp.getRsidR();
            rsidrdefault = ctp.getRsidRDefault();
        }
        p.setRsidP(rsidr);
        p.setRsidRDefault(rsidrdefault);
        CTPPr pPr = p.getPPr();
        if (pPr == null) {
            pPr = p.addNewPPr();
        }
        CTString pStyle = pPr.getPStyle();
        if (pStyle == null) {
            pStyle = pPr.addNewPStyle();
        }
        pStyle.setVal("Header");
        // start watermark paragraph
        CTR r = p.addNewR();
        CTRPr rPr = r.addNewRPr();
        rPr.addNewNoProof();
        int idx = 1;
        CTPicture pict = r.addNewPict();
        CTGroup group = CTGroup.Factory.newInstance();
        CTShapetype shapetype = group.addNewShapetype();
        shapetype.setId("_x0000_t136");
        shapetype.setCoordsize("1600,21600");
        shapetype.setSpt(136);
        shapetype.setAdj("10800");
        shapetype.setPath2("m@7,0l@8,0m@5,21600l@6,21600e");
        CTFormulas formulas = shapetype.addNewFormulas();
        formulas.addNewF().setEqn("sum #0 0 10800");
        formulas.addNewF().setEqn("prod #0 2 1");
        formulas.addNewF().setEqn("sum 21600 0 @1");
        formulas.addNewF().setEqn("sum 0 0 @2");
        formulas.addNewF().setEqn("sum 21600 0 @3");
        formulas.addNewF().setEqn("if @0 @3 0");
        formulas.addNewF().setEqn("if @0 21600 @1");
        formulas.addNewF().setEqn("if @0 0 @2");
        formulas.addNewF().setEqn("if @0 @4 21600");
        formulas.addNewF().setEqn("mid @5 @6");
        formulas.addNewF().setEqn("mid @8 @5");
        formulas.addNewF().setEqn("mid @7 @8");
        formulas.addNewF().setEqn("mid @6 @7");
        formulas.addNewF().setEqn("sum @6 0 @5");
        CTPath path = shapetype.addNewPath();
        path.setTextpathok(STTrueFalse.T);
        path.setConnecttype(STConnectType.CUSTOM);
        path.setConnectlocs("@9,0;@10,10800;@11,21600;@12,10800");
        path.setConnectangles("270,180,90,0");
        CTTextPath shapeTypeTextPath = shapetype.addNewTextpath();
        shapeTypeTextPath.setOn(STTrueFalse.T);
        shapeTypeTextPath.setFitshape(STTrueFalse.T);
        CTHandles handles = shapetype.addNewHandles();
        CTH h = handles.addNewH();
        h.setPosition("#0,bottomRight");
        h.setXrange("6629,14971");
        CTLock lock = shapetype.addNewLock();
        lock.setExt(STExt.EDIT);
        CTShape shape = group.addNewShape();
        shape.setId("PowerPlusWaterMarkObject" + idx);
        shape.setSpid("_x0000_s102" + (4 + idx));
        shape.setType("#_x0000_t136");
        shape.setStyle("position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
        shape.setWrapcoords("616 5068 390 16297 39 16921 -39 17155 7265 17545 7186 17467 -39 17467 18904 17467 10507 17467 8710 17545 18904 17077 18787 16843 18358 16297 18279 12554 19178 12476 20701 11774 20779 11228 21131 10059 21248 8811 21248 7563 20975 6316 20935 5380 19490 5146 14022 5068 2616 5068");
        shape.setFillcolor(color);
        shape.setStyle(shape.getStyle() + ";rotation:"+rotation);
        shape.setStroked(STTrueFalse.FALSE);
        CTTextPath shapeTextPath = shape.addNewTextpath();
        shapeTextPath.setStyle("font-family:&quot;Cambria&quot;;font-size:1pt");
        shapeTextPath.setString(text);
        pict.set(group);
    }

}
