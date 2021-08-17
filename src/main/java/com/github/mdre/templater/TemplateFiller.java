/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.github.mdre.templater;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
/**
 *
 * @author Marcelo D. Ré {@literal <marcelo.re@gmail.com>}
 */
public class TemplateFiller {
    private final static Logger LOGGER = Logger.getLogger(TemplateFiller.class .getName());
    static {
        if (LOGGER.getLevel() == null) {
            LOGGER.setLevel(Level.INFO);
        }
    }
    
    public static void fill(File template, File out, FillerCommand fillData ) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get(template.toURI())))
        ) {

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
}
