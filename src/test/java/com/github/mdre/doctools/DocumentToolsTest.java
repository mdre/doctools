/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.mdre.doctools;

import java.io.File;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author Marcelo D. Ré {@literal <marcelo.re@gmail.com>}
 */
public class DocumentToolsTest {
    
    public DocumentToolsTest() {
    }
    
    @BeforeClass
    public static void setUpClass() {
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
    }
    
    @After
    public void tearDown() {
    }

    /**
     * Test of fill method, of class DocumentToolsDocx4j.
     */
    @Test
    public void testFill() throws Exception {
        System.out.println("fill");
        File template = new File(System.getProperty("user.dir") + "/src/test/resources/template.docx");
        File out = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled.docx");
        FillerCommand fillData = new FillerCommand()
                                            .add("${nombre}", "Elba Gallo")
                                            .add("${numero}", "1234")
                                            .add("${texto}", "texto")
                                            .add("${header}", "Cabecera!!")
                                            .add("${footer}", "pata 1!!")
                                            .add("${footer2}", "pata 2!!")
                ;
        DocumentTools dft = new DocumentTools(template).fill(fillData).save(out);
        // TODO review the generated test code and remove the default call to fail.
        
    }
    
    @Test
    public void testAddWatermark() throws Exception {
        System.out.println("agregar un watermark");
        String timeStamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new Timestamp(System.currentTimeMillis()));
        File in = new File(System.getProperty("user.dir") + "/src/test/resources/wm_to_add_watermark.docx");
        //File in = new File("/home/mdre/tmp/1/template2.docx");
        File out = new File(System.getProperty("user.dir") + "/src/test/resources/wm_with_watermark.docx");
        
        DocumentTools dft = new DocumentTools(in)
                                        .addWatermark(timeStamp,"#ff0000",45)
//                                        .addWatermark(timeStamp)
                                        //.fill(fillData)
                                        .save(out)
                ;
    }
    
    @Test
    public void testRemoveWatermark() throws Exception {
        System.out.println("quitar un watermark");
        File in = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled_with_watermark.docx");
        File out = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled_without_watermark.docx");
        DocumentTools dft = new DocumentTools(in).removeWatermark().save(out);
    }
    
    @Test
    public void testHeader() throws Exception {
        System.out.println("agregar header");
        File in = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled.docx");
        File out = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled_with_header.docx");
        DocumentTools dft = new DocumentTools(in).addHeader("cabecera de prueba").save(out);
    }
    
    
    
    @Test
    public void testToPdf() throws Exception {
        System.out.println("Convertir a PDF");
        File in = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled.docx");
        File outPdf = new File(System.getProperty("user.dir") + "/src/test/resources/template_filled.pdf");
        DocumentTools dft = new DocumentTools(in).convertToPDF(outPdf);
    }
    
    
    
    
}
