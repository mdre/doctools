/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.mdre.doctools;

import java.io.File;
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
     * Test of fill method, of class DocumentTools.
     */
    @Test
    public void testFill() throws Exception {
        System.out.println("fill");
        File template = new File("/home/mdre/tmp/1/template.docx");
        File out = new File("/tmp/1/template_filled.docx");
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
        File in = new File("/tmp/1/template_filled.docx");
        File out = new File("/tmp/1/template_filled_with_watermark.docx");
        DocumentTools dft = new DocumentTools(in).addWatermark("Borrador").save(out);
    }
    
    @Test
    public void testRemoveWatermark() throws Exception {
        System.out.println("quitar un watermark");
        File in = new File("/tmp/1/template_filled_with_watermark.docx");
        File out = new File("/tmp/1/template_filled_without_watermark.docx");
        DocumentTools dft = new DocumentTools(in).removeWatermark().save(out);
    }
    
    @Test
    public void testHeader() throws Exception {
        System.out.println("agregar header");
        File in = new File("/tmp/1/template_filled.docx");
        File out = new File("/tmp/1/template_filled_with_header.docx");
        DocumentTools dft = new DocumentTools(in).addHeader("cabecera de prueba").save(out);
    }
    
    
    
    @Test
    public void testToPdf() throws Exception {
        System.out.println("Convertir a PDF");
        File in = new File("/tmp/1/template_filled.docx");
        File outPdf = new File("/tmp/1/template_filled.pdf");
        DocumentTools dft = new DocumentTools(in).convertToPDF(outPdf);
    }
    
    
    
    
}
