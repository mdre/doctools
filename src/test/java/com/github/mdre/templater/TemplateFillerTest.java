/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.github.mdre.templater;

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
public class TemplateFillerTest {
    
    public TemplateFillerTest() {
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
     * Test of fill method, of class TemplateFiller.
     */
    @Test
    public void testFill() throws Exception {
        System.out.println("fill");
        File template = new File("/home/mdre/tmp/1/template.docx");
        File out = new File("/tmp/1/template_filled.docx");
        FillerCommand fillData = new FillerCommand().add("${nombre}", "Elba Gallo").add("${numero}", "1234").add("${texto}", "texto");
        TemplateFiller.fill(template, out, fillData);
        // TODO review the generated test code and remove the default call to fail.
    }
    
}
