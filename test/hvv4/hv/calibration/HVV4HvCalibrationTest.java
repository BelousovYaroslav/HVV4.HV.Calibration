/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package hvv4.hv.calibration;

import java.util.Date;
import org.apache.log4j.BasicConfigurator;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author yaroslav
 */
public class HVV4HvCalibrationTest {
    
    public HVV4HvCalibrationTest() {
        BasicConfigurator.configure();
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
     * Test of GetAMSRoot method, of class HVV4HvCalibration.
     */
    @Test
    public void testGetAMSRoot() {
        System.out.println("GetAMSRoot");
        HVV4HvCalibration instance = new HVV4HvCalibration();
        String expResult = "";
        String result = instance.GetAMSRoot();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of ShowCalibration method, of class HVV4HvCalibration.
     */
    @Test
    public void testShowCalibration() {
        System.out.println("ShowCalibration");
        HVV4HvCalibration instance = new HVV4HvCalibration();
        instance.ShowCalibration();
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of main method, of class HVV4HvCalibration.
     */
    @Test
    public void testMain() {
        System.out.println("main");
        String[] args = null;
        HVV4HvCalibration.main(args);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of MessageBoxInfo method, of class HVV4HvCalibration.
     */
    @Test
    public void testMessageBoxInfo() {
        System.out.println("MessageBoxInfo");
        String strMessage = "";
        String strTitleBar = "";
        HVV4HvCalibration.MessageBoxInfo(strMessage, strTitleBar);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of MessageBoxError method, of class HVV4HvCalibration.
     */
    @Test
    public void testMessageBoxError() {
        System.out.println("MessageBoxError");
        String strMessage = "";
        String strTitleBar = "";
        HVV4HvCalibration.MessageBoxError(strMessage, strTitleBar);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of GetLocalDate method, of class HVV4HvCalibration.
     */
    @Test
    public void testGetLocalDate() {
        System.out.println("GetLocalDate");
        HVV4HvCalibration instance = new HVV4HvCalibration();
        Date expResult = null;
        Date result = instance.GetLocalDate();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }
    
}
