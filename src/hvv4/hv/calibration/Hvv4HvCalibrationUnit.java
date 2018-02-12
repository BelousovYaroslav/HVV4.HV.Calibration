/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package hvv4.hv.calibration;

import org.dom4j.Element;

/**
 *
 * @author yaroslav
 */
public class Hvv4HvCalibrationUnit {
    private final int m_nCode;
    public int GetCode() { return m_nCode; }
    
    private final int m_nValue;
    public int GetValue() { return m_nValue; }
    
    public Hvv4HvCalibrationUnit() {
        m_nCode  = 0;
        m_nValue = 0;
    }
    
    public Hvv4HvCalibrationUnit( int nCode, int nValue) {
        m_nCode  = nCode;
        m_nValue = nValue;
    }
    
    /*
    public Hvv4HvCalibrationUnit( int nCodePreset, Element e) {
        m_nCodePreset  = nCodePreset;
        m_nCodeCurrent = Integer.parseInt( e.element( "IMant").getText());
        m_nCodeVoltage = Integer.parseInt( e.element( "UMant").getText());
        m_nCurrent     = Integer.parseInt( e.element( "IReal").getText());
        m_nVoltage     = Integer.parseInt( e.element( "UReal").getText());
        bValid = true;
    }
    */
}
