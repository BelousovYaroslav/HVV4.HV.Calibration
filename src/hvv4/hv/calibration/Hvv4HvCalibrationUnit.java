/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package hvv4.hv.calibration;

/**
 *
 * @author yaroslav
 */
public class Hvv4HvCalibrationUnit {
    private final int m_nCodePreset;
    public int GetCodePreset() { return m_nCodePreset; }
    
    private final int m_nCodeCurrent;
    public int GetCodeCurrent() { return m_nCodeCurrent; }
    
    private final int m_nCodeVoltage;
    public int GetCodeVoltage() { return m_nCodeVoltage; }
    
    private final int m_nCurrent;
    public int GetCurrent() { return m_nCurrent; }
    
    private final int m_nVoltage;
    public int GetVoltage() { return m_nVoltage; }
    
    public Hvv4HvCalibrationUnit( int nCodePreset, int nCodeCurrent, int nCodeVoltage, int nCurrent, int nVoltage) {
        m_nCodePreset   = nCodePreset;
        m_nCodeCurrent  = nCodeCurrent;
        m_nCodeVoltage  = nCodeVoltage;
        m_nCurrent      = nCurrent;
        m_nVoltage      = nVoltage;
    }
}
