/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package hvv4.hv.calibration;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.Timer;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.DefaultTableModel;
import jssc.SerialPort;
import jssc.SerialPortException;
import jssc.SerialPortTimeoutException;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

/**
 *
 * @author yaroslav
 */
public class HVV4HvCalibration extends javax.swing.JFrame {
    
    private final String m_strAMSrootEnvVar;
    public String GetAMSRoot() { return m_strAMSrootEnvVar; }
    
    static Logger logger = Logger.getLogger( HVV4HvCalibration.class);
    TreeMap m_pCalibrationP;
    TreeMap m_pCalibrationI;
    TreeMap m_pCalibrationU;
    boolean m_bConnected;
    boolean m_bGettingData;
    
    int nSummCurrent, nSummVoltage;
    SerialPort m_port;
    Timer tRefreshState;
    
    /**
     * Creates new form HVV4HvCalibrationMainFrame
     */
    public HVV4HvCalibration() {
        initComponents();
        
        String strOS = System.getProperty("os.name");
        logger.info( "OS:" + strOS);
        if( strOS.contains("inu")) {
            //setResizable( true);
            Dimension d = getSize();
            d.height -= 40;
            setSize( d);
            //setResizable( false);
        }
        m_bConnected = false;
        m_bGettingData = false;
        
        setTitle( "Утилита проведения калибровки в/в модулей. 2018.02.08 13:30. (С) ФЛАВТ, 2018");
        
        tblPCalib.getSelectionModel().addListSelectionListener( new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                int nRow = tblPCalib.getSelectedRow();
                if( nRow != -1 && nRow < tblPCalib.getRowCount()) {
                    edtPcode.setText( tblPCalib.getModel().getValueAt( nRow, 0).toString());
                    edtPvalue.setText( tblPCalib.getModel().getValueAt( nRow, 1).toString());
                }
            }
            
        });
        
        tblICalib.getSelectionModel().addListSelectionListener( new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                int nRow = tblICalib.getSelectedRow();
                if( nRow != -1 && nRow < tblICalib.getRowCount()) {
                    edtIcode.setText( tblICalib.getModel().getValueAt( nRow, 0).toString());
                    edtIvalue.setText( tblICalib.getModel().getValueAt( nRow, 1).toString());
                }
            }
            
        });
        
        tblUCalib.getSelectionModel().addListSelectionListener( new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                int nRow = tblUCalib.getSelectedRow();
                if( nRow != -1 && nRow < tblUCalib.getRowCount()) {
                    edtUcode.setText( tblUCalib.getModel().getValueAt( nRow, 0).toString());
                    edtUvalue.setText( tblUCalib.getModel().getValueAt( nRow, 1).toString());
                }
            }
            
        });
        
        m_strAMSrootEnvVar = System.getenv( "AMS_ROOT");
        
        m_pCalibrationP = new TreeMap();
        m_pCalibrationI = new TreeMap();
        m_pCalibrationU = new TreeMap();
        
        tRefreshState = new Timer( 250, new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent e) {
                edtComPortValue.setEditable( !m_bConnected);
                btnConnect.setEnabled( !m_bConnected);
                btnDisconnect.setEnabled( m_bConnected & !m_bGettingData);
                btnTurnOff.setEnabled( m_bConnected & !m_bGettingData);
                btnPresetApply.setEnabled( m_bConnected & !m_bGettingData);
                btnReqCodes.setEnabled( m_bConnected);
                
                btnReqCodes.setVisible( !m_bGettingData);
                jProgressBar1.setVisible( m_bGettingData);
                
                btnAcceptPPoint.setEnabled( !m_bGettingData);
                btnRemovePPoint.setEnabled( !m_bGettingData && ( tblPCalib.getSelectedRow() != -1));
                btnLoadP.setEnabled( !m_bGettingData);
                btnSavePXML.setEnabled( !m_bGettingData);
                btnSavePXLS.setEnabled( !m_bGettingData);
                
                btnAcceptIPoint.setEnabled( !m_bGettingData);
                btnRemoveIPoint.setEnabled( !m_bGettingData && ( tblICalib.getSelectedRow() != -1));
                btnLoadI.setEnabled( !m_bGettingData);
                btnSaveIXML.setEnabled( !m_bGettingData);
                btnSaveIXLS.setEnabled( !m_bGettingData);
                
                btnAcceptUPoint.setEnabled( !m_bGettingData);
                btnRemoveUPoint.setEnabled( !m_bGettingData && ( tblUCalib.getSelectedRow() != -1));
                btnLoadU.setEnabled( !m_bGettingData);
                btnSaveUXML.setEnabled( !m_bGettingData);
                btnSaveUXLS.setEnabled( !m_bGettingData);
            }
        });
        tRefreshState.start();
        
        jProgressBar1.setVisible( false);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        lblComPortTitle = new javax.swing.JLabel();
        edtComPortValue = new javax.swing.JTextField();
        btnConnect = new javax.swing.JButton();
        btnDisconnect = new javax.swing.JButton();
        btnTurnOff = new javax.swing.JButton();
        lblPreset = new javax.swing.JLabel();
        edtPreset = new javax.swing.JTextField();
        btnPresetApply = new javax.swing.JButton();
        btnReqCodes = new javax.swing.JButton();
        jProgressBar1 = new javax.swing.JProgressBar();
        lblPC = new javax.swing.JLabel();
        lblMantigora = new javax.swing.JLabel();
        lblCalibrationDevices = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblUCalib = new javax.swing.JTable();
        btnLoadP = new javax.swing.JButton();
        btnSavePXLS = new javax.swing.JButton();
        btnSavePXML = new javax.swing.JButton();
        edtPcode = new javax.swing.JTextField();
        lblPCalibCode = new javax.swing.JLabel();
        lblPCalibValue = new javax.swing.JLabel();
        edtPvalue = new javax.swing.JTextField();
        lblICalibCode = new javax.swing.JLabel();
        edtIcode = new javax.swing.JTextField();
        lblICalibValue = new javax.swing.JLabel();
        edtIvalue = new javax.swing.JTextField();
        lblUCalibCode = new javax.swing.JLabel();
        lblUCalibValue = new javax.swing.JLabel();
        edtUvalue = new javax.swing.JTextField();
        edtUcode = new javax.swing.JTextField();
        jScrollPane2 = new javax.swing.JScrollPane();
        tblPCalib = new javax.swing.JTable();
        jScrollPane3 = new javax.swing.JScrollPane();
        tblICalib = new javax.swing.JTable();
        btnRemovePPoint = new javax.swing.JButton();
        btnAcceptPPoint = new javax.swing.JButton();
        btnLoadI = new javax.swing.JButton();
        btnSaveIXML = new javax.swing.JButton();
        btnSaveIXLS = new javax.swing.JButton();
        btnLoadU = new javax.swing.JButton();
        btnSaveUXML = new javax.swing.JButton();
        btnSaveUXLS = new javax.swing.JButton();
        btnRemoveIPoint = new javax.swing.JButton();
        btnAcceptIPoint = new javax.swing.JButton();
        btnRemoveUPoint = new javax.swing.JButton();
        btnAcceptUPoint = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setPreferredSize(new java.awt.Dimension(710, 560));
        setResizable(false);
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowClosing(java.awt.event.WindowEvent evt) {
                formWindowClosing(evt);
            }
        });
        getContentPane().setLayout(null);

        lblComPortTitle.setText("COM порт:");
        getContentPane().add(lblComPortTitle);
        lblComPortTitle.setBounds(10, 10, 100, 30);

        edtComPortValue.setText("/dev/ttyUSB0");
        getContentPane().add(edtComPortValue);
        edtComPortValue.setBounds(110, 10, 120, 30);

        btnConnect.setText("Соединить");
        btnConnect.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnConnectActionPerformed(evt);
            }
        });
        getContentPane().add(btnConnect);
        btnConnect.setBounds(240, 10, 220, 30);

        btnDisconnect.setText("Разъединить");
        btnDisconnect.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnDisconnectActionPerformed(evt);
            }
        });
        getContentPane().add(btnDisconnect);
        btnDisconnect.setBounds(470, 10, 220, 30);

        btnTurnOff.setText("Снять HV");
        btnTurnOff.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnTurnOffActionPerformed(evt);
            }
        });
        getContentPane().add(btnTurnOff);
        btnTurnOff.setBounds(470, 50, 220, 30);

        lblPreset.setText("Код уставки:");
        getContentPane().add(lblPreset);
        lblPreset.setBounds(10, 50, 100, 30);
        getContentPane().add(edtPreset);
        edtPreset.setBounds(110, 50, 120, 30);

        btnPresetApply.setText("Подать");
        btnPresetApply.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnPresetApplyActionPerformed(evt);
            }
        });
        getContentPane().add(btnPresetApply);
        btnPresetApply.setBounds(240, 50, 220, 30);

        btnReqCodes.setText("Получить");
        btnReqCodes.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnReqCodesActionPerformed(evt);
            }
        });
        getContentPane().add(btnReqCodes);
        btnReqCodes.setBounds(10, 90, 680, 30);

        jProgressBar1.setMaximum(5);
        getContentPane().add(jProgressBar1);
        jProgressBar1.setBounds(10, 90, 680, 30);

        lblPC.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblPC.setText("Калибровка уставки");
        getContentPane().add(lblPC);
        lblPC.setBounds(10, 130, 220, 20);

        lblMantigora.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblMantigora.setText("Калибровка токов");
        getContentPane().add(lblMantigora);
        lblMantigora.setBounds(240, 130, 220, 20);

        lblCalibrationDevices.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblCalibrationDevices.setText("Калибровка напряжений");
        getContentPane().add(lblCalibrationDevices);
        lblCalibrationDevices.setBounds(470, 130, 220, 20);

        tblUCalib.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Код", "Напряжение, В"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.Integer.class
            };
            boolean[] canEdit = new boolean [] {
                false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        tblUCalib.getTableHeader().setResizingAllowed(false);
        tblUCalib.getTableHeader().setReorderingAllowed(false);
        tblUCalib.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tblUCalibMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(tblUCalib);

        getContentPane().add(jScrollPane1);
        jScrollPane1.setBounds(470, 260, 220, 220);

        btnLoadP.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnLoadP.setText("LOAD");
        btnLoadP.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLoadPActionPerformed(evt);
            }
        });
        getContentPane().add(btnLoadP);
        btnLoadP.setBounds(10, 490, 70, 30);

        btnSavePXLS.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnSavePXLS.setText("XLSX");
        btnSavePXLS.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSavePXLSActionPerformed(evt);
            }
        });
        getContentPane().add(btnSavePXLS);
        btnSavePXLS.setBounds(90, 490, 70, 30);

        btnSavePXML.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnSavePXML.setText("XML");
        btnSavePXML.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSavePXMLActionPerformed(evt);
            }
        });
        getContentPane().add(btnSavePXML);
        btnSavePXML.setBounds(170, 490, 60, 30);

        edtPcode.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        getContentPane().add(edtPcode);
        edtPcode.setBounds(10, 180, 110, 30);

        lblPCalibCode.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblPCalibCode.setText("Уставка:");
        getContentPane().add(lblPCalibCode);
        lblPCalibCode.setBounds(10, 150, 110, 30);

        lblPCalibValue.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblPCalibValue.setText("Ток, мкА:");
        getContentPane().add(lblPCalibValue);
        lblPCalibValue.setBounds(130, 150, 100, 30);

        edtPvalue.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        edtPvalue.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyTyped(java.awt.event.KeyEvent evt) {
                edtPvalueKeyTyped(evt);
            }
        });
        getContentPane().add(edtPvalue);
        edtPvalue.setBounds(130, 180, 100, 30);

        lblICalibCode.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblICalibCode.setText("Код:");
        getContentPane().add(lblICalibCode);
        lblICalibCode.setBounds(240, 150, 110, 30);

        edtIcode.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        getContentPane().add(edtIcode);
        edtIcode.setBounds(240, 180, 110, 30);

        lblICalibValue.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblICalibValue.setText("Ток, мкА:");
        getContentPane().add(lblICalibValue);
        lblICalibValue.setBounds(360, 150, 100, 30);

        edtIvalue.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        edtIvalue.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyTyped(java.awt.event.KeyEvent evt) {
                edtIvalueKeyTyped(evt);
            }
        });
        getContentPane().add(edtIvalue);
        edtIvalue.setBounds(360, 180, 100, 30);

        lblUCalibCode.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblUCalibCode.setText("Код:");
        getContentPane().add(lblUCalibCode);
        lblUCalibCode.setBounds(470, 150, 110, 30);

        lblUCalibValue.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblUCalibValue.setText("Напряж., В:");
        getContentPane().add(lblUCalibValue);
        lblUCalibValue.setBounds(590, 150, 100, 30);

        edtUvalue.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        getContentPane().add(edtUvalue);
        edtUvalue.setBounds(590, 180, 100, 30);

        edtUcode.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        getContentPane().add(edtUcode);
        edtUcode.setBounds(470, 180, 110, 30);

        tblPCalib.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Код уставки", "Ток, мкА"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.Integer.class
            };
            boolean[] canEdit = new boolean [] {
                false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        tblPCalib.getTableHeader().setResizingAllowed(false);
        tblPCalib.getTableHeader().setReorderingAllowed(false);
        jScrollPane2.setViewportView(tblPCalib);

        getContentPane().add(jScrollPane2);
        jScrollPane2.setBounds(10, 260, 220, 220);

        tblICalib.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Код", "Ток, мкА"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Integer.class, java.lang.Integer.class
            };
            boolean[] canEdit = new boolean [] {
                false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        tblICalib.getTableHeader().setResizingAllowed(false);
        tblICalib.getTableHeader().setReorderingAllowed(false);
        jScrollPane3.setViewportView(tblICalib);

        getContentPane().add(jScrollPane3);
        jScrollPane3.setBounds(240, 260, 220, 220);

        btnRemovePPoint.setText("Удалить");
        btnRemovePPoint.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnRemovePPointActionPerformed(evt);
            }
        });
        getContentPane().add(btnRemovePPoint);
        btnRemovePPoint.setBounds(130, 220, 100, 30);

        btnAcceptPPoint.setText("Добавить");
        btnAcceptPPoint.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAcceptPPointActionPerformed(evt);
            }
        });
        getContentPane().add(btnAcceptPPoint);
        btnAcceptPPoint.setBounds(10, 220, 110, 30);

        btnLoadI.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnLoadI.setText("LOAD");
        btnLoadI.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLoadIActionPerformed(evt);
            }
        });
        getContentPane().add(btnLoadI);
        btnLoadI.setBounds(240, 490, 70, 30);

        btnSaveIXML.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnSaveIXML.setText("XML");
        btnSaveIXML.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSaveIXMLActionPerformed(evt);
            }
        });
        getContentPane().add(btnSaveIXML);
        btnSaveIXML.setBounds(400, 490, 60, 30);

        btnSaveIXLS.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnSaveIXLS.setText("XLSX");
        btnSaveIXLS.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSaveIXLSActionPerformed(evt);
            }
        });
        getContentPane().add(btnSaveIXLS);
        btnSaveIXLS.setBounds(320, 490, 70, 30);

        btnLoadU.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnLoadU.setText("LOAD");
        btnLoadU.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLoadUActionPerformed(evt);
            }
        });
        getContentPane().add(btnLoadU);
        btnLoadU.setBounds(470, 490, 70, 30);

        btnSaveUXML.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnSaveUXML.setText("XML");
        btnSaveUXML.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSaveUXMLActionPerformed(evt);
            }
        });
        getContentPane().add(btnSaveUXML);
        btnSaveUXML.setBounds(630, 490, 60, 30);

        btnSaveUXLS.setFont(new java.awt.Font("Dialog", 0, 10)); // NOI18N
        btnSaveUXLS.setText("XLSX");
        btnSaveUXLS.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSaveUXLSActionPerformed(evt);
            }
        });
        getContentPane().add(btnSaveUXLS);
        btnSaveUXLS.setBounds(550, 490, 70, 30);

        btnRemoveIPoint.setText("Удалить");
        btnRemoveIPoint.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnRemoveIPointActionPerformed(evt);
            }
        });
        getContentPane().add(btnRemoveIPoint);
        btnRemoveIPoint.setBounds(360, 220, 100, 30);

        btnAcceptIPoint.setText("Добавить");
        btnAcceptIPoint.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAcceptIPointActionPerformed(evt);
            }
        });
        getContentPane().add(btnAcceptIPoint);
        btnAcceptIPoint.setBounds(240, 220, 110, 30);

        btnRemoveUPoint.setText("Удалить");
        btnRemoveUPoint.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnRemoveUPointActionPerformed(evt);
            }
        });
        getContentPane().add(btnRemoveUPoint);
        btnRemoveUPoint.setBounds(590, 220, 100, 30);

        btnAcceptUPoint.setText("Добавить");
        btnAcceptUPoint.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAcceptUPointActionPerformed(evt);
            }
        });
        getContentPane().add(btnAcceptUPoint);
        btnAcceptUPoint.setBounds(470, 220, 110, 30);

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    public void ShowCalibrationP() {
        DefaultTableModel mdl = ( DefaultTableModel) tblPCalib.getModel();
        mdl.getDataVector().removeAllElements();
        Set set = m_pCalibrationP.entrySet();
        Iterator it = set.iterator();
        while( it.hasNext()) {
            Map.Entry entry = ( Map.Entry) it.next();
            int nCode = ( int) entry.getKey();            
            int nValue = ( int) entry.getValue();            
            mdl.addRow( new Integer[] { nCode, nValue });
        }
    }
    
    public void ShowCalibrationI() {
        DefaultTableModel mdl = ( DefaultTableModel) tblICalib.getModel();
        mdl.getDataVector().removeAllElements();
        Set set = m_pCalibrationI.entrySet();
        Iterator it = set.iterator();
        while( it.hasNext()) {
            Map.Entry entry = ( Map.Entry) it.next();
            int nCode = ( int) entry.getKey();            
            int nValue = ( int) entry.getValue();            
            mdl.addRow( new Integer[] { nCode, nValue });
        }
    }
    
    public void ShowCalibrationU() {
        DefaultTableModel mdl = ( DefaultTableModel) tblUCalib.getModel();
        mdl.getDataVector().removeAllElements();
        Set set = m_pCalibrationU.entrySet();
        Iterator it = set.iterator();
        while( it.hasNext()) {
            Map.Entry entry = ( Map.Entry) it.next();
            int nCode = ( int) entry.getKey();            
            int nValue = ( int) entry.getValue();            
            mdl.addRow( new Integer[] { nCode, nValue });
        }
    }
    
    private void btnLoadPActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadPActionPerformed
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXMLFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                            "etc" + File.separator +
                                            "calibration"));
        
        int returnVal = fc.showOpenDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            //This is where a real application would open the file.
            logger.info( "LoadFile opening: " + file.getName());
            
            TreeMap newCalib = new TreeMap();
            try {
                SAXReader reader = new SAXReader();
                URL url = file.toURI().toURL();
                Document document = reader.read( url);
                Element program = document.getRootElement();
                if( program.getName().equals( "CalibrationP")) {
                    // iterate through child elements of root
                    for ( Iterator i = program.elementIterator(); i.hasNext(); ) {
                        Element element = ( Element) i.next();
                        
                        if( element != null && "cpoint".equals( element.getName())) {
                            String strCode  = element.attributeValue( "code");
                            String strValue = element.attributeValue( "value");

                            try {
                                int nCode = Integer.parseInt( strCode);
                                int nValue = Integer.parseInt( strValue);
                                newCalib.put( nCode, nValue);
                            }
                            catch( NumberFormatException ex) {
                                logger.warn( ex);
                            }
                        }
                        else {
                            logger.warn( "not <cpoint>!");
                        }
                    }
                    
                    m_pCalibrationP = newCalib;
                    ShowCalibrationP();
                }
                else
                    logger.error( "There is no 'CalibrationP' root-tag in pointed XML");
                
                
            } catch( MalformedURLException ex) {
                logger.error( "MalformedURLException caught while loading settings!", ex);
            } catch( DocumentException ex) {
                logger.error( "DocumentException caught while loading settings!", ex);
            }
        
        } else {
            logger.info("LoadProgram cancelled.");
        }
    }//GEN-LAST:event_btnLoadPActionPerformed

    
    private void btnSavePXMLActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSavePXMLActionPerformed
        Document saveFile = DocumentHelper.createDocument();
        Element calibration = saveFile.addElement( "CalibrationP");
        
        Set set = m_pCalibrationP.entrySet();
        Iterator it = set.iterator();
        while( it.hasNext()) {
            Map.Entry entry = ( Map.Entry) it.next();
            
            int nCode  = ( int) entry.getKey();            
            int nValue = ( int) entry.getValue();            
            
            Element e = calibration.addElement( "cpoint");
            e.addAttribute( "code", "" + nCode);
            e.addAttribute( "value", "" + nValue);
        }
        
        OutputFormat format = OutputFormat.createPrettyPrint();
        
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXMLFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                            "etc" + File.separator +
                                            "calibration"));
        
        int returnVal = fc.showSaveDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            String strFilePathName = fc.getSelectedFile().getAbsolutePath();
            if( !strFilePathName.endsWith( ".xml"))
                strFilePathName += ".xml";
            File file = new File( strFilePathName);            
            XMLWriter writer;
            try {
                writer = new XMLWriter( new FileWriter( file.getAbsolutePath()), format);
                writer.write( saveFile);
                writer.close();
            } catch (IOException ex) {
                logger.error( "IOException: ", ex);
            }
        
        } else {
            logger.error( "Пользователь не указал имя файла куда сохранять программу.");
        }
    }//GEN-LAST:event_btnSavePXMLActionPerformed

    private void btnConnectActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnConnectActionPerformed
        if( edtComPortValue.getText().isEmpty()) {
            logger.info( "Connect to no-port? Ha (3 times)");
            return;
        }
        
        m_port = new SerialPort( edtComPortValue.getText());
        try {
            //Открываем порт
            m_port.openPort();

            //Выставляем параметры
            m_port.setParams( 38400,
                                 SerialPort.DATABITS_8,
                                 0,
                                 SerialPort.PARITY_NONE);
            
        }
        catch( SerialPortException ex) {
            logger.error( "COM-Communication exception", ex);
            m_bConnected = false;
            return;
        }
        m_bConnected = true;
    }//GEN-LAST:event_btnConnectActionPerformed

    private void btnDisconnectActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnDisconnectActionPerformed
        m_bConnected = false;
        try {
            m_port.closePort();
        } catch( SerialPortException ex) {
            logger.error( "COM-Communication exception", ex);
        }
    }//GEN-LAST:event_btnDisconnectActionPerformed

    private void btnReqCodesActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnReqCodesActionPerformed
        m_bGettingData = true;
        
        jProgressBar1.setValue( 0);
        edtIcode.setText( "");
        edtUcode.setText( "");
                    
        nSummCurrent = 0;
        nSummVoltage = 0;
        
        new Timer( 1000, new ActionListener() {

            @Override
            public void actionPerformed( ActionEvent e) {
                boolean bFail = false;
                
                try {
                    byte aBytes[] = new byte[1];
                    aBytes[0] = 0x05;

                    m_port.writeBytes( aBytes);
                    logger.info( "POLLING (0x05): sent");

                    byte bBytes[] = new byte[5];
                    bBytes = m_port.readBytes( 5, 800);
                
                    logger.info( String.format( "REPOSND RECEIVED: 0x%02X 0x%02X 0x%02X 0x%02X 0x%02X", bBytes[0], bBytes[1], bBytes[2], bBytes[3], bBytes[4]));
                    
                    if( bBytes[4] == 0x0D) {
                        //ответ корректный
                        int nB0 = bBytes[0] & 0xFF;
                        int nB1 = bBytes[1] & 0xFF;
                        int nB2 = bBytes[2] & 0xFF;
                        int nB3 = bBytes[3] & 0xFF;
                        
                        int nUval = ( nB2 << 8) + nB3;
                        int nIval = ( nB0 << 8) + nB1;
                        
                        nSummVoltage += nUval;
                        nSummCurrent += nIval;
                    }
                    else {
                        //а ответ пришёл корявый!
                        logger.warn( "POLLING FAILED!");
                        byte cBytes[] = new byte[ 255];
                        cBytes = m_port.readBytes();
                        bFail = true;
                    }
                } catch( SerialPortException ex) {
                    
                    logger.error( "COM-Communication exception", ex);
                    m_bGettingData = false;
                    try {
                        m_port.closePort();
                    } catch( SerialPortException ex2) {
                        logger.error( "COM-Communication exception", ex2);
                    }
                    m_bConnected = false;
                    ( ( Timer) e.getSource()).stop();
                    
                    edtIcode.setBackground( Color.red);
                    edtUcode.setBackground( Color.red);
                    new Timer( 1000, new ActionListener() {

                        @Override
                        public void actionPerformed(ActionEvent e) {
                            (( Timer) e.getSource()).stop();
                            edtIcode.setBackground( null);
                            edtUcode.setBackground( null);
                        }
                    }).start();
                    return;
                }
                catch( SerialPortTimeoutException ex) {
                    logger.error( "POLLING TIMEOUT exception", ex);
                    bFail = true;
                }
                
                if( bFail) {
                    m_bGettingData = false;
                    ( ( Timer) e.getSource()).stop();
                    
                    edtIcode.setBackground( Color.red);
                    edtUcode.setBackground( Color.red);
                    new Timer( 1000, new ActionListener() {

                        @Override
                        public void actionPerformed(ActionEvent e) {
                            (( Timer) e.getSource()).stop();
                            edtIcode.setBackground( null);
                            edtUcode.setBackground( null);
                        }
                    }).start();
                    return;
                }
                
                int nValue = jProgressBar1.getValue();
                if( ++nValue == 5) {
                    ( ( Timer) e.getSource()).stop();
                    m_bGettingData = false;
                    
                    edtIcode.setText( "" + nSummCurrent / 5);
                    edtUcode.setText( "" + nSummVoltage / 5);
                    
                    edtIcode.setBackground( Color.green);
                    edtUcode.setBackground( Color.green);
                    new Timer( 1000, new ActionListener() {

                        @Override
                        public void actionPerformed(ActionEvent e) {
                            (( Timer) e.getSource()).stop();
                            edtIcode.setBackground( null);
                            edtUcode.setBackground( null);
                        }
                    }).start();
                }
                jProgressBar1.setValue( nValue);
            }
        }).start();
    }//GEN-LAST:event_btnReqCodesActionPerformed

    private void formWindowClosing(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowClosing
        tRefreshState.stop();
        
        if( m_port != null && m_port.isOpened()) {
            try {
                m_port.closePort();
            } catch (SerialPortException ex) {
                logger.error( "COM-Communication exception", ex);
            }
        }
    }//GEN-LAST:event_formWindowClosing

    private void btnPresetApplyActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnPresetApplyActionPerformed
        int nPreset = 0;
        boolean bFail = false;
        
        try {
            nPreset = Integer.parseInt( edtPreset.getText());
            if( nPreset < 0 || nPreset > 0xFFFF)
                bFail = true;
        }
        catch( NumberFormatException ex) {
            bFail = true;
        }
            
        if( bFail) {
            edtPreset.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtPreset.setBackground( null);
                }
            }).start();
            return;
        }
            
        byte aBytes[] = new byte[3];
        aBytes[0] = 0x01;
        aBytes[1] = ( byte) ( nPreset & 0xFF);
        aBytes[2] = ( byte) ( ( nPreset & 0xFF00) >> 8);
        
        
        byte bBytes[] = new byte[1];
        bBytes[0] = 0x02;
                        
        try {
            m_port.writeBytes( aBytes);
            logger.info( String.format( "SET PRESET (0x%02X 0x%02X 0x%02X): sent", aBytes[0], aBytes[1], aBytes[2]));
            Thread.sleep( 500);
            m_port.writeBytes( bBytes);
            logger.info( String.format( "UPDATE (0x%02X): sent", bBytes[0]));
            
        } catch( SerialPortException ex) {
            logger.error( "COM-Communication exception", ex);
        } catch( InterruptedException ex) {
            logger.error( "InterruptedException", ex);
        }
        
        edtPcode.setText( edtPreset.getText());
    }//GEN-LAST:event_btnPresetApplyActionPerformed

    private void btnTurnOffActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnTurnOffActionPerformed
        byte aBytes[] = new byte[1];
        aBytes[0] = 0x03;
        try {
            m_port.writeBytes( aBytes);
            logger.info( "TURN OFF (0x03): sent");   
        } catch( SerialPortException ex) {
            logger.error( "COM-Communication exception", ex);
        }
    }//GEN-LAST:event_btnTurnOffActionPerformed

    private void btnSavePXLSActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSavePXLSActionPerformed
        String strFilePathName;
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXLSXFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                    "etc" + File.separator +
                                    "calibration"));
        
        int returnVal = fc.showSaveDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            strFilePathName = fc.getSelectedFile().getAbsolutePath();
            if( !strFilePathName.endsWith( ".xlsx"))
                strFilePathName += ".xlsx";
            
        } else {
            logger.error( "Пользователь не указал имя файла куда сохранять программу.");
            return;
        }
        
        File file = new File( strFilePathName);
        if( file.exists())
            file.delete();        
        try {

            //Create Blank workbook
            XSSFWorkbook workbook = new XSSFWorkbook(); 

            //Create a blank sheet
            XSSFSheet spreadsheet = workbook.createSheet( "TODAY_DATE");

            //Create row object
            XSSFRow row;

            /*
            //This data needs to be written (Object[])
            Map < String, Object[] > empinfo =  new TreeMap < String, Object[] >();
            empinfo.put( "1", new Object[] { "CODE", "EMP NAME", "DESIGNATION" });
            empinfo.put( "2", new Object[] { "tp01", "Gopal", "Technical Manager" });
            empinfo.put( "3", new Object[] { "tp02", "Manisha", "Proof Reader" });
            empinfo.put( "4", new Object[] { "tp03", "Masthan", "Technical Writer" });
            empinfo.put( "5", new Object[] { "tp04", "Satish", "Technical Writer" });
            empinfo.put( "6", new Object[] { "tp05", "Krishna", "Technical Writer" });
            
            
            //Iterate over data and write to sheet
            Set < String > keyid = empinfo.keySet();
            int rowid = 0;

            for (String key : keyid) {
               row = spreadsheet.createRow(rowid++);
               Object [] objectArr = empinfo.get(key);
               int cellid = 0;

               for (Object obj : objectArr) {
                  Cell cell = row.createCell(cellid++);
                  cell.setCellValue((String)obj);
               }
            }
            */
            
            int rowid = 0;
            row = spreadsheet.createRow(rowid++);
            
            Cell cell;
            cell = row.createCell( 0); cell.setCellValue( "CODE");
            cell = row.createCell( 1); cell.setCellValue( "VALUE, mcA");
            
            Set set = m_pCalibrationP.entrySet();
            Iterator it = set.iterator();
            while( it.hasNext()) {
                
                Map.Entry entry = (Map.Entry) it.next();
            
                row = spreadsheet.createRow( rowid++);
                cell = row.createCell( 0); cell.setCellValue( ( int) entry.getKey());
                cell = row.createCell( 1); cell.setCellValue( ( int) entry.getValue());
            }
            
            //Create file system using specific name
            FileOutputStream out = new FileOutputStream( new File(strFilePathName));

            //write operation workbook using file out object 
            workbook.write(out);
            out.close();

        } catch( IOException ex) {
            logger.error( "IOException exception", ex);
        }
    }//GEN-LAST:event_btnSavePXLSActionPerformed

    private void tblUCalibMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tblUCalibMouseClicked
        if( evt.getButton() == MouseEvent.BUTTON2) {
            int nSelection = tblUCalib.getSelectedRow();
            if( nSelection != -1) {
                //TODO REMOVE
            }
        }
    }//GEN-LAST:event_tblUCalibMouseClicked

    private void btnRemovePPointActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnRemovePPointActionPerformed
        int nSelection = tblPCalib.getSelectedRow();
        if( nSelection == -1) return;
        DefaultTableModel mdl = ( DefaultTableModel) tblPCalib.getModel();
        mdl.removeRow( nSelection);
    }//GEN-LAST:event_btnRemovePPointActionPerformed

    private void btnAcceptPPointActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAcceptPPointActionPerformed
        boolean bFail = false;
        Integer nCode = 0, nValue = 0;
        
        try {
            nCode = Integer.parseInt( edtPcode.getText());
        }
        catch( NumberFormatException ex) {
            bFail = true;
            edtPcode.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtPcode.setBackground( null);
                }
            }).start();
        }
        
        try {
            nValue = Integer.parseInt( edtPvalue.getText());
        }
        catch( NumberFormatException ex) {
            bFail = true;
            edtPvalue.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtPvalue.setBackground( null);
                }
            }).start();
        }
        
        if( !bFail) {
            m_pCalibrationP.put( nCode, nValue);
            ShowCalibrationP();
        }
    }//GEN-LAST:event_btnAcceptPPointActionPerformed

    private void btnLoadIActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadIActionPerformed
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXMLFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                            "etc" + File.separator +
                                            "calibration"));
        
        int returnVal = fc.showOpenDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            //This is where a real application would open the file.
            logger.info( "LoadFile opening: " + file.getName());
            
            TreeMap newCalib = new TreeMap();
            try {
                SAXReader reader = new SAXReader();
                URL url = file.toURI().toURL();
                Document document = reader.read( url);
                Element program = document.getRootElement();
                if( program.getName().equals( "CalibrationI")) {
                    // iterate through child elements of root
                    for ( Iterator i = program.elementIterator(); i.hasNext(); ) {
                        Element element = ( Element) i.next();
                        
                        if( element != null && "cpoint".equals( element.getName())) {
                            String strCode  = element.attributeValue( "code");
                            String strValue = element.attributeValue( "value");

                            try {
                                int nCode = Integer.parseInt( strCode);
                                int nValue = Integer.parseInt( strValue);
                                newCalib.put( nCode, nValue);
                            }
                            catch( NumberFormatException ex) {
                                logger.warn( ex);
                            }
                        }
                        else {
                            logger.warn( "not <cpoint>!");
                        }
                    }
                    
                    m_pCalibrationI = newCalib;
                    ShowCalibrationI();
                }
                else
                    logger.error( "There is no 'CalibrationI' root-tag in pointed XML");
                
                
            } catch( MalformedURLException ex) {
                logger.error( "MalformedURLException caught while loading settings!", ex);
            } catch( DocumentException ex) {
                logger.error( "DocumentException caught while loading settings!", ex);
            }
        
        } else {
            logger.info("LoadProgram cancelled.");
        }
    }//GEN-LAST:event_btnLoadIActionPerformed

    private void btnSaveIXMLActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSaveIXMLActionPerformed
        Document saveFile = DocumentHelper.createDocument();
        Element calibration = saveFile.addElement( "CalibrationI");
        
        Set set = m_pCalibrationI.entrySet();
        Iterator it = set.iterator();
        while( it.hasNext()) {
            Map.Entry entry = ( Map.Entry) it.next();
            
            int nCode  = ( int) entry.getKey();            
            int nValue = ( int) entry.getValue();            
            
            Element e = calibration.addElement( "cpoint");
            e.addAttribute( "code", "" + nCode);
            e.addAttribute( "value", "" + nValue);
        }
        
        OutputFormat format = OutputFormat.createPrettyPrint();
        
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXMLFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                            "etc" + File.separator +
                                            "calibration"));
        
        int returnVal = fc.showSaveDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            String strFilePathName = fc.getSelectedFile().getAbsolutePath();
            if( !strFilePathName.endsWith( ".xml"))
                strFilePathName += ".xml";
            File file = new File( strFilePathName);            
            XMLWriter writer;
            try {
                writer = new XMLWriter( new FileWriter( file.getAbsolutePath()), format);
                writer.write( saveFile);
                writer.close();
            } catch (IOException ex) {
                logger.error( "IOException: ", ex);
            }
        
        } else {
            logger.error( "Пользователь не указал имя файла куда сохранять программу.");
        }
    }//GEN-LAST:event_btnSaveIXMLActionPerformed

    private void btnSaveIXLSActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSaveIXLSActionPerformed
        String strFilePathName;
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXLSXFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                    "etc" + File.separator +
                                    "calibration"));
        
        int returnVal = fc.showSaveDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            strFilePathName = fc.getSelectedFile().getAbsolutePath();
            if( !strFilePathName.endsWith( ".xlsx"))
                strFilePathName += ".xlsx";
            
        } else {
            logger.error( "Пользователь не указал имя файла куда сохранять программу.");
            return;
        }
        
        File file = new File( strFilePathName);
        if( file.exists())
            file.delete();        
        try {

            //Create Blank workbook
            XSSFWorkbook workbook = new XSSFWorkbook(); 

            //Create a blank sheet
            XSSFSheet spreadsheet = workbook.createSheet( "TODAY_DATE");

            //Create row object
            XSSFRow row;

            /*
            //This data needs to be written (Object[])
            Map < String, Object[] > empinfo =  new TreeMap < String, Object[] >();
            empinfo.put( "1", new Object[] { "CODE", "EMP NAME", "DESIGNATION" });
            empinfo.put( "2", new Object[] { "tp01", "Gopal", "Technical Manager" });
            empinfo.put( "3", new Object[] { "tp02", "Manisha", "Proof Reader" });
            empinfo.put( "4", new Object[] { "tp03", "Masthan", "Technical Writer" });
            empinfo.put( "5", new Object[] { "tp04", "Satish", "Technical Writer" });
            empinfo.put( "6", new Object[] { "tp05", "Krishna", "Technical Writer" });
            
            
            //Iterate over data and write to sheet
            Set < String > keyid = empinfo.keySet();
            int rowid = 0;

            for (String key : keyid) {
               row = spreadsheet.createRow(rowid++);
               Object [] objectArr = empinfo.get(key);
               int cellid = 0;

               for (Object obj : objectArr) {
                  Cell cell = row.createCell(cellid++);
                  cell.setCellValue((String)obj);
               }
            }
            */
            
            int rowid = 0;
            row = spreadsheet.createRow(rowid++);
            
            Cell cell;
            cell = row.createCell( 0); cell.setCellValue( "CODE");
            cell = row.createCell( 1); cell.setCellValue( "VALUE, mcA");
            
            Set set = m_pCalibrationI.entrySet();
            Iterator it = set.iterator();
            while( it.hasNext()) {
                
                Map.Entry entry = (Map.Entry) it.next();
            
                row = spreadsheet.createRow( rowid++);
                cell = row.createCell( 0); cell.setCellValue( ( int) entry.getKey());
                cell = row.createCell( 1); cell.setCellValue( ( int) entry.getValue());
            }
            
            //Create file system using specific name
            FileOutputStream out = new FileOutputStream( new File(strFilePathName));

            //write operation workbook using file out object 
            workbook.write(out);
            out.close();

        } catch( IOException ex) {
            logger.error( "IOException exception", ex);
        }
    }//GEN-LAST:event_btnSaveIXLSActionPerformed

    private void btnLoadUActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLoadUActionPerformed
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXMLFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                            "etc" + File.separator +
                                            "calibration"));
        
        int returnVal = fc.showOpenDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            //This is where a real application would open the file.
            logger.info( "LoadFile opening: " + file.getName());
            
            TreeMap newCalib = new TreeMap();
            try {
                SAXReader reader = new SAXReader();
                URL url = file.toURI().toURL();
                Document document = reader.read( url);
                Element program = document.getRootElement();
                if( program.getName().equals( "CalibrationU")) {
                    // iterate through child elements of root
                    for ( Iterator i = program.elementIterator(); i.hasNext(); ) {
                        Element element = ( Element) i.next();
                        
                        if( element != null && "cpoint".equals( element.getName())) {
                            String strCode  = element.attributeValue( "code");
                            String strValue = element.attributeValue( "value");

                            try {
                                int nCode = Integer.parseInt( strCode);
                                int nValue = Integer.parseInt( strValue);
                                newCalib.put( nCode, nValue);
                            }
                            catch( NumberFormatException ex) {
                                logger.warn( ex);
                            }
                        }
                        else {
                            logger.warn( "not <cpoint>!");
                        }
                    }
                    
                    m_pCalibrationU = newCalib;
                    ShowCalibrationU();
                }
                else
                    logger.error( "There is no 'CalibrationU' root-tag in pointed XML");
                
                
            } catch( MalformedURLException ex) {
                logger.error( "MalformedURLException caught while loading settings!", ex);
            } catch( DocumentException ex) {
                logger.error( "DocumentException caught while loading settings!", ex);
            }
        
        } else {
            logger.info("LoadProgram cancelled.");
        }
    }//GEN-LAST:event_btnLoadUActionPerformed

    private void btnSaveUXMLActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSaveUXMLActionPerformed
        Document saveFile = DocumentHelper.createDocument();
        Element calibration = saveFile.addElement( "CalibrationU");
        
        Set set = m_pCalibrationU.entrySet();
        Iterator it = set.iterator();
        while( it.hasNext()) {
            Map.Entry entry = ( Map.Entry) it.next();
            
            int nCode  = ( int) entry.getKey();            
            int nValue = ( int) entry.getValue();            
            
            Element e = calibration.addElement( "cpoint");
            e.addAttribute( "code", "" + nCode);
            e.addAttribute( "value", "" + nValue);
        }
        
        OutputFormat format = OutputFormat.createPrettyPrint();
        
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXMLFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                            "etc" + File.separator +
                                            "calibration"));
        
        int returnVal = fc.showSaveDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            String strFilePathName = fc.getSelectedFile().getAbsolutePath();
            if( !strFilePathName.endsWith( ".xml"))
                strFilePathName += ".xml";
            File file = new File( strFilePathName);            
            XMLWriter writer;
            try {
                writer = new XMLWriter( new FileWriter( file.getAbsolutePath()), format);
                writer.write( saveFile);
                writer.close();
            } catch (IOException ex) {
                logger.error( "IOException: ", ex);
            }
        
        } else {
            logger.error( "Пользователь не указал имя файла куда сохранять программу.");
        }
    }//GEN-LAST:event_btnSaveUXMLActionPerformed

    private void btnSaveUXLSActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSaveUXLSActionPerformed
        String strFilePathName;
        final JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new JFileDialogXLSXFilter());
        fc.setCurrentDirectory( new File( GetAMSRoot() + File.separator +
                                    "etc" + File.separator +
                                    "calibration"));
        
        int returnVal = fc.showSaveDialog( this);
        if( returnVal == JFileChooser.APPROVE_OPTION) {
            strFilePathName = fc.getSelectedFile().getAbsolutePath();
            if( !strFilePathName.endsWith( ".xlsx"))
                strFilePathName += ".xlsx";
            
        } else {
            logger.error( "Пользователь не указал имя файла куда сохранять программу.");
            return;
        }
        
        File file = new File( strFilePathName);
        if( file.exists())
            file.delete();        
        try {

            //Create Blank workbook
            XSSFWorkbook workbook = new XSSFWorkbook(); 

            //Create a blank sheet
            XSSFSheet spreadsheet = workbook.createSheet( "TODAY_DATE");

            //Create row object
            XSSFRow row;

            /*
            //This data needs to be written (Object[])
            Map < String, Object[] > empinfo =  new TreeMap < String, Object[] >();
            empinfo.put( "1", new Object[] { "CODE", "EMP NAME", "DESIGNATION" });
            empinfo.put( "2", new Object[] { "tp01", "Gopal", "Technical Manager" });
            empinfo.put( "3", new Object[] { "tp02", "Manisha", "Proof Reader" });
            empinfo.put( "4", new Object[] { "tp03", "Masthan", "Technical Writer" });
            empinfo.put( "5", new Object[] { "tp04", "Satish", "Technical Writer" });
            empinfo.put( "6", new Object[] { "tp05", "Krishna", "Technical Writer" });
            
            
            //Iterate over data and write to sheet
            Set < String > keyid = empinfo.keySet();
            int rowid = 0;

            for (String key : keyid) {
               row = spreadsheet.createRow(rowid++);
               Object [] objectArr = empinfo.get(key);
               int cellid = 0;

               for (Object obj : objectArr) {
                  Cell cell = row.createCell(cellid++);
                  cell.setCellValue((String)obj);
               }
            }
            */
            
            int rowid = 0;
            row = spreadsheet.createRow(rowid++);
            
            Cell cell;
            cell = row.createCell( 0); cell.setCellValue( "CODE");
            cell = row.createCell( 1); cell.setCellValue( "VALUE, V");
            
            Set set = m_pCalibrationU.entrySet();
            Iterator it = set.iterator();
            while( it.hasNext()) {
                
                Map.Entry entry = (Map.Entry) it.next();
            
                row = spreadsheet.createRow( rowid++);
                cell = row.createCell( 0); cell.setCellValue( ( int) entry.getKey());
                cell = row.createCell( 1); cell.setCellValue( ( int) entry.getValue());
            }
            
            //Create file system using specific name
            FileOutputStream out = new FileOutputStream( new File(strFilePathName));

            //write operation workbook using file out object 
            workbook.write(out);
            out.close();

        } catch( IOException ex) {
            logger.error( "IOException exception", ex);
        }
    }//GEN-LAST:event_btnSaveUXLSActionPerformed

    private void btnRemoveIPointActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnRemoveIPointActionPerformed
        int nSelection = tblICalib.getSelectedRow();
        if( nSelection == -1) return;
        DefaultTableModel mdl = ( DefaultTableModel) tblICalib.getModel();
        mdl.removeRow( nSelection);
    }//GEN-LAST:event_btnRemoveIPointActionPerformed

    private void btnAcceptIPointActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAcceptIPointActionPerformed
        boolean bFail = false;
        Integer nCode = 0, nValue = 0;
        
        try {
            nCode = Integer.parseInt( edtIcode.getText());
        }
        catch( NumberFormatException ex) {
            bFail = true;
            edtIcode.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtIcode.setBackground( null);
                }
            }).start();
        }
        
        try {
            nValue = Integer.parseInt( edtIvalue.getText());
        }
        catch( NumberFormatException ex) {
            bFail = true;
            edtIvalue.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtIvalue.setBackground( null);
                }
            }).start();
        }
        
        if( !bFail) {
            m_pCalibrationI.put( nCode, nValue);
            ShowCalibrationI();
        }
    }//GEN-LAST:event_btnAcceptIPointActionPerformed

    private void btnRemoveUPointActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnRemoveUPointActionPerformed
        int nSelection = tblUCalib.getSelectedRow();
        if( nSelection == -1) return;
        DefaultTableModel mdl = ( DefaultTableModel) tblUCalib.getModel();
        mdl.removeRow( nSelection);
    }//GEN-LAST:event_btnRemoveUPointActionPerformed

    private void btnAcceptUPointActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAcceptUPointActionPerformed
        boolean bFail = false;
        Integer nCode = 0, nValue = 0;
        
        try {
            nCode = Integer.parseInt( edtUcode.getText());
        }
        catch( NumberFormatException ex) {
            bFail = true;
            edtUcode.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtUcode.setBackground( null);
                }
            }).start();
        }
        
        try {
            nValue = Integer.parseInt( edtUvalue.getText());
        }
        catch( NumberFormatException ex) {
            bFail = true;
            edtUvalue.setBackground( Color.red);
            new Timer( 1000, new ActionListener() {
                @Override
                public void actionPerformed( ActionEvent e) {
                    ( ( Timer) e.getSource()).stop();
                    edtUvalue.setBackground( null);
                }
            }).start();
        }
        
        if( !bFail) {
            m_pCalibrationU.put( nCode, nValue);
            ShowCalibrationU();
        }
    }//GEN-LAST:event_btnAcceptUPointActionPerformed

    private void edtPvalueKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_edtPvalueKeyTyped
       new Timer( 10, new ActionListener() {
            
            @Override
            public void actionPerformed(ActionEvent e) {
                Timer tt = ( Timer) e.getSource();
                tt.stop();
                edtIvalue.setText( edtPvalue.getText());
            }
        }).start();
    }//GEN-LAST:event_edtPvalueKeyTyped

    private void edtIvalueKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_edtIvalueKeyTyped
        new Timer( 10, new ActionListener() {
            
            @Override
            public void actionPerformed(ActionEvent e) {
                Timer tt = ( Timer) e.getSource();
                tt.stop();
                edtPvalue.setText( edtIvalue.getText());
            }
        }).start();
    }//GEN-LAST:event_edtIvalueKeyTyped

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(HVV4HvCalibration.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(HVV4HvCalibration.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(HVV4HvCalibration.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(HVV4HvCalibration.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        //главная переменная окружения
        String strAMSrootEnvVar = System.getenv( "AMS_ROOT");
        if( strAMSrootEnvVar == null) {
            MessageBoxError( "Не задана переменная окружения AMS_ROOT!", "HVV4.CALIB");
            return;
        }
        
        //настройка логгера
        String strlog4jPropertiesFile = strAMSrootEnvVar + File.separator +
                                        "etc" + File.separator +
                                        "log4j" + File.separator +
                                        "log4j.hvv4.hv.calib.properties";
        File file = new File( strlog4jPropertiesFile);
        if(!file.exists())
            System.out.println("It is not possible to load the given log4j properties file :" + file.getAbsolutePath());
        else
            PropertyConfigurator.configure( file.getAbsolutePath());

        logger.info( "");
        logger.info( "");
        logger.info( "");
        logger.info( "******");
        logger.info( "******");
        logger.info( "******");
        
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new HVV4HvCalibration().setVisible(true);
            }
        });        
    }

    /**
     * Функция для сообщения пользователю информационного сообщения
     * @param strMessage сообщение
     * @param strTitleBar заголовок
     */
    public static void MessageBoxInfo( String strMessage, String strTitleBar)
    {
        JOptionPane.showMessageDialog( null, strMessage, strTitleBar, JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * Функция для сообщения пользователю сообщения об ошибке
     * @param strMessage сообщение
     * @param strTitleBar заголовок
     */
    public static void MessageBoxError( String strMessage, String strTitleBar)
    {
        JOptionPane.showMessageDialog( null, strMessage, strTitleBar, JOptionPane.ERROR_MESSAGE);
    }
    
    public Date GetLocalDate() {
        Date dt = new Date( System.currentTimeMillis() - 1000 * 60 * 60 * -1);//GetSettings().GetTimeZoneShift());
        return dt;
    }
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnAcceptIPoint;
    private javax.swing.JButton btnAcceptPPoint;
    private javax.swing.JButton btnAcceptUPoint;
    private javax.swing.JButton btnConnect;
    private javax.swing.JButton btnDisconnect;
    private javax.swing.JButton btnLoadI;
    private javax.swing.JButton btnLoadP;
    private javax.swing.JButton btnLoadU;
    private javax.swing.JButton btnPresetApply;
    private javax.swing.JButton btnRemoveIPoint;
    private javax.swing.JButton btnRemovePPoint;
    private javax.swing.JButton btnRemoveUPoint;
    private javax.swing.JButton btnReqCodes;
    private javax.swing.JButton btnSaveIXLS;
    private javax.swing.JButton btnSaveIXML;
    private javax.swing.JButton btnSavePXLS;
    private javax.swing.JButton btnSavePXML;
    private javax.swing.JButton btnSaveUXLS;
    private javax.swing.JButton btnSaveUXML;
    private javax.swing.JButton btnTurnOff;
    private javax.swing.JTextField edtComPortValue;
    private javax.swing.JTextField edtIcode;
    private javax.swing.JTextField edtIvalue;
    private javax.swing.JTextField edtPcode;
    private javax.swing.JTextField edtPreset;
    private javax.swing.JTextField edtPvalue;
    private javax.swing.JTextField edtUcode;
    private javax.swing.JTextField edtUvalue;
    private javax.swing.JProgressBar jProgressBar1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JLabel lblCalibrationDevices;
    private javax.swing.JLabel lblComPortTitle;
    private javax.swing.JLabel lblICalibCode;
    private javax.swing.JLabel lblICalibValue;
    private javax.swing.JLabel lblMantigora;
    private javax.swing.JLabel lblPC;
    private javax.swing.JLabel lblPCalibCode;
    private javax.swing.JLabel lblPCalibValue;
    private javax.swing.JLabel lblPreset;
    private javax.swing.JLabel lblUCalibCode;
    private javax.swing.JLabel lblUCalibValue;
    private javax.swing.JTable tblICalib;
    private javax.swing.JTable tblPCalib;
    private javax.swing.JTable tblUCalib;
    // End of variables declaration//GEN-END:variables
}
