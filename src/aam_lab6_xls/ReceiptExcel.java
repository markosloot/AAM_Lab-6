package aam_lab6_xls;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.ImageIcon;
import javax.swing.JLabel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReceiptExcel extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread { // Поток запуска MS Excel

        public void run() {
            
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Текущий катаолог
            try {
                modifData(dir + "receipt_template.xls", dir + "receipt.xls", jTextField_FIO.getText(),
                        jTextField_Date.getText(), jTextField_Adres.getText(), jTextField_Any.getText()); // Вызов метода создания отчета
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "receipt.xls").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "receipt.xls")); // Запуск отчета в MS Excel
                }
            } catch (Exception ex) {
                System.err.println("Error modifData!");
                ex.printStackTrace();
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    // Метод создания отчета
    private void modifData(String inputFileName, String outputFileName, String fio, String date,
            String adres, String any) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFileName))); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Первый лист в документе MS Excel
        sheet.getRow(13).getCell(1).setCellValue(fio);
        sheet.getRow(16).getCell(2).setCellValue(date);
        sheet.getRow(18).getCell(1).setCellValue(adres);
        sheet.getRow(23).getCell(1).setCellValue(any);
        
        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public ReceiptExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton1 = new javax.swing.JButton();
        jTextField_FIO = new javax.swing.JTextField();
        jTextField_Date = new javax.swing.JTextField();
        jTextField_Adres = new javax.swing.JTextField();
        jTextField_Any = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Работа с Excel");
        setResizable(false);
        getContentPane().setLayout(null);

        jButton1.setText("в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(150, 370, 66, 22);

        jTextField_FIO.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_FIO);
        jTextField_FIO.setBounds(50, 160, 280, 20);

        jTextField_Date.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Date);
        jTextField_Date.setBounds(110, 200, 160, 20);

        jTextField_Adres.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Adres);
        jTextField_Adres.setBounds(60, 230, 260, 20);

        jTextField_Any.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Any);
        jTextField_Any.setBounds(50, 310, 280, 30);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/aam_lab6_xls/screen.png"))); // NOI18N
        getContentPane().add(jLabel2);
        jLabel2.setBounds(0, 0, 380, 420);

        setSize(new java.awt.Dimension(396, 463));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton1ActionPerformed

    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        
        
        //</editor-fold>
        

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ReceiptExcel().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JTextField jTextField_Adres;
    private javax.swing.JTextField jTextField_Any;
    private javax.swing.JTextField jTextField_Date;
    private javax.swing.JTextField jTextField_FIO;
    // End of variables declaration//GEN-END:variables
}
