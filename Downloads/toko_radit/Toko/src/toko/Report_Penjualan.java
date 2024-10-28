/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package toko;

import java.awt.Graphics;
import java.awt.Graphics2D;
import java.sql.*;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;

import java.awt.print.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.Font;

/**
 *
 * @author winda
 */
public class Report_Penjualan extends javax.swing.JFrame {

    Connection conn = Koneksi.GetConnection();
    String sql;
    ResultSet rs;

    /**
     * Creates new form Report_Penjualan
     */
    public Report_Penjualan() {
        initComponents();
        setLocationRelativeTo(this);
        GetData();
    }

    public void exportToExcel(JTable table) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data Sales");

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < table.getColumnCount(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(table.getColumnName(i));
        }

        for (int i = 0; i < table.getRowCount(); i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < table.getColumnCount(); j++) {
                Cell cell = row.createCell(j);
                Object value = table.getValueAt(i, j);
                if (value instanceof Number) {
                    cell.setCellValue(((Number) value).doubleValue());
                    cell.setCellType(CellType.NUMERIC);
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                    cell.setCellType(CellType.BOOLEAN);
                } else {
                    cell.setCellValue(value.toString());
                    cell.setCellType(CellType.STRING); 
                }
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream("Data_Sales.xlsx")) {
            workbook.write(outputStream);
            JOptionPane.showMessageDialog(null, "Data berhasil diekspor ke Excel.");
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error saat menyimpan file: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void printTable(JTable table) {
        try {
            PrinterJob printerJob = PrinterJob.getPrinterJob();

            printerJob.setPrintable(new Printable() {
                public int print(Graphics graphics, PageFormat pageFormat, int pageIndex) throws PrinterException {
                    if (pageIndex > 0) {
                        return NO_SUCH_PAGE; 
                    }

                    Graphics2D g2d = (Graphics2D) graphics;

                    double margin = 36; 
                    double pageWidth = pageFormat.getImageableWidth();
                    double pageHeight = pageFormat.getImageableHeight();

                    g2d.translate(pageFormat.getImageableX() + margin, pageFormat.getImageableY() + margin);

                    g2d.setFont(new Font("Arial", Font.BOLD, 16)); 
                    g2d.drawString("Data Sales Report", 0, 0); 

                    Font headerFont = new Font("Arial", Font.BOLD, 12);
                    Font dataFont = new Font("Arial", Font.PLAIN, 10);
                    g2d.setFont(headerFont);

                    int[] columnWidths = new int[table.getColumnCount()];
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        columnWidths[i] = g2d.getFontMetrics().stringWidth(table.getColumnName(i)) + 10; 
                    }

                    for (int i = 0; i < table.getRowCount(); i++) {
                        for (int j = 0; j < table.getColumnCount(); j++) {
                            String cellValue = String.valueOf(table.getValueAt(i, j));
                            int cellWidth = g2d.getFontMetrics().stringWidth(cellValue) + 10;
                            columnWidths[j] = Math.max(columnWidths[j], cellWidth);
                        }
                    }

                    int headerYPosition = 30;
                    int currentX = 0; 
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        g2d.drawString(table.getColumnName(i), currentX, headerYPosition);
                        currentX += columnWidths[i];
                    }

                    g2d.setFont(dataFont);
                    int rowHeight = 15;
                    for (int i = 0; i < table.getRowCount(); i++) {
                        currentX = 0; 
                        for (int j = 0; j < table.getColumnCount(); j++) {
                            String cellValue = String.valueOf(table.getValueAt(i, j));
                            g2d.drawString(cellValue, currentX, headerYPosition + (i + 1) * rowHeight);
                            currentX += columnWidths[j]; 
                        }
                    }

                    return PAGE_EXISTS; 
                }
            });

            if (printerJob.printDialog()) {
                printerJob.print(); 
            }
        } catch (PrinterException e) {
            JOptionPane.showMessageDialog(null, "Error during printing: " + e.getMessage());
        }
    }

    public void GetData() {
        try {
            sql = "SELECT * FROM tb_sales";
            PreparedStatement ps = conn.prepareStatement(sql);
            rs = ps.executeQuery();

            DefaultTableModel tbl = new DefaultTableModel();
            tbl.addColumn("No");
            tbl.addColumn("Kode Faktur");
            tbl.addColumn("Id Barang");
            tbl.addColumn("Nama Barang");
            tbl.addColumn("Harga Satuan");
            tbl.addColumn("Jumlah");
            tbl.addColumn("Total Harga");
            tbl.addColumn("Status");

            int i = 0;
            while (rs.next()) {
                i++;
                String kodeFaktur = rs.getString("kode_faktur");
                int idBarang = rs.getInt("id_barang");
                String namaBarang = rs.getString("nama_barang");
                int hargaSatuan = rs.getInt("harga_barang"); 
                int jumlah = rs.getInt("jumlah"); 
                int totalHarga = rs.getInt("total"); 
                String status = rs.getString("status_pembayaran");

                tbl.addRow(new Object[]{
                    i,
                    kodeFaktur,
                    idBarang,
                    namaBarang,
                    hargaSatuan, 
                    jumlah, 
                    totalHarga, 
                    status
                });
            }

            tbl_report.setModel(tbl);

        } catch (Exception e) {
            e.printStackTrace(); 
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane2 = new javax.swing.JScrollPane();
        jPanel1 = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jScrollPane3 = new javax.swing.JScrollPane();
        tbl_report = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenu2 = new javax.swing.JMenu();
        jMenu3 = new javax.swing.JMenu();
        jMenu4 = new javax.swing.JMenu();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jButton1.setBackground(new java.awt.Color(153, 153, 153));
        jButton1.setFont(new java.awt.Font("Rockwell Condensed", 1, 14)); // NOI18N
        jButton1.setForeground(new java.awt.Color(51, 255, 0));
        jButton1.setText("Export");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton1, new org.netbeans.lib.awtextra.AbsoluteConstraints(650, 280, 90, 40));

        jButton2.setBackground(new java.awt.Color(153, 153, 153));
        jButton2.setFont(new java.awt.Font("Rockwell Condensed", 1, 12)); // NOI18N
        jButton2.setForeground(new java.awt.Color(255, 204, 0));
        jButton2.setText("Cetak");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton2, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 280, 90, 40));

        tbl_report.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane3.setViewportView(tbl_report);

        jPanel1.add(jScrollPane3, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 50, 750, 220));

        jLabel1.setBackground(new java.awt.Color(204, 204, 255));
        jLabel1.setFont(new java.awt.Font("Rockwell Condensed", 1, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(51, 204, 0));
        jLabel1.setText("LAPORAN  PENJUALAN");
        jPanel1.add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 10, -1, -1));

        jMenu1.setText("Data Barang");
        jMenu1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jMenu1MouseClicked(evt);
            }
        });
        jMenuBar1.add(jMenu1);

        jMenu2.setText("Transaksi");
        jMenu2.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jMenu2MouseClicked(evt);
            }
        });
        jMenuBar1.add(jMenu2);

        jMenu3.setText("Inventaris");
        jMenu3.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jMenu3MouseClicked(evt);
            }
        });
        jMenuBar1.add(jMenu3);

        jMenu4.setText("Report");
        jMenu4.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jMenu4MouseClicked(evt);
            }
        });
        jMenuBar1.add(jMenu4);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, 328, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        exportToExcel(tbl_report);
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jMenu1MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jMenu1MouseClicked
        // TODO add your handling code here:
        this.dispose();
        new DataBarang().show();
    }//GEN-LAST:event_jMenu1MouseClicked

    private void jMenu2MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jMenu2MouseClicked
        // TODO add your handling code here:
        this.dispose();
        new Form_Transaksi().show();
    }//GEN-LAST:event_jMenu2MouseClicked

    private void jMenu3MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jMenu3MouseClicked
        // TODO add your handling code here:
        this.dispose();
        new Inventaris().show();
    }//GEN-LAST:event_jMenu3MouseClicked

    private void jMenu4MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jMenu4MouseClicked
        // TODO add your handling code here:
        this.dispose();
        new Report_Penjualan().show();
    }//GEN-LAST:event_jMenu4MouseClicked

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        // TODO add your handling code here:
        printTable(tbl_report);
    }//GEN-LAST:event_jButton2ActionPerformed

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
            java.util.logging.Logger.getLogger(Report_Penjualan.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Report_Penjualan.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Report_Penjualan.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Report_Penjualan.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Report_Penjualan().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenu jMenu4;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JTable tbl_report;
    // End of variables declaration//GEN-END:variables
}
