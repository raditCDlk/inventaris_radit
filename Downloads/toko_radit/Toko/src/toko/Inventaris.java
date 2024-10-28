package toko;

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
import java.sql.*;
import java.text.NumberFormat;
import java.util.Locale;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;

/**
 *
 * @author winda
 */
public class Inventaris extends javax.swing.JFrame {

    Connection conn = Koneksi.GetConnection();
    ResultSet rs;

    /**
     * Creates new form Inventaris
     */
    public Inventaris() {
        initComponents();
        setLocationRelativeTo(this);
        GetData();
    }

    public static String formatRupiah(double amount) {
        // Buat NumberFormat untuk format Rupiah
        NumberFormat numberFormat = NumberFormat.getCurrencyInstance(new Locale("id", "ID"));
        return numberFormat.format(amount);
    }

    public void GetData() {
        try {
            String cbvalue = cb_barang.getSelectedItem().toString();
            String orderBy = "";

            if (cbvalue.equals("ID_Barang")) {
                orderBy = "id";
            } else if (cbvalue.equals("Nama_Barang")) {
                orderBy = "nama";
            }

            String sql = "select * from tb_barang order by " + orderBy + " asc";

            PreparedStatement ps = conn.prepareStatement(sql);
            rs = ps.executeQuery();

            tbl_barang.setModel(new DefaultTableModel());

            DefaultTableModel tbl = new DefaultTableModel();
            tbl.addColumn("No");
            tbl.addColumn("Id Barang");
            tbl.addColumn("Nama Barang");
            tbl.addColumn("Merk");
            tbl.addColumn("Stok");
            tbl.addColumn("Harga");
            tbl.addColumn("Tanggal Masuk");
            tbl.addColumn("Expired");

            tbl_barang.setModel(tbl);
            int i = 0;
            while (rs.next()) {
                i++;
                tbl.addRow(new Object[]{
                    i,
                    rs.getString("id"),
                    rs.getString("nama"),
                    rs.getString("merk"),
                    rs.getString("stok"),
                    rs.getString("harga"),
                    rs.getString("tgl_masuk"),
                    rs.getString("exp"),});
            }
            tbl_barang.setModel(tbl);

        } catch (Exception e) {
        }

    }

    public void SearchData(String keyword) {
        try {
            String cbvalue = cb_barang.getSelectedItem().toString();
            String searchColumn = "";

          
            if (cbvalue.equals("ID_Barang")) {
                searchColumn = "id";
            } else if (cbvalue.equals("Nama_Barang")) {
                searchColumn = "nama";
            }

            
            String sql = "SELECT * FROM tb_barang WHERE " + searchColumn + " LIKE ? ORDER BY " + searchColumn + " ASC";

            PreparedStatement ps = conn.prepareStatement(sql);

           
            ps.setString(1, "%" + keyword + "%"); 

            rs = ps.executeQuery();

           
            tbl_barang.setModel(new DefaultTableModel());

           
            DefaultTableModel tbl = new DefaultTableModel();
            tbl.addColumn("No");
            tbl.addColumn("Id Barang");
            tbl.addColumn("Nama Barang");
            tbl.addColumn("Merk");
            tbl.addColumn("Stok");
            tbl.addColumn("Harga");
            tbl.addColumn("Tanggal Masuk");
            tbl.addColumn("Expired");

            tbl_barang.setModel(tbl);
            int i = 0;

            
            while (rs.next()) {
                i++;
                tbl.addRow(new Object[]{
                    i,
                    rs.getString("id"),
                    rs.getString("nama"),
                    rs.getString("merk"),
                    rs.getString("stok"),
                    rs.getString("harga"),
                    rs.getString("tgl_masuk"),
                    rs.getString("exp"),});
            }
            tbl_barang.setModel(tbl);

        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Error: " + e.getMessage());
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

        jPanel1 = new javax.swing.JPanel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        txt_harga = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        cb_barang = new javax.swing.JComboBox<>();
        txt_search = new javax.swing.JTextField();
        txt_nama = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        txt_stok = new javax.swing.JLabel();
        txt_exp = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tbl_barang = new javax.swing.JTable();
        jButton1 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenu2 = new javax.swing.JMenu();
        jMenu3 = new javax.swing.JMenu();
        jMenu4 = new javax.swing.JMenu();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel2.setText("Cari Barang       :");
        jPanel1.add(jLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 40, -1, -1));

        jLabel3.setText("Nama Barang    :");
        jPanel1.add(jLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 90, -1, -1));

        txt_harga.setText("Rp-");
        jPanel1.add(txt_harga, new org.netbeans.lib.awtextra.AbsoluteConstraints(630, 40, -1, -1));

        jLabel5.setText("Stok Barang   :");
        jPanel1.add(jLabel5, new org.netbeans.lib.awtextra.AbsoluteConstraints(370, 40, -1, -1));

        cb_barang.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "ID_Barang", "Nama_Barang" }));
        cb_barang.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                cb_barangItemStateChanged(evt);
            }
        });
        cb_barang.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyTyped(java.awt.event.KeyEvent evt) {
                cb_barangKeyTyped(evt);
            }
        });
        jPanel1.add(cb_barang, new org.netbeans.lib.awtextra.AbsoluteConstraints(100, 40, 100, -1));
        jPanel1.add(txt_search, new org.netbeans.lib.awtextra.AbsoluteConstraints(204, 40, 80, -1));

        txt_nama.setText("-");
        jPanel1.add(txt_nama, new org.netbeans.lib.awtextra.AbsoluteConstraints(100, 90, 170, 20));

        jLabel7.setText("Exp.                 :");
        jPanel1.add(jLabel7, new org.netbeans.lib.awtextra.AbsoluteConstraints(370, 90, -1, -1));

        txt_stok.setText("-");
        jPanel1.add(txt_stok, new org.netbeans.lib.awtextra.AbsoluteConstraints(450, 40, 50, 20));

        txt_exp.setText("-");
        jPanel1.add(txt_exp, new org.netbeans.lib.awtextra.AbsoluteConstraints(450, 90, 110, 20));

        jLabel10.setText("Harga Barang   :");
        jPanel1.add(jLabel10, new org.netbeans.lib.awtextra.AbsoluteConstraints(530, 40, -1, -1));

        tbl_barang.setModel(new javax.swing.table.DefaultTableModel(
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
        tbl_barang.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tbl_barangMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(tbl_barang);

        jPanel1.add(jScrollPane1, new org.netbeans.lib.awtextra.AbsoluteConstraints(80, 140, 660, 210));

        jButton1.setBackground(new java.awt.Color(153, 153, 153));
        jButton1.setFont(new java.awt.Font("Rockwell Condensed", 1, 12)); // NOI18N
        jButton1.setForeground(new java.awt.Color(255, 204, 0));
        jButton1.setText("Cari");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton1, new org.netbeans.lib.awtextra.AbsoluteConstraints(300, 40, -1, -1));

        jLabel1.setBackground(new java.awt.Color(204, 204, 255));
        jLabel1.setFont(new java.awt.Font("Rockwell Condensed", 1, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 204, 0));
        jLabel1.setText("INVENTARIS TOKO");
        jPanel1.add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 0, 190, 30));

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
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, 782, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, 371, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

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

    private void cb_barangKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_cb_barangKeyTyped
        // TODO add your handling code here:
    }//GEN-LAST:event_cb_barangKeyTyped

    private void cb_barangItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_cb_barangItemStateChanged
        // TODO add your handling code here:
        GetData();
    }//GEN-LAST:event_cb_barangItemStateChanged

    private void tbl_barangMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tbl_barangMouseClicked
        // TODO add your handling code here:
        int row = tbl_barang.rowAtPoint(evt.getPoint());

//        txt_id.setText(tbl_barang.getValueAt(row, 1).toString());
        txt_nama.setText(tbl_barang.getValueAt(row, 2).toString());
//        txt_merk.setText(tbl_barang.getValueAt(row, 3).toString());
        txt_stok.setText(tbl_barang.getValueAt(row, 4).toString());
        txt_harga.setText(formatRupiah(Integer.valueOf(tbl_barang.getValueAt(row, 5).toString())));
        txt_exp.setText(tbl_barang.getValueAt(row, 7).toString());
    }//GEN-LAST:event_tbl_barangMouseClicked

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        SearchData(txt_search.getText());
    }//GEN-LAST:event_jButton1ActionPerformed

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
            java.util.logging.Logger.getLogger(Inventaris.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Inventaris.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Inventaris.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Inventaris.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Inventaris().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> cb_barang;
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenu jMenu4;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable tbl_barang;
    private javax.swing.JLabel txt_exp;
    private javax.swing.JLabel txt_harga;
    private javax.swing.JLabel txt_nama;
    private javax.swing.JTextField txt_search;
    private javax.swing.JLabel txt_stok;
    // End of variables declaration//GEN-END:variables
}