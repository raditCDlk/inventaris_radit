/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package toko;

import java.sql.*;
import javax.swing.JOptionPane;

/**
 *
 * @author adity
 */
public class Koneksi {

    public static Connection GetConnection() {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            Connection koneksi = DriverManager.getConnection("jdbc:mysql://localhost/db_toko", "root", "");
               return koneksi;
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "error : "+ e);
            return null;
        }

    }

}
