/*
 * Search the database for a client and populate his information into the
 * ticket tool.
 */
package com.microage.tickettool;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import javax.swing.JOptionPane;

/**
 *
 * @author Brian Milne
 */
public class SearchDB {
    public static void SearchDB() {
        // check if name field is empty
        if (Main.main.clientName.getText().trim().length() != 0 ) {
            // search through db for the name
            Connection c = null;
            Statement stmt = null;
            try {
                Class.forName("org.sqlite.JDBC");
                c = DriverManager.getConnection("jdbc:sqlite:lib\\rogerswg.db");
                c.setAutoCommit(false);
                stmt = c.createStatement();
                ResultSet rs = stmt.executeQuery("SELECT * FROM COMPANY WHERE NAME LIKE '%" + Main.main.clientName.getText().toLowerCase() + "%'");
                if (!rs.isBeforeFirst() ) {    
                    System.out.println("No data"); 
                } else {  
                    // Populate client data.
                    Main.main.clientName.setText(rs.getString("name")); 
                    Main.main.title.setText(rs.getString("title"));
                    Main.main.location.setSelectedItem(rs.getString("location"));
                    Main.main.floor.setSelectedItem(rs.getString("floor"));
                    Main.main.phone.setText(rs.getString("phone"));                   
                    Main.main.make.setSelectedItem("HP");
                    Main.main.make3.setSelectedItem("HP");
                    Main.main.model.setSelectedItem(rs.getString("model"));
                    Main.main.serial.setText(rs.getString("serial"));
                    Main.main.asset.setText(rs.getString("asset")); 
                    Main.main.pcName.setText(rs.getString("pcname"));
                    Main.main.model3.setSelectedItem(rs.getString("model"));
                    Main.main.serial3.setText(rs.getString("serial"));
                    Main.main.asset3.setText(rs.getString("asset")); 
                    Main.main.pcName3.setText(rs.getString("pcname"));
                    // Asset collection for database update.
                    AutoUpdateDB.oldTitle = rs.getString("title");
                    AutoUpdateDB.oldLocation = rs.getString("location");
                    AutoUpdateDB.oldFloor = rs.getString("floor");
                    AutoUpdateDB.oldPhone = rs.getString("phone");
                    AutoUpdateDB.oldMake = rs.getString("make");
                    AutoUpdateDB.oldModel = rs.getString("model");
                    AutoUpdateDB.oldSerial = rs.getString("serial");
                    AutoUpdateDB.oldAsset = rs.getString("asset");   
                    AutoUpdateDB.oldPCName = rs.getString("pcname");                                    
                }             
                rs.close();
                stmt.close();
                c.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "Search Error", "Error", JOptionPane.ERROR_MESSAGE);       
            }           
        }
    }
    
    public static void SearchRequesterDB() {
        // check if requester name field is empty
        if (Main.main.requester.getText().trim().length() != 0 ) {
            // search through db for the requester name
            Connection c = null;
            Statement stmt = null;
            try {
                Class.forName("org.sqlite.JDBC");
                c = DriverManager.getConnection("jdbc:sqlite:lib\\rogerswg.db");
                c.setAutoCommit(false);
                stmt = c.createStatement();
                ResultSet rs = stmt.executeQuery("SELECT * FROM COMPANY WHERE NAME LIKE '%" + Main.main.requester.getText().toLowerCase() + "%'");
                if (!rs.isBeforeFirst() ) {    
                    System.out.println("No data"); 
                } else {  
                    // Populate client data.
                    Main.main.requester.setText(rs.getString("name"));                                
                }             
                rs.close();
                stmt.close();
                c.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "Search Error", "Error", JOptionPane.ERROR_MESSAGE);       
            }           
        }
    }
}

