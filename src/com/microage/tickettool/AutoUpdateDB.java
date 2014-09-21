/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
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
public class AutoUpdateDB {
    
    // Global Variables
    static String oldTitle; static String oldLocation; static String oldFloor; 
    static String oldPhone; static String oldMake; static String oldModel; 
    static String oldSerial; static String oldAsset; static String oldPCName;
    
    static String newTitle; static String newLocation; static String newFloor; 
    static String newPhone; static String newMake; static String newModel; 
    static String newSerial; static String newAsset; static String newPCName;
    
    public static void AutoUpdateDB() {
        // ************** Update Database *********************
        boolean update = false;
        // Update Title
        if (!Main.main.title.getText().equals(oldTitle)) {
            newTitle = Main.main.title.getText();
            update = true;
        } else
            newTitle = oldTitle;
        // Update Location
        if (!Main.main.location.getSelectedItem().equals(oldLocation)) {
            newLocation = Main.main.location.getSelectedItem().toString();
            update = true;
        } else
            newLocation = oldLocation;
        // Update Floor
        if (!Main.main.floor.getSelectedItem().equals(oldFloor)) {
            newFloor = Main.main.floor.getSelectedItem().toString();
            update = true;
        } else
            newFloor = oldFloor;
        // Update Phone
        if (!Main.main.phone.getText().equals(oldPhone)) {
            newPhone = Main.main.phone.getText();
            update = true;
        } else
            newPhone = oldPhone;
        
        // Update PC Info if HP
        boolean pcUpdate = false;
        if (Main.main.make.getSelectedItem().equals("HP")) {
            // Update Model
            if (!Main.main.model.getSelectedItem().equals(oldModel)) {
                newModel = Main.main.model.getSelectedItem().toString();
                pcUpdate = true;
            } else
                newModel = oldModel;
            // Update S/N
            if (!Main.main.serial.getText().equals(oldSerial)) {
                newSerial = Main.main.serial.getText().toString();
                pcUpdate = true;
            } else
                newSerial = oldSerial;
            // Update Asset Tag
            if (!Main.main.asset.getText().equals(oldAsset)) {
                newAsset = Main.main.asset.getText().toString();
                pcUpdate = true;
            } else
                newAsset = oldAsset;
            // Update PC Name
            if (!Main.main.pcName.getText().equals(oldPCName)) {
                newPCName = Main.main.pcName.getText();
                pcUpdate = true;
            } else
                newPCName = oldPCName;
        }
  
        boolean newUser = false;
        // Add new user
        if (!checkDatabase()) {
            newUser = true;
        } 
        
        if (update == true || pcUpdate == true || newUser == true) {
            int dialogButton = JOptionPane.YES_NO_OPTION;
            int dialogResult = JOptionPane.showConfirmDialog (null, "Client or computer info has changed. Would you like to update the database?","Warning",dialogButton);
            if(dialogResult == JOptionPane.YES_OPTION){
                if (newUser == true) {
                    addToDatabase();
                } else {
                    // Update database                
                    Connection c = null;
                    Statement stmt = null;
                    try {
                        Class.forName("org.sqlite.JDBC");
                        c = DriverManager.getConnection("jdbc:sqlite:lib\\rogerswg.db");
                        c.setAutoCommit(false);
                        String sql;
                        stmt = c.createStatement();
                        if (update == true) {     
                            // Update existing user
                            sql = "UPDATE company SET title='" +newTitle+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET location='" +newLocation+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET floor='" +newFloor+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET phone='" +newPhone+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);

                        }
                        if (pcUpdate == true) {
                            sql = "UPDATE company SET make='HP' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET model='" +newModel+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET serial='" +newSerial+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET asset='" +newAsset+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                            sql = "UPDATE company SET pcname='" +newPCName+"' WHERE name ='"+Main.main.clientName.getText()+"' ";
                            stmt.executeUpdate(sql);
                        }
                        stmt.close();
                        c.commit();
                        c.close();
                    } catch ( Exception e ) {
                        System.err.println( e.getClass().getName() + ": " + e.getMessage());
                        System.exit(0);
                    }
                }
            }
        }
    }
    
    private static boolean checkDatabase() {
        boolean result = false;
        Connection c = null;
        Statement stmt = null;
        try {
            Class.forName("org.sqlite.JDBC");
            c = DriverManager.getConnection("jdbc:sqlite:lib\\rogerswg.db");
            c.setAutoCommit(false);
            stmt = c.createStatement();
            ResultSet rs = stmt.executeQuery("SELECT * FROM COMPANY WHERE NAME LIKE '%" + Main.main.clientName.getText().toLowerCase() + "%'");
            if (!rs.isBeforeFirst() )   
                result = false;
            else 
                result = true;                    
            rs.close();
            stmt.close();
            c.close();
        } catch ( Exception e ) {
            JOptionPane.showMessageDialog(null, "Auto Update Database Error", "Error", JOptionPane.ERROR_MESSAGE);
        }
        return result;
    }
    
    private static void addToDatabase() {
        Connection c = null;
        Statement stmt = null;
        try {
            Class.forName("org.sqlite.JDBC");
            c = DriverManager.getConnection("jdbc:sqlite:lib\\rogerswg.db");
            c.setAutoCommit(false);          
            stmt = c.createStatement();
            String sql = "INSERT INTO COMPANY (PCNAME,MAKE,MODEL,SERIAL,ASSET,NAME,TITLE,LOCATION,FLOOR,PHONE) " +
                         "VALUES ('"+Main.main.pcName.getText()+"', '"+
                                     Main.main.make.getSelectedItem()+"', '"+
                                     Main.main.model.getSelectedItem()+"', '"+
                                     Main.main.serial.getText()+"', '"+
                                     Main.main.asset.getText()+"', '"+
                                     Main.main.clientName.getText()+"', '"+
                                     Main.main.title.getText()+"', '"+
                                     Main.main.location.getSelectedItem()+"', '"+
                                     Main.main.floor.getSelectedItem()+"', '"+
                                     Main.main.phone.getText()+"' );"; 
            stmt.executeUpdate(sql);     
            stmt.close();
            c.commit();
            c.close();
        } catch ( Exception e ) {
            System.err.println( e.getClass().getName() + ": " + e.getMessage() );
            System.exit(0);
        }
    }
}
