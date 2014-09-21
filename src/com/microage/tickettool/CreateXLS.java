/*
 * Generate the excel file that will be imported into Clearview.
 */
package com.microage.tickettool;

import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author Brian Milne
 */
public class CreateXLS {
    
    protected static String attachmentName;
            
    protected static void CreateXLS(String location) {
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
        Date date = new Date();
        new File(location).mkdirs();
        try{
            String filename = location + Main.main.clientName.getText() + "_" + dateFormat.format(date) + ".xls";
            attachmentName = Main.main.clientName.getText() + "_" + dateFormat.format(date) + ".xls";
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("SAMPLE");  
            HSSFRow rowhead = sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Record");
            rowhead.createCell(1).setCellValue("Location");
            rowhead.createCell(2).setCellValue("ContactFirst");
            rowhead.createCell(3).setCellValue("ContactLast");
            rowhead.createCell(4).setCellValue("Phone");
            rowhead.createCell(5).setCellValue("Ext");
            rowhead.createCell(6).setCellValue("Email");
            rowhead.createCell(7).setCellValue("Department");       
            rowhead.createCell(8).setCellValue("TicketOrigin");
            rowhead.createCell(9).setCellValue("Reference");
            rowhead.createCell(10).setCellValue("CSR_Code");
            rowhead.createCell(11).setCellValue("Primary_Tech_Code");
            rowhead.createCell(12).setCellValue("Item_no");
            rowhead.createCell(13).setCellValue("Serial");
            rowhead.createCell(14).setCellValue("Asset");
            rowhead.createCell(15).setCellValue("VIP");
            rowhead.createCell(16).setCellValue("Category");
            rowhead.createCell(17).setCellValue("Problem");
            rowhead.createCell(18).setCellValue("Subproblem");
            rowhead.createCell(19).setCellValue("Cause_Code");
            rowhead.createCell(20).setCellValue("Repair_Code");
            rowhead.createCell(21).setCellValue("TopCallDrivers");
            rowhead.createCell(22).setCellValue("Software");
            rowhead.createCell(23).setCellValue("Reported_Date");
            rowhead.createCell(24).setCellValue("Entry_Date");
            rowhead.createCell(25).setCellValue("Dis_Time");
            rowhead.createCell(26).setCellValue("Arv_Time");
            rowhead.createCell(27).setCellValue("Cmp_Time");
            rowhead.createCell(28).setCellValue("BCH_CODE");
            rowhead.createCell(29).setCellValue("Survey_Date");
            rowhead.createCell(30).setCellValue("Call_Back_Confirm_Closure");
            rowhead.createCell(31).setCellValue("Reoccurence");
            rowhead.createCell(32).setCellValue("Taken_Actions");
            rowhead.createCell(33).setCellValue("MemoS");
            rowhead.createCell(34).setCellValue("MemoR");
            rowhead.createCell(35).setCellValue("MemoC");
            rowhead.createCell(36).setCellValue("Requestor");

            HSSFRow row =   sheet.createRow((short)1);
            row.createCell(0).setCellValue("1"); // Record ID count
            row.createCell(1).setCellValue(LocationCode.LocationCode()); // Location Code
            String fName = "";
            String lName = "";
            if (Main.main.clientName.getText().contains(" ")) {          
                int space = Main.main.clientName.getText().indexOf(" ");
                fName = Main.main.clientName.getText().substring(0, space);
                lName = Main.main.clientName.getText().substring(space+1, Main.main.clientName.getText().length());
            } else fName = Main.main.clientName.getText();
            row.createCell(2).setCellValue(fName); // First Name
            row.createCell(3).setCellValue(lName); // Last Name
            row.createCell(4).setCellValue(Main.main.phone.getText()); // Phone
            row.createCell(5).setCellValue(""); // Ext
            row.createCell(6).setCellValue(fName + "." + lName + "@rci.rogers.com"); // Email
            row.createCell(7).setCellValue("Rogers"); // Department
            row.createCell(8).setCellValue(Main.main.notification.getSelectedItem().toString()); // Notification Type
            row.createCell(9).setCellValue(Main.main.chTicket.getText()); // Incident Reference Number
            row.createCell(10).setCellValue(""); // CSR_Code
            row.createCell(11).setCellValue(System.getProperty("user.name")); //Primary Tech
            if (Main.main.tab.getSelectedIndex() == 0) {
                row.createCell(12).setCellValue(Main.main.model.getSelectedItem().toString()); // Item Number Model code
                row.createCell(13).setCellValue(Main.main.serial.getText()); // Serial
                row.createCell(14).setCellValue(Main.main.asset.getText()); // Asset
            } else if (Main.main.tab.getSelectedIndex() == 1) {
                row.createCell(12).setCellValue(Main.main.model3.getSelectedItem().toString()); // Item Number Model code
                row.createCell(13).setCellValue(Main.main.serial3.getText()); // Serial
                row.createCell(14).setCellValue(Main.main.asset3.getText()); // Asset
            } else if (Main.main.tab.getSelectedIndex() == 2) {
                row.createCell(12).setCellValue(""); // Item Number Model code
                row.createCell(13).setCellValue(Main.main.serial4.getText()); // Serial
                row.createCell(14).setCellValue(Main.main.asset4.getText()); // Asset
            }
            String sType = "";
            String cType = "";
            if (Main.main.tab.getSelectedIndex() == 1){
                sType = "IMAC-VIP";
                cType = "REFRESH";
            } else if (Main.main.tab.getSelectedIndex() == 2){
                sType = "IMAC-VIP";
                cType = "IMAC";
            } else if (Main.main.imacRadio.isSelected()) {
                sType = "IMAC-VIP";
                cType = "IMAC";
            } else if (Main.main.dssRadio.isSelected()){
                sType = "SEV4-VIP";
                cType = "DESKSIDE";
            } 
            int vip = 0;
            if (Main.main.title.getText().contains("EA") || Main.main.title.getText().contains("VP") || 
                Main.main.title.getText().contains("CEO") || Main.main.title.getText().contains("Chairman") || 
                Main.main.title.getText().contains("President") || Main.main.title.getText().contains("AA"))
                vip = 1;
            row.createCell(15).setCellValue(vip); // Service Type
            row.createCell(16).setCellValue(cType); // Category
            row.createCell(17).setCellValue(""); // Problem
            row.createCell(18).setCellValue(""); // Subproblem
            row.createCell(19).setCellValue(""); // Cause
            row.createCell(20).setCellValue(""); // Repair
            row.createCell(21).setCellValue(""); // Top Call Driver
            row.createCell(22).setCellValue(""); // Software
            Date stime = (Date)Main.main.startDate.getValue();
            Date etime = (Date)Main.main.endDate.getValue();
            DateFormat excelFormat = new SimpleDateFormat("M/d/yyyy h:mm aa");  
            row.createCell(23).setCellValue(excelFormat.format(stime)); // Reported Date
            row.createCell(24).setCellValue(excelFormat.format(stime)); // Entry Date
            row.createCell(25).setCellValue(""); // Dispatch Time
            row.createCell(26).setCellValue(excelFormat.format(stime)); // Arrival Time
            row.createCell(27).setCellValue(excelFormat.format(etime)); // Completion Time
            row.createCell(28).setCellValue(""); // BCH_CODE Close Code
            row.createCell(29).setCellValue(""); // Survey Date
            int myInt = (Main.main.confirmClose.isSelected()) ? 1 : 0;
            row.createCell(30).setCellValue(myInt); // Call back confirm close 1 or 0
            int rCheck;
            if (Main.main.reoccurence.getSelectedItem() == "Yes")
                rCheck = 1;
            else
                rCheck = 0;
            row.createCell(31).setCellValue(rCheck); // Reoccurence?
            row.createCell(32).setCellValue(Main.main.actionTaken.getText().toString()); // Taken Action
            if (Main.main.tab.getSelectedIndex() == 0) {
                row.createCell(33).setCellValue(Main.main.issue.getText()); // MemoS
                if (Main.main.status.getSelectedItem() == "Closed") {
                    row.createCell(34).setCellValue(Main.main.resolution.getText()); // MemoR
                } else
                    row.createCell(35).setCellValue(Main.main.resolution.getText()); // MemoC
            } else if (Main.main.tab.getSelectedIndex() == 2) {
                row.createCell(33).setCellValue("New Install Request # " + Main.main.chTicket4.getText() + " " + Main.main.description.getText()); // MemoS
                row.createCell(34).setCellValue("Completed the new install \r" +
                        "Deskside Tech Name: " + Main.main.technician.getSelectedItem() + "\r" +
                        "Client's Name: " + Main.main.clientName.getText() + "\r" +
                        "Date: " + excelFormat.format(stime) + " - " + excelFormat.format(etime) + "\r" +
                        "ISM Ticket Number: " + Main.main.chTicket4.getText() + "\r" +
                        "Location: " + Main.main.location.getSelectedItem() +  "\r" +
                        "Floor: " + Main.main.floor.getSelectedItem() +  "\r" +
                        "Desk: " + Main.main.desk4.getText() + "\r" +
                        "Dock Provided: " + Main.main.dock4.isSelected() + "\r" +
                        "  New Install " + "\r" +
                        "Type: " + Main.main.type4.getSelectedItem() + "\r" +
                        "Make: " + Main.main.make4.getSelectedItem()+ "\r" +
                        "Model: " + Main.main.model4.getText() + "\r" +
                        "Serial: " + Main.main.serial4.getText() + "\r" +
                        "Asset Tag: " + Main.main.asset4.getText() + "\r" +
                        "Product Num: " + Main.main.partNum4.getText() + "\r" +
                        "PC Name: " + Main.main.pcName4.getText() + "\r" +
                        "Description: " + Main.main.description.getText()); // MemoR
            } else {
                row.createCell(33).setCellValue("PC Hardware Refresh request # " + Main.main.chTicket.getText()); // MemoS
                row.createCell(34).setCellValue("Refreshed the client's computer. \r" + 
                        "Date: " + excelFormat.format(stime) + " - " + excelFormat.format(etime) + " \r" +
                        "ISM Ticket Number: " + Main.main.chTicket.getText() + "\r" +
                        "Location: " + Main.main.location.getSelectedItem() +  "\r" +
                        "Floor: " + Main.main.floor.getSelectedItem() +  "\r" +
                        "Desk: " + Main.main.desk.getText() + "\r" +
                        "Admin Access: " + Main.main.admin3.getSelectedItem() + "\r" +
                        "Application List: " + Main.main.appList3.getText() + "\r" +
                        "Dock Provided: " + Main.main.dock.isSelected() + "\r" +
                        "  Old Computer " + "\r" +
                        "Make: " + Main.main.make3.getSelectedItem() + "\r" +
                        "Model: " + Main.main.model3.getSelectedItem() + "\r" +
                        "Serial: " + Main.main.serial3.getText() + "\r" +
                        "Asset Tag: " + Main.main.asset3.getText() + "\r" +
                        "Product Num: " + Main.main.partNum3.getText() + "\r" +
                        "PC Name: " + Main.main.pcName3.getText() + "\r" +
                        "  New Computer " + "\r" +
                        "Make: " + Main.main.make2.getSelectedItem() + "\r" +
                        "Model: " + Main.main.model2.getSelectedItem() + "\r" +
                        "Serial: " + Main.main.serial2.getText() + "\r" +
                        "Asset Tag: " + Main.main.asset2.getText() + "\r" +
                        "Product Num: " + Main.main.partNum2.getText() + "\r" +
                        "PC Name: " + Main.main.pcName2.getText()); // MemoR
            }
            if (Main.main.requester.getText().equals("")) 
                row.createCell(36).setCellValue(Main.main.clientName.getText());
            else
                row.createCell(36).setCellValue(Main.main.requester.getText());
            FileOutputStream fileOut =  new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
        } catch ( Exception ex ) {
            JOptionPane.showMessageDialog(null, "Failed to generate Excel file.", "Error", JOptionPane.ERROR_MESSAGE);
       }          
    }
}
