/*
 * Generates an email that is sent to the coordinator and technician.
 */
package com.microage.tickettool;

import java.awt.Image;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.swing.ImageIcon;
import javax.swing.JOptionPane;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

/**
 *
 * @author Brian Milne
 */
public class SendEmail {
    
    public static void SendEmail() {
        final String username = "microage.wgs@gmail.com";
        final String password = "M1(r0w@g3";
        String body = "";

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587"); //587
           
        Session session = Session.getInstance(props,
          new javax.mail.Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                }
          });
        
        try {
            //String toName = "rsomerville@tor.microage.ca";
            //String toName2 = "ritani@tor.microage.ca";
            String toName = "cro_mos@hotmail.com";
            String toName2 = "cro_mos@hotmail.com";
            // Find CC address to email technician.
            String ccName = System.getProperty("user.name") + "@tor.microage.ca";
            // Date for Subject Line
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
            Date date = new Date();
            dateFormat.format(date);
            Date stime = (Date)Main.main.startDate.getValue();
            DateFormat excelFormat = new SimpleDateFormat("M/d/yyyy h:mm aa");  
            Date etime = (Date)Main.main.endDate.getValue();
       
            // Set Recipients
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress("microage.wgs@gmail.com"));
            message.addRecipients(Message.RecipientType.TO, InternetAddress.parse(toName));
            message.addRecipients(Message.RecipientType.TO, InternetAddress.parse(toName2));
            try {
                message.addRecipients(Message.RecipientType.CC, InternetAddress.parse(ccName));
            } catch (Exception e) {}
            
            // ******************** DSS or IMAC email message *****************
            if (Main.main.tab.getSelectedIndex() == 0) {
                if (Main.main.status.getSelectedItem().equals("Closed"))
                    message.setSubject("New Ticket for " + Main.main.clientName.getText() + " " + dateFormat.format(date));
                else {
                    message.setSubject("***** Rush ****** New Ticket for " + Main.main.clientName.getText() + " " + dateFormat.format(date));
                    message.addHeader("X-Priority", "2");
                }
                body = "This email was automatically generated by the Rogers WGS Ticket Tool." +
                        "<br>" + "<B>Requester's Name: </B>" + Main.main.requester.getText() +
                        "<br>" + "<B>Affected Client's Name: </B>" + Main.main.clientName.getText() +
                        "<br>" + "<B>Title: </B>" + Main.main.title.getText() +
                        "<br>" + "<B>Location: </B>" + Main.main.location.getSelectedItem() + " <B>Floor</B> " + Main.main.floor.getSelectedItem() +
                        "<br>" + "<B>Phone Number: </B>" + Main.main.phone.getText() +
                        "<br>" + "<B>Notification Type: </B>" + Main.main.notification.getSelectedItem() +
                        "<br>" + "<B>Reoccurence: </B>" + Main.main.reoccurence.getSelectedItem() +
                        "<br>" + "<B>Action Taken: </B>" + Main.main.actionTaken.getText() +
                        "<br><br>" + "<B>Device Make: </B>" + Main.main.make.getSelectedItem() +
                        "<br>" + "<B>Model: </B>" + Main.main.model.getSelectedItem() +
                        "<br>" + "<B>Serial Number: </B>" + Main.main.serial.getText() +
                        "<br>" + "<B>Asset: </B>" + Main.main.asset.getText() +
                        "<br>" + "<B>Start Time: </B>" + excelFormat.format(stime) + 
                        "<br>" + "<B>Completion Time: </B>" + excelFormat.format(etime) +
                        "<br>" + "<B>Status: </B>" + Main.main.status.getSelectedItem() +
                        "<br><br>" + "<B>Issue: </B>" + Main.main.issue.getText() +
                        "<br><br>" + "<B>Resolution: </B>" + Main.main.resolution.getText() +
                        "<br><br>" + "<B>Technician: </B>" + Main.main.technician.getSelectedItem();
            }
            // ****************** PC Refresh Email ********************** 
            else if (Main.main.tab.getSelectedIndex() == 1) {     
                message.setSubject("New Install for " + Main.main.clientName.getText() + " " + dateFormat.format(date));
                body = "This email was automatically generated by the Rogers WGS Ticket Tool." +
                        "<br>" + "<B>Deskside Tech Name: </B>" + Main.main.technician.getSelectedItem() +
                        "<br>" + "<B>Client's Name: </B>" + Main.main.clientName.getText() +
                        "<br>" + "<B>Date: </B>" + excelFormat.format(stime) + " - " + excelFormat.format(etime) +
                        "<br>" + "<B>ISM Ticket Number: </B>" + Main.main.chTicket.getText() +
                        "<br>" + "<B>Location: </B>" + Main.main.location.getSelectedItem() + 
                        "<br>" + "<B>Floor: </B>" + Main.main.floor.getSelectedItem() + 
                        "<br>" + "<B>Desk: </B>" + Main.main.desk.getText() +
                        "<br>" + "<B>Admin Access: </B>" + Main.main.admin3.getSelectedItem() +
                        "<br>" + "<B>Application List: </B>" + Main.main.appList3.getText() +
                        "<br>" + "<B>Dock Provided: </B>" + Main.main.dock.isSelected() +
                        "<br>" +
                        "<br>" + "<B>  Old Computer </B>" +
                        "<br>" + "<B>Make: </B>" + Main.main.make3.getSelectedItem() +
                        "<br>" + "<B>Model: </B>" + Main.main.model3.getSelectedItem() +
                        "<br>" + "<B>Serial: </B>" + Main.main.serial3.getText() +
                        "<br>" + "<B>Asset Tag: </B>" + Main.main.asset3.getText() +
                        "<br>" + "<B>Product Num: </B>" + Main.main.partNum3.getText() +
                        "<br>" + "<B>PC Name: </B>" + Main.main.pcName3.getText() +
                        "<br>" +
                        "<br>" + "<B>  New Computer </B>" +
                        "<br>" + "<B>Make: </B>" + Main.main.make2.getSelectedItem() +
                        "<br>" + "<B>Model: </B>" + Main.main.model2.getSelectedItem() +
                        "<br>" + "<B>Serial: </B>" + Main.main.serial2.getText() +
                        "<br>" + "<B>Asset Tag: </B>" + Main.main.asset2.getText() +
                        "<br>" + "<B>Product Num: </B>" + Main.main.partNum2.getText() +
                        "<br>" + "<B>PC Name: </B>" + Main.main.pcName2.getText();
            }
            // ****************** PC Refresh Email **********************   
            else if (Main.main.tab.getSelectedIndex() == 2) {   
                message.setSubject("PC Refresh for " + Main.main.clientName.getText() + " " + dateFormat.format(date));
                body = "This email was automatically generated by the Rogers WGS Ticket Tool." +
                        "<br>" + "<B>Deskside Tech Name: </B>" + Main.main.technician.getSelectedItem() +
                        "<br>" + "<B>Client's Name: </B>" + Main.main.clientName.getText() +
                        "<br>" + "<B>Date: </B>" + excelFormat.format(stime) + " - " + excelFormat.format(etime) +
                        "<br>" + "<B>ISM Ticket Number: </B>" + Main.main.chTicket4.getText() +
                        "<br>" + "<B>Location: </B>" + Main.main.location.getSelectedItem() + 
                        "<br>" + "<B>Floor: </B>" + Main.main.floor.getSelectedItem() + 
                        "<br>" + "<B>Desk: </B>" + Main.main.desk4.getText() +
                        "<br>" + "<B>Dock Provided: </B>" + Main.main.dock4.isSelected() +
                        "<br>" +
                        "<br>" + "<B>  New Install </B>" +
                        "<br>" + "<B>Type: </B>" + Main.main.type4.getSelectedItem() +
                        "<br>" + "<B>Make: </B>" + Main.main.make4.getSelectedItem()+
                        "<br>" + "<B>Model: </B>" + Main.main.model4.getText() +
                        "<br>" + "<B>Serial: </B>" + Main.main.serial4.getText() +
                        "<br>" + "<B>Asset Tag: </B>" + Main.main.asset4.getText() +
                        "<br>" + "<B>Product Num: </B>" + Main.main.partNum4.getText() +
                        "<br>" + "<B>PC Name: </B>" + Main.main.pcName4.getText() +
                        "<br>" + "<B>Description: </B>" + Main.main.description.getText() ;
            }
            BodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText(body);
            messageBodyPart.setContent(body, "text/html");               

            // Add Attachment
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);
            MimeBodyPart attachment = new MimeBodyPart();
            String fileName = "i:\\Rogers\\Tickets\\" + CreateXLS.attachmentName;
            DataSource source = new FileDataSource(fileName);
            attachment.setDataHandler(new DataHandler(source));
            attachment.setFileName(CreateXLS.attachmentName);
            multipart.addBodyPart(attachment);
            message.setContent(multipart);

            // Send Email
            Transport.send(message);
            
            // Email confirmation window.
            final ImageIcon icon = new ImageIcon("lib\\mail.png");
            Image img = icon.getImage();  
            Image newimg = img.getScaledInstance(100, 100,  java.awt.Image.SCALE_SMOOTH);  
            ImageIcon newIcon = new ImageIcon(newimg); 
            JOptionPane.showMessageDialog(null, "Email sent!", "Email", JOptionPane.INFORMATION_MESSAGE, newIcon);
        } catch (MessagingException e) {
            JOptionPane.showMessageDialog(null, "Unable to email new ticket request. \nCheck your internet connection.\nAlternatively email your log file to RWGSUpdate@tor.microage.ca", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}
