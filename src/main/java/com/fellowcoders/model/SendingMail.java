package com.fellowcoders.model;

import java.util.Properties;

import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class SendingMail
{

    public void sendmail() {
    	
    	
   
    	
    	StringBuffer  sb = new StringBuffer();
    	
    	FetchData fetchDataObj = new FetchData("E:\\ECLIPSE IDE\\sendingmailforxcell\\EXCELLFILE\\STUDENTS &COURSE DETAILS.xlsx");
    	int courseRowCount = 0;
    	int studentRowCount = 0 ;
     	int courseColumnCount = 0;
    	int studentColumnCount = 0 ;
    	if(fetchDataObj.isSheetExist("STUDENT_DETAILS")) {
    		studentRowCount = fetchDataObj.getRowCount("STUDENT_DETAILS");
    		studentColumnCount = fetchDataObj.getColumnCount("STUDENT_DETAILS");
    		

    	}

    	if(fetchDataObj.isSheetExist("COURSE_DETAILS")) {
    		courseRowCount = fetchDataObj.getRowCount("COURSE_DETAILS");
    		courseColumnCount = fetchDataObj.getColumnCount("COURSE_DETAILS");

     	}
    	
    	StringBuffer header = new StringBuffer();
    	header.append("<tr>");

         for (int j = 0 ; j < studentColumnCount ; j++) {
        		
        	 header.append("<td>");
        	 header.append(fetchDataObj.getCellData("STUDENT_DETAILS",j,1));
    			if (studentColumnCount != j) {
        			sb.append("</td>");

    			}
    
    		
    		}
         
         for (int m =0 ; m <courseColumnCount; m++) {
        	 header.append("<td>");
        	 header.append(fetchDataObj.getCellData("COURSE_DETAILS",m , 1));
        	 header.append("</td>");

			}
     	header.append("</tr>");

		

    	for (int i = 2 ; i <= studentRowCount ; i++) {
    		if ("Yes".equalsIgnoreCase(fetchDataObj.getCellData("STUDENT_DETAILS","MAIL_SENT_FLAG" , i))) {
    			continue;
    		}
    		sb.append("<html>");

    		sb.append("<table>");
    		sb.append(header.toString());
			sb.append("<tr>");

    		for (int j = 0 ; j < studentColumnCount ; j++) {
        		
    			sb.append("<td>");
    			sb.append(fetchDataObj.getCellData("STUDENT_DETAILS",j,i));
    			if (studentColumnCount != j) {
        			sb.append("</td>");

    			}
    		
    		}
    		String toMailId = fetchDataObj.getCellData("STUDENT_DETAILS","EMAIL" , i);
    		int rowNum = fetchDataObj.getCellRowNum("COURSE_DETAILS","COURSE_CODE" , fetchDataObj.getCellData("STUDENT_DETAILS","COURSE_CODE" , i));
    		if (i == 1){
        		rowNum ++;
        		rowNum ++;
    		}
			for (int m =0 ; m <courseColumnCount; m++) {
    			sb.append("<td>");
    			sb.append(fetchDataObj.getCellData("COURSE_DETAILS",m , rowNum));
    			sb.append("</td>");

			}
			sb.append("</tr>");
			sb.append("</table>");
			sb.append("</html>");
			if (i != 1) {
				  boolean mailSendSuccessfully = sendMail(toMailId, sb.toString());
				    if (mailSendSuccessfully) {
			    		boolean flagUpdated = fetchDataObj.setCellData("STUDENT_DETAILS", "MAIL_SENT_FLAG", i, "Yes");
				    }
				    sb = new StringBuffer();
			}
		  


    	}
	

    	
    }
    
    
    public static boolean sendMail(String toMailId, String messageTable){
    	 // Recipient's email ID needs to be mentioned.
    	
    	boolean mailSentSuccessfully = false;
        String to = toMailId;

        // Sender's email ID needs to be mentioned
        String from = "debaryapal11@gmail.com";

        // Assuming you are sending email from through gmails smtp
        String host = "172.16.113.70";

        // Get system properties
        Properties properties = System.getProperties();

        // Setup mail server
        properties.put("mail.smtp.host", host);
        properties.put("mail.smtp.port", "25");
      //  properties.put("mail.smtp.port", "465");
      //  properties.put("mail.smtp.ssl.enable", "true");
       // properties.put("mail.smtp.auth", "true");

        
        
        
        
       
        
        
        // Get the Session object.// and pass username and password
        Session session = Session.getDefaultInstance(properties);

        // Used to debug SMTP issues
        session.setDebug(true);

        try {
            // Create a default MimeMessage object.
            MimeMessage message = new MimeMessage(session);

            // Set From: header field of the header.
            message.setFrom(new InternetAddress(from));

            // Set To: header field of the header.
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

            // Set Subject: header field
            message.setSubject("Student Course Enrollment");

            MimeBodyPart messageBodyPart = new MimeBodyPart();
            MimeMultipart multipart = new MimeMultipart();
			messageBodyPart = new MimeBodyPart();
			StringBuffer messagebody = new StringBuffer();
			messagebody.append("Please find enrollment details");
			messagebody.append("\n");
			messagebody.append(messageTable);
			messageBodyPart.setContent(messagebody.toString(), "text/html");
			multipart.addBodyPart(messageBodyPart);
            // Now set the actual message
            message.setContent(multipart);

            System.out.println("sending...");
            // Send message
            Transport.send(message);
            System.out.println("Sent message successfully....");
            mailSentSuccessfully = true;
        } catch (MessagingException mex) {
            mex.printStackTrace();
            mailSentSuccessfully = false;

        }
		return mailSentSuccessfully;

    }

}
