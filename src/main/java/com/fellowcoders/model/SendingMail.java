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

    public static void main(String[] args) {
    	
    	
   
    	
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
    		String name = fetchDataObj.getCellData("STUDENT_DETAILS","NAME" , i);

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
		
			if (i != 1) {
				  boolean mailSendSuccessfully = sendMail(toMailId,name, sb.toString());
				    if (mailSendSuccessfully) {
			    		boolean flagUpdated = fetchDataObj.setCellData("STUDENT_DETAILS", "MAIL_SENT_FLAG", i, "Yes");
				    }
				    sb = new StringBuffer();
			}
		  


    	}
	

    	
    }
    
    public static boolean sendMail(String toMailId,String name, String messageTable ){
        // Recipient's email ID needs to be mentioned.
       
        boolean mailSentSuccessfully = false;
            String to = toMailId;

            // Sender's email ID needs to be mentioned
            final String from = "debaryapal11@gmail.com";

            // Assuming you are sending email from through gmails smtp
            String host = "smtp.gmail.com";

            // Get system properties
            Properties properties = new Properties();    

            // Setup mail server
            properties.put("mail.smtp.host", host);
            properties.put("mail.smtp.port", "587");
           
            properties.put("mail.smtp.socketFactory.port", "465");    
            properties.put("mail.smtp.socketFactory.class",    
                      "javax.net.ssl.SSLSocketFactory");    
            properties.put("mail.smtp.auth", "true");    
          //  properties.put("mail.smtp.port", "465");
          //  properties.put("mail.smtp.ssl.enable", "true");
           // properties.put("mail.smtp.auth", "true");

           
           
           
           
           
           
           
            // Get the Session object.// and pass username and password
           // Session session = Session.getDefaultInstance(properties);

            Session session = Session.getDefaultInstance(properties,    
                    new javax.mail.Authenticator() {    
                    protected javax.mail.PasswordAuthentication getPasswordAuthentication() {    
                    return new javax.mail.PasswordAuthentication(from,"deb1arya@76");  
                    }    
                   });    
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
            	messagebody.append("Hi "+ name + ",");
            	messagebody.append("\n");
            	messagebody.append("We are delighted to inform you that you are successfully enrolled with us for the Crack Job Batch 1 (CJB1). Please find your details below:\r\n");
            	messagebody.append("\n");
            	messagebody.append("<html>");
            	messagebody.append("<table>");

            	messagebody.append(messageTable);
            	messagebody.append("\n");

            	messagebody.append("</table>");

            	messagebody.append("</html>");
            	messagebody.append("Thank you for joining us and trusting our services. Please reply back for any queries or changes in your details.\r\n");
            	messagebody.append("Regards\r\n CrackjobTeam\r\nDigital Education Foundation\r\ncontact:8910274229\n");
            	



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




