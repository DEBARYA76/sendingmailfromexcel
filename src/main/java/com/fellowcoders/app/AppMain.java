package com.fellowcoders.app;


import com.fellowcoders.model.*;
import java.util.Properties;
import javax.mail.*;

public class AppMain {

	public static void main(String[] args)
	{
		

    	StringBuffer  sb = new StringBuffer();
    	SendingMail send=new SendingMail();
    	
    	FetchData fetchDataObj = new FetchData("E:\\ECLIPSE IDE\\sendingmailforxcell\\EXCELLFILE\\STUDENTS&COURSE DETAILS FOR CRACKJOB.xlsx");
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
    	
    		/*sb.append(header.toString());
			sb.append("<tr>");

    		for (int j = 0 ; j < studentColumnCount ; j++) {
        		
    			sb.append("<td>");
    			sb.append(fetchDataObj.getCellData("STUDENT_DETAILS",j,i));
    			if (studentColumnCount != j) {
        			sb.append("</td>");

    			}
    		
    		}*/
    		String toMailId = fetchDataObj.getCellData("STUDENT_DETAILS","EMAIL" , i);
    		String name = fetchDataObj.getCellData("STUDENT_DETAILS","NAME" , i);

    		int rowNum = fetchDataObj.getCellRowNum("COURSE_DETAILS","COURSE_CODE" , fetchDataObj.getCellData("STUDENT_DETAILS","COURSE_CODE" , i));
    		if (i == 1){
        		rowNum ++;
        		rowNum ++;
    		}
			/*for (int m =0 ; m <courseColumnCount; m++) {
    			sb.append("<td>");
    			sb.append(fetchDataObj.getCellData("COURSE_DETAILS",m , rowNum));
    			sb.append("</td>");

			}
			sb.append("</tr>");*/
		
			if (i != 1) {
				 // boolean mailSendSuccessfully = send.sendMail(toMailId,name, sb.toString());
				boolean mailSendSuccessfully = send.sendMail(toMailId,name, i, courseRowCount ,studentRowCount,courseColumnCount ,studentColumnCount );
				    if (mailSendSuccessfully) {
			    		boolean flagUpdated = fetchDataObj.setCellData("STUDENT_DETAILS", "MAIL_SENT_FLAG", i, "Yes");
				    }
				    sb = new StringBuffer();
			}
		  


    	}
	

	}

}
