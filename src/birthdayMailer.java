import java.io.File;
import jxl.Sheet;
import jxl.Cell;

import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import javax.mail.*;  
import javax.mail.internet.*;  
import javax.activation.*;  

import jxl.Workbook;
import jxl.read.biff.BiffException;

public class birthdayMailer {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		File list = new File("C:/birthdaMailerList/birthdayList.xls");
		
		Workbook w1;
        try {
			w1 = Workbook.getWorkbook(list);
		
			Sheet sheet1 = w1.getSheet(0);
        
			//System.out.println("sheet1.getColumns() "+sheet1.getColumns());
			//System.out.println("sheet1.getRows() "+sheet1.getRows());
        
			for (int x = 1; x < sheet1.getRows(); x++) {
				Cell cell1 = sheet1.getCell(1, x);
				//System.out.println("cell1 contents: "+cell1.getContents().toString());
				
				Date date = new Date();
				Calendar currentDate = Calendar.getInstance();
				currentDate.setTime(date);
			    //int year = currentDate.get(Calendar.YEAR);
			    int todayMonth = currentDate.get(Calendar.MONTH);
			    int todayDay = currentDate.get(Calendar.DAY_OF_MONTH);
			    
			    DateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
			    Date bday = (Date)formatter.parse(cell1.getContents().toString()); 
			    
			    
			    Calendar listDate = Calendar.getInstance();
			    listDate.setTime(bday);
			    //int year = listDate.get(Calendar.YEAR);
			    int listMonth = listDate.get(Calendar.MONTH);
			    int listDay = listDate.get(Calendar.DAY_OF_MONTH);
			    
			    String birthdayName = sheet1.getCell(0, x).getContents().toString();
			    
			    if(listMonth == todayMonth && listDay == todayDay){
			    	System.out.println("birthday today: "+ birthdayName);
			    
				    String to = "email_id@gmail.com";//change accordingly  
				    String from = "email_id@gmail.com";//change accordingly  
			      
				    //Get the session object  
				      Properties properties = System.getProperties();
				      
				      properties.put("mail.smtp.starttls.enable", "true"); 
				      properties.put("mail.smtp.host", "smtp.gmail.com");
				      properties.put("mail.smtp.auth", "true");
				     
				      Session session = Session.getDefaultInstance(properties,new Authenticator() {
	
		                    protected PasswordAuthentication getPasswordAuthentication() {
		                        return new PasswordAuthentication(
		                                "email_id@gmail.com", "password");// Specify the Username and the PassWord
		                    }
				      });  
				  
				     //compose the message  
				      
			         MimeMessage message = new MimeMessage(session);  
			         message.setFrom(new InternetAddress(from));  
			         message.addRecipient(Message.RecipientType.TO,new InternetAddress(to));  
			         message.setSubject("Birthday Reminder");  
			         message.setText("Today is "+ birthdayName + "'s birthday.");  
			  
			         // Send message  
			         Transport.send(message);  
			         System.out.println("message sent successfully....");  
			    }
			}
			
        } catch (BiffException | IOException | ParseException | MessagingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
