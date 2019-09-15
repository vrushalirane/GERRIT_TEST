package com;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


public class CruciableList {
	
	public static void main(String[] args){
		Connection con=null;
		try{    
			ReportDao dao=new ReportDao();
			Properties props=dao.getPropValues();
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
			Calendar c = Calendar.getInstance();
			long today = new Date().getTime()-2;
			System.out.println("millis:::::"+today);
			Class.forName("com.mysql.jdbc.Driver");
			c.add(Calendar.DAY_OF_MONTH, -30);  
			System.out.println("Date:::::::::::::"+sdf.format(c.getTime()));
			long lastmonth = (c.getTime()).getTime();
			System.out.println("lastmonth::::"+lastmonth);
			//PasswordSecurityService passwordSecurityService=SecurityServiceFactory.getPasswordSecurityService(); 
			ArrayList<String[]> list = new ArrayList<String[]>();
			// sonar for vulnerability
			con = DriverManager.getConnection("jdbc:mysql://itddhsdevalma01:3306/fisheyedb", "fisheyeuser",
					"3lignpd247");
			Statement stmt = con.createStatement();
			String sql ="select * from cru_review where cru_create_date <"+today+" and cru_create_date>"+lastmonth;
			System.out.println("sql::::"+sql);
			ResultSet rs=stmt.executeQuery(sql);
	 		ResultSetMetaData md = rs.getMetaData();
	 		int columns = md.getColumnCount();
	 		md=rs.getMetaData();
	 		
	 		//System.out.println("inside GetExcel data:::::::::::"+columns);
	 		String[] Header = new String[columns];
	 		for (int k=1;k<=columns;k++){
	 			//System.out.println("META DATA:::::::::::::"+md.getColumnName(k));
	 			Header[k-1]=md.getColumnName(k);
	 		}
	 		list.add(Header);
	 		while(rs.next())  {
	 		//	System.out.println("inside GetExcel data:::::::::::");
	 		String[] row = new String[columns];
	 	     for(int i=1; i<=columns; ++i){      
	 	      row[i-1]=rs.getString(i);
	 		}
	 	     list.add(row);
	 	    
	 		}
	 		
	 		createExcel( list, props );
	 		sendEmail( props);
	 		if(con!=null){
				con.close();
			}
			if(rs!=null){
				rs.close();
			}
			if(stmt!=null){
				stmt.close();
			}
			
			
			
		}catch(Exception e){
				e.printStackTrace();
	}  
			  
}  
	
	public  static void sendEmail(Properties config){
	     String to = config.getProperty("email.to1")==null?"":config.getProperty("email.to1");
	     String cc = config.getProperty("email.cc1")==null?"":config.getProperty("email.cc1");
	     String from = config.getProperty("email.from");
	     String host = config.getProperty("email.host");//or IP address  
	 
	    //Get the session object  
	     Properties properties = System.getProperties();  
	     properties.setProperty("mail.smtp.host", host);  
	     Session session = Session.getDefaultInstance(properties);  
	     //compose the message  
	     try{  
	        MimeMessage message = new MimeMessage(session);
	        Multipart multipart=new MimeMultipart();
	        BodyPart messageBodyPart = new MimeBodyPart();
	        message.setFrom(new InternetAddress(from));  
	        //message.addRecipient(Message.RecipientType.TO,new InternetAddress(to));  
	        message.setSubject(config.getProperty("email.subject1"));  
	       
	        String[] email=to.split(";");

	        InternetAddress[] addressTo = new InternetAddress[email.length];
	        //int counter = 0;
	        int counter=0;
	        for ( counter=0;counter<email.length;counter++) {
	        	
	             addressTo[counter] = new InternetAddress(email[counter]);
	             
	        }

	        // CC
	       
	        message.addRecipients(Message.RecipientType.TO,addressTo); 
	        if(cc.length()>0){
	        	email=cc.split(";");
	        InternetAddress[] addressCC = new InternetAddress[email.length];
	        counter=0;
	        System.out.println("email  >>"+email.length);
	        for (counter=0;cc.length()>0 && counter<email.length;counter++) {
	        	
	        	addressCC[counter] = new InternetAddress(email[counter]);
	             System.out.println("mail   "+addressCC[counter]);
	        }
	      
	       
	        if(addressCC.length>0)
	        	message.addRecipients(Message.RecipientType.CC,addressCC);
	        
	        }
	        for (final File fileEntry : (new File(config.getProperty("filepath1")+File.separator)).listFiles()) {
	        	DataSource source = new FileDataSource(fileEntry);
	        	messageBodyPart=new MimeBodyPart();
		        messageBodyPart.setDataHandler(new DataHandler(source));
		        messageBodyPart.setFileName(fileEntry.getName());
		        
		        multipart.addBodyPart(messageBodyPart);
	        }
	        
	        message.setContent(multipart);
	        //session.setDebug(true);
	     // Send message 
	        Transport.send(message);  
	        System.out.println("message sent successfully....");  
	 
	     }catch (Exception ex) {
	    	 ex.printStackTrace();}  
	  }

	 public static void createExcel(ArrayList<String[]> list,Properties props ){
		 try {
	         String filename = new String(props.getProperty("filepath1")+File.separator+"Crucible.xls");

	         // This will output the full path where the file will be written to...
	         File exlFile = new File(filename);

				if (!exlFile.exists()) {
					exlFile.createNewFile();
				}
				
	         WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlFile);
	         WritableSheet writableSheet = writableWorkbook.createSheet("Sheet1", 0);
	         for(int i=0;i<list.size();i++){
	        	 String[] row = list.get(i);
	        	 for (int j=0;j<row.length;j++){
	        		 Label lb1= new Label(j,i,row[j]);
	        		 writableSheet.addCell(lb1);
	        	 }
	        		 
	        	
	         }
	         

	         //Write and close the workbook
	         writableWorkbook.write();
	         writableWorkbook.close();

	     } catch (Exception e) {
	         e.printStackTrace();
	     } 
	 }
	
}
