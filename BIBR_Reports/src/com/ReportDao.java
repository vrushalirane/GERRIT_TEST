package com;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
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

import com.deloitte.passwordSecurity.service.PasswordSecurityService;
import com.deloitte.passwordSecurity.service.SecurityServiceFactory;
import com.deloitte.passwordSecurity.service.PasswordSecurityService;
import com.deloitte.passwordSecurity.service.SecurityServiceFactory;
public class ReportDao {
	Properties prop = null;
	
	 public Properties  getPropValues() throws IOException {
		 
			InputStream inputStream=null;
			prop = new Properties();
			try {
				
				String propFileName = "configuration.properties";

				inputStream = getClass().getClassLoader().getResourceAsStream(propFileName);

				if (inputStream != null) {
					prop.load(inputStream);
				} else {
					throw new FileNotFoundException("property file '" + propFileName + "' not found in the classpath");
				}

			} catch (Exception e) {
				System.out.println("Exception: " + e);
			} finally {
				inputStream.close();
			}
			return prop;
		}

	 public  void sendEmail(Properties config){
	     String to = config.getProperty("email.to")==null?"":config.getProperty("email.to");
	     String cc = config.getProperty("email.cc")==null?"":config.getProperty("email.cc");
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
	        message.setSubject(config.getProperty("email.subject"));  
	       
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
	        for (final File fileEntry : (new File(config.getProperty("R2.report.path")+"/")).listFiles()) {
	        	System.out.println("fileEntry::::::"+fileEntry.getAbsolutePath());
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
	 

		
	 public ArrayList<String[]> getExcelData(String query,int m) throws SQLException{

	 	 ArrayList<String[]> list = new ArrayList<String[]>();
	 	Connection con=null;
	 	ResultSet rs=null;
	 	Statement stmt=null;
	 	try{    
	 		PasswordSecurityService passwordSecurityService=SecurityServiceFactory.getPasswordSecurityService(); 
	 		 con=getConnection(prop.getProperty("connectionString"+m),prop.getProperty("username"+m),passwordSecurityService.decrypt(prop.getProperty("password"+m)));
	 		 stmt=con.createStatement(); 
	 		//System.out.println("prop.getProperty(query)::::::"+query);
	 		 rs=stmt.executeQuery(query);
	 		ResultSetMetaData md = rs.getMetaData();
	 		int columns = md.getColumnCount();
	 		md=rs.getMetaData();
	 		
	 		System.out.println("inside GetExcel data:::::::::::"+columns);
	 		String[] Header = new String[columns];
	 		for (int k=1;k<=columns;k++){
	 			//System.out.println("META DATA:::::::::::::"+md.getColumnName(k));
	 			Header[k-1]=md.getColumnName(k);
	 		}
	 		list.add(Header);
	 		while(rs.next())  {
	 			System.out.println("inside GetExcel data:::::::::::");
	 		String[] row = new String[columns];
	 	     for(int i=1; i<=columns; ++i){      
	 	      row[i-1]=rs.getString(i);
	 		}
	 	     list.add(row);
	 	    
	 		}
	 		
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
	 			if(con!=null){
					con.close();
				}
				if(rs!=null){
					rs.close();
				}
				if(stmt!=null){
					stmt.close();
				}
			    
	 		} 
	 	return list;
	 }
	 public Connection getConnection(String con,String username,String password) throws SQLException,ClassNotFoundException{
			Connection connection=null;
			Class.forName("oracle.jdbc.driver.OracleDriver"); 
			connection=DriverManager.getConnection(con,username,password);
			return connection;
		}
	
	 public void createExcel(int m ){
		 try {
			 String timeLog = new SimpleDateFormat("ddMMyyyy").format(new Date());
				
	         /*String filename = new String(prop.getProperty("filepath")+File.separator+prop.getProperty("date")+File.separator+prop.getProperty("filename"+m)+".xls");*/
	         String[] reportTypeList= prop.getProperty("report.type.list").split(",");
	         for (String reportType : reportTypeList){
	        	 System.out.println(timeLog);
	         ArrayList<ReportConfigBo> configList = fetchReportConfigFromDB(prop,reportType);
	         System.out.println("LIST SIZE"+configList.size());
	         System.out.println("exlFile===>"+reportType);
	         // This will output the full path where the file will be written to...
	         File exlFile = new File(prop.getProperty("R2.report.path")+"/"+reportType+"_"+timeLog+"_"+prop.getProperty("env"+m)+".xls");
	         System.out.println("exlFile===>"+exlFile.getPath());
				if (!exlFile.exists()) {
					exlFile.createNewFile();
				}
				
	         WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlFile);
	        
	         for (ReportConfigBo configBo: configList){
	        	 System.out.println("configBo.getSheetname():::::::::"+configBo.getSheetname());
	         WritableSheet writableSheet = writableWorkbook.createSheet(configBo.getSheetname(), 0);
	         //System.out.println("configBo.getQuery():::"+configBo.getQuery());
	         ArrayList<String[]> list =getExcelData(configBo.getQuery(),m);
	         System.out.println("list SIZE:::::::::"+list.size());
	         for(int i=0;i<list.size();i++){
	        	 String[] row = list.get(i);
	        	 for (int j=0;j<row.length;j++){
	        		 Label lb1= new Label(j,i,row[j]);
	        		 writableSheet.addCell(lb1);
	        	 }
	        		 
	        	
	         	}
	         if (list.size()<1){
	        	 System.out.println("List size is list.size()"+list.size());
	        	 Label lb1= new Label(1,1,"    ");
        		 writableSheet.addCell(lb1);
	         	}
	         } 

	         //Write and close the workbook
	         writableWorkbook.write();
	         writableWorkbook.close();
	         }
	     } catch (Exception e) {
	         e.printStackTrace();
	     } 
	 }
	 public ArrayList<ReportConfigBo> fetchReportConfigFromDB(Properties props,String ReportType) throws SQLException{

		 ArrayList<ReportConfigBo> list = new ArrayList<ReportConfigBo>();
		try{    
			//Connection con 
			PasswordSecurityService passwordSecurityService=SecurityServiceFactory.getPasswordSecurityService(); 
			Connection con=getConnection(props.getProperty("config.connection"),props.getProperty("config.username"),passwordSecurityService.decrypt(props.getProperty("config.password")));
			Statement stmt=con.createStatement(); 
			System.out.println("ReportType:::"+ReportType);
			String query= "Select * from BI_BR_REPORTS_QUERY where report_type='"+ReportType+"'  order by order_of_sheet desc ";
			ResultSet rs=stmt.executeQuery(query);
			
			while(rs.next())  {
				ReportConfigBo reportBo =new ReportConfigBo();
				System.out.println("QUERY:::::::::\n"+rs.getString(5));
				reportBo.setReportName(rs.getString(1));
				reportBo.setReportId(rs.getString(2));
				reportBo.setFilename(rs.getString(3));
				reportBo.setSheetname(rs.getString(4));
				reportBo.setQuery(rs.getString(6));
				reportBo.setReportType(rs.getString(5));
				list.add(reportBo);
			}
			if(con!=null){
				con.close();
			}
			if(rs!=null){
				rs.close();
			}
		    
			}catch(Exception e){
				e.printStackTrace();
			} 
		return list;
	}

}
