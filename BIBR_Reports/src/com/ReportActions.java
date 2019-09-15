package com;


import java.sql.Connection;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import com.deloitte.passwordSecurity.service.PasswordSecurityService;
import com.deloitte.passwordSecurity.service.SecurityServiceFactory;

public class ReportActions {
	public static void main(String[] args){
		Connection con=null;
		try{    
			ReportDao dao=new ReportDao();
			Properties props=dao.getPropValues();
			
		        
			int i=1;
			PasswordSecurityService passwordSecurityService=SecurityServiceFactory.getPasswordSecurityService(); 
			while(props.getProperty("connectionString"+i)!=null){
				dao.createExcel(i );
				
			//con=dao.getConnection(props.getProperty("connectionString"+i),props.getProperty("username"+i),passwordSecurityService.decrypt(props.getProperty("password"+i)));
			
			
			/*ArrayList<String[]> detail=dao.getExcelData(props, con,"detail.query");
			//dao.update100ErrorsReport(detail, props, i);
			
			ArrayList<String[]> summary=dao.getExcelData(props, con,"summary.query");*/
			
			//con.close();
			i++;
		}
			dao.sendEmail(props);
			}catch(Exception e){
				e.printStackTrace();
			}  
			  
			}  
}
