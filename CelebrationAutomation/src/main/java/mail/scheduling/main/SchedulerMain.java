package mail.scheduling.main;


import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.mail.MessagingException;

import javax.mail.internet.InternetAddress;
import javax.naming.NamingException;

import org.apache.chemistry.opencmis.client.api.CmisObject;
import org.apache.chemistry.opencmis.client.api.Document;
import org.apache.chemistry.opencmis.client.api.Folder;
import org.apache.chemistry.opencmis.client.api.ItemIterable;
import org.apache.chemistry.opencmis.client.api.Session;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.monitorjbl.xlsx.StreamingReader;
import com.sap.ecm.api.EcmService;

import mail.documentRepository.CustomDocument;
import mail.documentRepository.SharedConstants;

@Component
public class SchedulerMain {
	
	static Logger logger = LoggerFactory.getLogger(SchedulerMain.class);
//	
	
//	@Scheduled(cron = "0 0 10 * * MON-FRI")
	@Scheduled(fixedRate = 24*60*60*1000, initialDelay = 60000)
	public void sendNotificationMail() throws NamingException, MessagingException, IOException{		
		// read excel file from the folder
		
		CustomDocument cDoc = new CustomDocument();
		 cDoc.setUniqueKey(SharedConstants.REP_KEY);
		 cDoc.setUniqueName(SharedConstants.REP_NAME);
		 
		 Session openCmisSession = null;
		 
		 EcmService ecmSvc = cDoc.getECMService();
		 openCmisSession =cDoc.createSession(ecmSvc);
		 
		 Folder root = openCmisSession.getRootFolder();
		 Folder MainFolder =  cDoc.getFolder(root,SharedConstants.FOLDER_NAME);
		 Document bDayExcel = null ;
		 
		 ItemIterable<CmisObject> children = MainFolder.getChildren();
			for (CmisObject o : children) {
		        if (o instanceof Document) {
		        	bDayExcel = (Document) o;
		        	break;
		        } 
		      }
		 
			InputStream excelFile = bDayExcel.getContentStream().getStream();
			 Workbook workbook = StreamingReader.builder()
				        .rowCacheSize(16384)    // number of rows to keep in memory (defaults to 10)
				        .bufferSize(8192)     // buffer size to use when reading InputStream to file (defaults to 1024)
				        .open(excelFile); 
			 Sheet sheetBday = workbook.getSheet("Birthday");
			 int index = 0;
			  
			 List<Person> team = new ArrayList<Person>();
			 
			 for (Row row : sheetBday) {
				 Person person = new Person();
			    for (Cell cCell : row) {
					    	if ( index == 0)
					    	 {
					    		person.setName(cCell.getStringCellValue());
					    		
					    		
			                 }
						    else if (index == 1)
			               	 {
						    	person.setEmail(cCell.getStringCellValue());
						    	
			               	 }
						    else if (index == 2)
						    {
						    person.setMobile(cCell.getStringCellValue());
						    }
	                    	else if (index == 3)
	                    	 {
	                    		person.setBday(cCell.getStringCellValue());
	                    	 }
	                    	else if (index == 4)
	                    	 {
	                    		person.setAuth(cCell.getStringCellValue());
	                    	 }
					    	                    	 
			    	index = index +1;
			    }
			    team.add(person);
			     
			    index = 0;
			  }
			// sending reminder mails
				Calendar todayCal = Calendar.getInstance();
				Boolean sendMailFlag = false;
				 String[] strDays = new String[] { "Sunday", "Monday", "Tuesday", "Wednesday", "Thusday",
					        "Friday", "Saturday" };
				 String[] strMonths = new String[] { "Jan", "Feb", "Mar", "Apr", "May",
					        "Jun", "Jul","Aug","Sep","Oct","Nov","Dec" };
					if(todayCal.get(Calendar.DAY_OF_WEEK) != 1 && todayCal.get(Calendar.DAY_OF_WEEK) != 7){ 
				String mailHtml = "Hi,</br>";
				mailHtml = mailHtml + "<p>Upcoming Celebration Dates:</p>";
				mailHtml = mailHtml + "<table style='width:100%; border:1px solid black; border-collapse:collapse'>";
				mailHtml = mailHtml + "<tr style='border:1px solid black; border-collapse:collapse; text-align: left' >";
				mailHtml = mailHtml + "<th style='border:1px solid black; border-collapse:collapse; text-align: left ' >Name</th>";
				mailHtml = mailHtml + "<th style='border:1px solid black; border-collapse:collapse; text-align: left' >Birth Date</th>";
				mailHtml = mailHtml + "<th style='border:1px solid black; border-collapse:collapse; text-align: left' >Birth WeekDay</th>";
				mailHtml = mailHtml + "<th style='border:1px solid black; border-collapse:collapse; text-align: left' >Celebration Date</th>";
				mailHtml = mailHtml + "<th style='border:1px solid black; border-collapse:collapse; text-align: left' >Celebration WeekDay</th>";
				mailHtml = mailHtml + "<th style='border:1px solid black; border-collapse:collapse; text-align: left' >Days Left(Excluding Weekend)</th>";
				mailHtml = mailHtml + "</tr>";
				
				String reminderMailTo = "";
				
				for (Person person:team){
					todayCal = Calendar.getInstance();
					String[] parts = person.getBday().split("\\.");
					if(person.getAuth().equalsIgnoreCase("ADMIN"))
					{ reminderMailTo = reminderMailTo + person.getEmail()+"," ;}
					logger.debug("bday"+parts[0]+parts[1]+parts[2]);
					if(parts.length >0){
						Calendar bcal = Calendar.getInstance();
						bcal.set(bcal.get(Calendar.YEAR), Integer.parseInt(parts[1]) - 1, Integer.parseInt(parts[0])); //Year, month and day of month
						
						int aDayofWeek = bcal.get(Calendar.DAY_OF_WEEK);
						
						if(aDayofWeek == Calendar.SUNDAY)
						{bcal.add(Calendar.DATE, 1);}
						else if (aDayofWeek == Calendar.SATURDAY)
						{bcal.add(Calendar.DATE, 2);}	
						
						bcal.set(Calendar.HOUR_OF_DAY, 0);  
						bcal.set(Calendar.MINUTE, 0);  
						bcal.set(Calendar.SECOND, 0);  
						bcal.set(Calendar.MILLISECOND, 0);
						
						todayCal.set(Calendar.HOUR_OF_DAY, 0);  
						todayCal.set(Calendar.MINUTE, 0);  
						todayCal.set(Calendar.SECOND, 0);  
						todayCal.set(Calendar.MILLISECOND, 0);
						
						int diffInDays = 0;
						if(todayCal.before(bcal)){
						while(!todayCal.equals(bcal))
			            {
			                int day = todayCal.get(Calendar.DAY_OF_WEEK);
			                if ((day != Calendar.SATURDAY) && (day != Calendar.SUNDAY))
			                {
			                	diffInDays++;
			                	}
			                todayCal.add(Calendar.DATE, 1);
			            }}
						else
						{
							diffInDays = -1;
						}
						
						if(diffInDays > 0 && diffInDays < 7){
						sendMailFlag = true;
						
						mailHtml = mailHtml + "<tr style='border:1px solid black; border-collapse:collapse; text-align: left' >";
						mailHtml = mailHtml + "<td style='border:1px solid black; border-collapse:collapse; text-align: left' >"+person.getName()+"</td>";
						mailHtml = mailHtml + "<td style='border:1px solid black; border-collapse:collapse; text-align: left' >"+Integer.valueOf(parts[0])+" "+strMonths[(Integer.valueOf(parts[1])-1)]+"</td>";
						mailHtml = mailHtml + "<td style='border:1px solid black; border-collapse:collapse; text-align: left' >"+strDays[aDayofWeek - 1]+"</td>";
						mailHtml = mailHtml + "<td style='border:1px solid black; border-collapse:collapse; text-align: left' >"+bcal.get(Calendar.DATE)+" "+ strMonths[bcal.get(Calendar.MONTH)]+"</td>";
						mailHtml = mailHtml + "<td style='border:1px solid black; border-collapse:collapse; text-align: left' >"+strDays[bcal.get(Calendar.DAY_OF_WEEK) - 1]+"</td>";
						mailHtml = mailHtml + "<td style='border:1px solid black; border-collapse:collapse; text-align: left' >"+diffInDays+"</td>";
						mailHtml = mailHtml + "</tr>";
						}
					}
				}
				reminderMailTo = reminderMailTo.substring(0, reminderMailTo.length() - 1);
				mailHtml = mailHtml + "</table>";
				
				if(sendMailFlag){
				JavaMail jMail = new JavaMail();
			
				String from = "donotreply.hrcngahr@gmail.com";
				
				jMail.setMailSubject("Birthday Celebration Reminder - HRC Team");
				jMail.setFromAddress(InternetAddress.parse(from));
		        jMail.setToAddress(InternetAddress.parse(reminderMailTo));
		        jMail.setMailHtml(mailHtml);
		        jMail.setMailSession();
		        jMail.setMimeMessage();
		        jMail.sendMail();
		      
		        }}
					
					// sending birthday emails to all team members
					String bMailHtml = null; 
					String mailSubject = null;
					String mailTo = "";
					Calendar calToday = Calendar.getInstance();
					Date todayDate = calToday.getTime();
					String day = new SimpleDateFormat("dd").format(todayDate);
					String month = new SimpleDateFormat("MM").format(todayDate);
					bMailHtml = "<p style='color:rgb(237,76,92)'>";
					bMailHtml = bMailHtml +"May this birthday be filled with lots of happy hours and also your life with many happy birthdays, that are yet to come.";
					bMailHtml = bMailHtml +"</p>";
					bMailHtml = bMailHtml +"<img src=\"cid:image\" alt='Happy Birthday'><br><br><br>";
					bMailHtml = bMailHtml +"<div style='color:orange'>- HRC Team</div>";
					
					for (Person person : team)
					{	
						mailTo = mailTo + person.getEmail()+",";
					}
					mailTo = mailTo.substring(0, mailTo.length() - 1);
					
					for (Person person : team)
						
					{	
						mailSubject = "Happy Birthday - ";
						String[] parts = person.getBday().split("\\.");
						Calendar bcal = Calendar.getInstance();
						bcal.set(bcal.get(Calendar.YEAR), Integer.parseInt(parts[1]) - 1, Integer.parseInt(parts[0])); //Year, month and day of month
						if(bcal.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY)
						{ bcal.add(Calendar.DATE, -1);
						mailSubject = "Happy Birthday In Advance -";
						}
						else if (bcal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY)
						{ bcal.add(Calendar.DATE, -2);
						mailSubject = "Happy Birthday In Advance - ";}	
						
					String bday = new SimpleDateFormat("dd").format(bcal.getTime());
					String bmonth = new SimpleDateFormat("MM").format(bcal.getTime());
					if(day.equalsIgnoreCase(bday) && month.equalsIgnoreCase(bmonth))
					{	mailSubject = mailSubject + person.getName() + " (" + parts[0] + "/"+parts[1]+")"; 
					 
						JavaMail jMail = new JavaMail();
						
						String from = "donotreply.hrcngahr@gmail.com";
						
						jMail.setMailSubject(mailSubject);
						jMail.setFromAddress(InternetAddress.parse(from));
				        jMail.setToAddress(InternetAddress.parse(mailTo));
				        jMail.setMailHtml(bMailHtml);
				        jMail.setMailImage("images/happyBirthday.gif");
				        jMail.setMailSession();
				        jMail.setMimeMessage();
				        jMail.sendMail();
					}
					}
		
			
		 
	}
}
