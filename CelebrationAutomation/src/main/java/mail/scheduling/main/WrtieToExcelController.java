package mail.scheduling.main;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

import javax.naming.NamingException;

import org.apache.chemistry.opencmis.client.api.CmisObject;
import org.apache.chemistry.opencmis.client.api.Document;
import org.apache.chemistry.opencmis.client.api.Folder;
import org.apache.chemistry.opencmis.client.api.ItemIterable;
import org.apache.chemistry.opencmis.client.api.Session;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.sap.ecm.api.EcmService;

import mail.documentRepository.CustomDocument;
import mail.documentRepository.SharedConstants;

@RestController
public class WrtieToExcelController {
	private static String urlString =  "https://hrcngahr.wufoo.com/api/v3/forms/zv6fg4i1hs864p/entries.json";
	private static String authString = "Basic SDFOVS1GRVZILUlNQkQtNDkwNjpjb21faHJjX2ludF8xMjM=";

@GetMapping("/WriteToExcel")
public  ResponseEntity<?>  writeInExcel() throws NamingException, IOException {
		
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
				 if(MainFolder !=null){
				 
					 ItemIterable<CmisObject> children = MainFolder.getChildren();
						for (CmisObject o : children) {
					        if (o instanceof Document) {
					        	bDayExcel = (Document) o;
					        	break;
					        } 
					      }
						if(bDayExcel != null){
							 Document doc = cDoc.getDocumentBySession(openCmisSession, bDayExcel.getId());
							 doc.delete(true);}
				
				 }
				 else
				 { MainFolder = cDoc.createNewFolder(root,SharedConstants.FOLDER_NAME);}
					 // read from the API
					 URL url = new URL(urlString);
					 HttpURLConnection conn = (HttpURLConnection) url.openConnection();
					 conn.setRequestMethod("GET");
						conn.setRequestProperty("Accept", "application/json");
						conn.setRequestProperty("Authorization", authString);

						if (conn.getResponseCode() != 200) {
							throw new RuntimeException("Failed : HTTP error code : "
									+ conn.getResponseCode());
						}
						StringBuffer sb = new StringBuffer();
						BufferedReader br = new BufferedReader(new InputStreamReader(
							(conn.getInputStream())));

						String output;
						while ((output = br.readLine()) != null) {
							sb.append(output);
						}
						conn.disconnect();
						JSONObject myResponse = new JSONObject(sb.toString());
					 // loop through the data
						JSONArray entries = myResponse.getJSONArray("Entries");
						List<Person> team = new ArrayList<Person>();
						for(int i = 0; i < entries.length(); i++)
						{
						   JSONObject jObj = entries.getJSONObject(i);
						   	Person nPerson = new Person();
							 nPerson.setEmail(jObj.getString("Field5"));
							 nPerson.setName(jObj.getString("Field2") + " " + jObj.getString("Field3"));
							 String bDay =jObj.getString("Field8");
							 String[] parts  = bDay.split("-");
							 nPerson.setBday(parts[2]+"."+parts[1]+".00");
							 nPerson.setMobile(jObj.getString("Field6"));
							 nPerson.setAuth(jObj.getString("Field10"));
							 team.add(nPerson);
						}
					
					
					
		// remove the excel 
					
					
		// save a new excel	
					 XSSFWorkbook newWorkbook = new XSSFWorkbook();
				     XSSFSheet newsheet = newWorkbook.createSheet("Birthday");
				     int rowNum = 0;
				     for(Person p:team)
				     {
				    	 Row row = newsheet.createRow(rowNum++);
				    	 int colNum = 0;
				    	 Cell cell1 = row.createCell(colNum++);
				    	 cell1.setCellValue(p.getName());
				    	 Cell cell2 = row.createCell(colNum++);
				    	 cell2.setCellValue(p.getEmail());
				    	 Cell cell3 = row.createCell(colNum++);
				    	 cell3.setCellValue(p.getMobile());
				    	 Cell cell4 = row.createCell(colNum++);
				    	 cell4.setCellValue(p.getBday());
				    	 Cell cell5 = row.createCell(colNum++);
				    	 cell5.setCellValue(p.getAuth());
				     }
				     ByteArrayOutputStream baos = new ByteArrayOutputStream();
				     newWorkbook.write(baos);
				     newWorkbook.close();
				     byte[] xls = baos.toByteArray();
					 String mimeType = URLConnection.guessContentTypeFromName(SharedConstants.FILE_NAME+".xlsx");
				        if (mimeType == null || mimeType.length() == 0) {
				            mimeType = "application/octet-stream";
				        }
				     Timestamp timestamp = new Timestamp(System.currentTimeMillis()); 
					 cDoc.createNewDocument(openCmisSession, MainFolder, SharedConstants.FILE_NAME+"#"+timestamp.getTime(),xls,mimeType); 
	
					 return ResponseEntity.ok().body("Success");
}

}
