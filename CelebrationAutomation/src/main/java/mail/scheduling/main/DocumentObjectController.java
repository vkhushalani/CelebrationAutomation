package mail.scheduling.main;


import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLConnection;
import java.sql.Timestamp;
//import java.util.Calendar;

import javax.naming.NamingException;
import javax.servlet.http.HttpServletRequest;

import org.apache.chemistry.opencmis.client.api.Document;
import org.apache.chemistry.opencmis.client.api.Folder;
import org.apache.chemistry.opencmis.client.api.Session;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.sap.ecm.api.EcmService;

import mail.documentRepository.CustomDocument;
import mail.documentRepository.SharedConstants;



@RestController
public class DocumentObjectController {
	Logger logger = LoggerFactory.getLogger(DocumentObjectController.class);
	
	 @DeleteMapping("/attachments/{docId}")
	 public ResponseEntity<?> deleteAttachment(@PathVariable("docId") String docId) throws NamingException {
		 
		 CustomDocument cDoc = new CustomDocument();
		 cDoc.setUniqueKey(SharedConstants.REP_KEY);
		 cDoc.setUniqueName(SharedConstants.REP_NAME);
		 Session openCmisSession = null;
		 
		 EcmService ecmSvc = cDoc.getECMService();
		 openCmisSession = ecmSvc.connect(cDoc.getUniqueName(),cDoc.getUniqueKey());
		 
		 Document doc = cDoc.getDocumentBySession(openCmisSession, docId);
		 doc.delete(true);
		 return ResponseEntity.ok().body("Successfully Deleted DocId:" + docId);
		 
	 }
	 @PostMapping("/attachment")
	 public ResponseEntity<AttachmentsObject> uploadAttachment(@RequestParam("file") MultipartFile uploadfile,HttpServletRequest request) throws NamingException, UnsupportedEncodingException, IOException{
		 CustomDocument cDoc = new CustomDocument();
		 cDoc.setUniqueKey(SharedConstants.REP_KEY);
		 cDoc.setUniqueName(SharedConstants.REP_NAME);
		 
		 Session openCmisSession = null;
		 
		 EcmService ecmSvc = cDoc.getECMService();
		 openCmisSession =cDoc.createSession(ecmSvc);
		 
		 Folder root = openCmisSession.getRootFolder();
		 Folder AttFolder =  cDoc.getFolder(root,SharedConstants.FOLDER_NAME);
		 if(AttFolder == null){
			 AttFolder = cDoc.createNewFolder(root,SharedConstants.FOLDER_NAME);
		 }
		 String mimeType = URLConnection.guessContentTypeFromName(uploadfile.getName());
	        if (mimeType == null || mimeType.length() == 0) {
	            mimeType = "application/octet-stream";
	        }
	     Timestamp timestamp = new Timestamp(System.currentTimeMillis()); 
		 Document doc = cDoc.createNewDocument(openCmisSession, AttFolder, uploadfile.getOriginalFilename()+"#"+timestamp.getTime(),uploadfile.getBytes(),mimeType); 
		 AttachmentsObject att = new AttachmentsObject();
		 att.setDocId(doc.getId());
		 att.setDocName(doc.getName());
		 return ResponseEntity.ok().body(att);

	 }
}
