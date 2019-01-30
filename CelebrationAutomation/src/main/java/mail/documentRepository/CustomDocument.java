package mail.documentRepository;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.Map;

import javax.naming.InitialContext;
import javax.naming.NamingException;

import org.apache.chemistry.opencmis.client.api.CmisObject;
import org.apache.chemistry.opencmis.client.api.Document;
import org.apache.chemistry.opencmis.client.api.Folder;
import org.apache.chemistry.opencmis.client.api.ItemIterable;
import org.apache.chemistry.opencmis.client.api.Session;
import org.apache.chemistry.opencmis.commons.PropertyIds;
import org.apache.chemistry.opencmis.commons.data.ContentStream;
import org.apache.chemistry.opencmis.commons.enums.VersioningState;
import org.apache.chemistry.opencmis.commons.exceptions.CmisObjectNotFoundException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sap.ecm.api.EcmService;
import com.sap.ecm.api.RepositoryOptions;
import com.sap.ecm.api.RepositoryOptions.Visibility;

public class CustomDocument {
	private String uniqueName ;
	private String uniqueKey;

	Logger logger = LoggerFactory.getLogger(CustomDocument.class);
	public EcmService getECMService() throws NamingException{
        InitialContext ctx = new InitialContext();
        String lookupName = "java:comp/env/" + "EcmService";
        EcmService ecmSvc = (EcmService) ctx.lookup(lookupName);
        return ecmSvc;
	}
	public RepositoryOptions getRepository(){
		
		RepositoryOptions options = new RepositoryOptions();
        options.setUniqueName(uniqueName);
        options.setRepositoryKey(uniqueKey);
        options.setVisibility(Visibility.PROTECTED);
        return options;
	}
	public EcmService CreateRepository(EcmService ecmSvc,RepositoryOptions rep){
		ecmSvc.createRepository(rep);
		return ecmSvc;
		}
	public Folder createNewFolder(Folder folder,String fname){
		
		Map<String, String> newFolderProps = new HashMap<String, String>();
	      newFolderProps.put(PropertyIds.OBJECT_TYPE_ID, "cmis:folder");
	      newFolderProps.put(PropertyIds.NAME, fname);
	      return folder.createFolder(newFolderProps);
	}
	public Document createNewDocument (Session openCmisSession,Folder folder,String docName,byte[] Content, String mimeType) throws UnsupportedEncodingException{
		 Map<String, Object> properties = new HashMap<String, Object>();
	      properties.put(PropertyIds.OBJECT_TYPE_ID, "cmis:document");
	      properties.put(PropertyIds.NAME, docName);
	      InputStream stream = new ByteArrayInputStream(Content);
	      ContentStream contentStream = openCmisSession.getObjectFactory()
	                                    .createContentStream(docName,
	                                    Content.length,mimeType,stream);
	     
	       return folder.createDocument(properties,contentStream, VersioningState.NONE);
		
	}
	public Session createSession(EcmService ecmSvc){
		 try
		 {
			 return ecmSvc.connect(uniqueName,uniqueKey); 
		 }
		 catch (CmisObjectNotFoundException e) {
			 
			 RepositoryOptions rep = getRepository();
			 ecmSvc = CreateRepository(ecmSvc, rep);
			 return ecmSvc.connect(uniqueName,uniqueKey); 
			 
		 }
	}
	public String getUniqueName() {
		return uniqueName;
	}

	public void setUniqueName(String uniqueName) {	
		this.uniqueName = uniqueName;
	}

	public String getUniqueKey() {
		return uniqueKey;
	}

	public void setUniqueKey(String uniqueKey) {
		this.uniqueKey = uniqueKey;
	}
	public Folder getFolder(Folder folder, String folderName) {
		ItemIterable<CmisObject> children = folder.getChildren();
		for (CmisObject o : children) {
	        if (o instanceof Folder) {
	        	logger.debug(o.getName());
	          if(o.getName().equalsIgnoreCase(folderName))
	          { return (Folder)o;}
	        } 
	      }
		return null;
	}
	
	public Document getDocumentBySession(Session openCmisSession,String docId){
		Document doc = (Document) openCmisSession.getObject(docId);
        return doc;
		
	}
	public Document getDocument(Folder folder,String docId){
		ItemIterable<CmisObject> children = folder.getChildren();
		for (CmisObject o : children) {
	        if (o instanceof Document) {
	        	logger.debug(o.getName());
	          if(o.getId().equalsIgnoreCase(docId))
	          { return (Document) o;}
	        } 
	      }
		return null;
	}
	
	public ContentStream getDocumentStreamById(Session openCmisSession,String documentId) {
        ContentStream contentStream = null;
        Document doc = getDocumentBySession(openCmisSession,documentId);
        if (doc != null) {
            contentStream = doc.getContentStream();
        }
        return contentStream;
    }
	
	

}
