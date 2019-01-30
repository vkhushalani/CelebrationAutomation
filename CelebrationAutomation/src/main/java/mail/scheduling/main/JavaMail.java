package mail.scheduling.main;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;

import java.io.IOException;
import java.io.InputStream;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.Message.RecipientType;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.util.ByteArrayDataSource;
import javax.naming.InitialContext;
import javax.naming.NamingException;

import org.apache.poi.util.IOUtils;

public class JavaMail {
	
	private InternetAddress[] fromAddress;
	private InternetAddress[] toAddress;
	private String mailSubject;
	private String mailHtml;
	private String mailImage;
	private Session mailSession;
	private MimeMessage mimeMessage;
	
	public void setMailSession() throws NamingException {
		InitialContext ctx;
		ctx = new InitialContext();
		mailSession = (Session)ctx.lookup("java:comp/env/mail/Session");
	}
	
	public Session getMailSession(){
		return mailSession;
	}
	
	public void setMimeMessage() throws AddressException, MessagingException, IOException{
		
		mimeMessage = new MimeMessage(mailSession);
		mimeMessage.setFrom(fromAddress[0]);
        mimeMessage.setRecipients(RecipientType.TO, toAddress);
        mimeMessage.setSubject(mailSubject, "UTF-8");
        MimeMultipart multiPart = new MimeMultipart("related");
        MimeBodyPart part = new MimeBodyPart();
        part.setContent(mailHtml, "text/html; charset=utf-8");
        multiPart.addBodyPart(part);
        if(mailImage != null && mailImage.length() > 0)
        {
        	
        	part = new MimeBodyPart();
        	ClassLoader classloader = Thread.currentThread().getContextClassLoader();
        	InputStream imageStream = classloader.getResourceAsStream(mailImage);
             DataSource fds = new ByteArrayDataSource(IOUtils.toByteArray(imageStream), "image/gif");
             part.setDataHandler(new DataHandler(fds));
             part.setHeader("Content-ID","<image>");
             
            multiPart.addBodyPart(part);
        }
        mimeMessage.setContent(multiPart);
        mimeMessage.saveChanges();
	} 
	
	public MimeMessage getMimeMessage(){
		return mimeMessage;
	}
	
	public void sendMail() throws MessagingException{
		Transport transport = mailSession.getTransport();
		transport.connect();
		transport.sendMessage(mimeMessage, mimeMessage.getAllRecipients());
		transport.close();
	}

	public String getMailSubject() {
		return mailSubject;
	}

	public void setMailSubject(String mailSubject) {
		this.mailSubject = mailSubject;
	}

	public String getMailHtml() {
		return mailHtml;
	}

	public void setMailHtml(String mailHtml) {
		this.mailHtml = mailHtml;
	}

	public InternetAddress[] getFromAddress() {
		return fromAddress;
	}

	public void setFromAddress(InternetAddress[] fromAddress) {
		this.fromAddress = fromAddress;
	}

	public InternetAddress[] getToAddress() {
		return toAddress;
	}

	public void setToAddress(InternetAddress[] toAddress) {
		this.toAddress = toAddress;
	}

	public String getMailImage() {
		return mailImage;
	}

	public void setMailImage(String mailImage) {
		this.mailImage = mailImage;
	}

	
	

}
