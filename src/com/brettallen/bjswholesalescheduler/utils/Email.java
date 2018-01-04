package com.brettallen.bjswholesalescheduler.utils;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class Email
{
	private final String USER_SENDER = "brettallen.business";
	private final String EMAIL_SENDER = "brettallen.business@gmail.com";
	public static final String DEV_INFO = "Contact Developer (Brett Allen): brettallen.business@gmail.com\n\n";

    private final Session session;

	public boolean errorInEmailProcess;
	
	public Email(final String password)
	{
		errorInEmailProcess = false;

        Properties props = new Properties();
		
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.host", "smtp.gmail.com");
		props.put("mail.smtp.port", "465");
		props.put("mail.smtp.socketFactory.port", "465");
		props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		
		session = Session.getDefaultInstance(props, new Authenticator() 
		{
			@Override
			protected PasswordAuthentication getPasswordAuthentication()
			{
				return new PasswordAuthentication(USER_SENDER, password);
			}			
		});
	}
	
	public void sendGenericMessage(String sendTo, String subject, String textMessage) throws MessagingException
	{
		MimeMessage msg = new MimeMessage(session);

		msg.setFrom(new InternetAddress(EMAIL_SENDER));
		msg.setRecipients(Message.RecipientType.TO, InternetAddress.parse(sendTo));
		msg.setSubject(subject);
		msg.setText(textMessage);

		Transport.send(msg);

		System.out.println("Your message was sent successfully...");
	}
	
	public void sendAttachmentMessage(String sendTo, String employeeName, String subject, String[] fileNames) throws MessagingException
	{
		Message msg = new MimeMessage(session);

		msg.setFrom(new InternetAddress(EMAIL_SENDER));
		msg.setRecipients(Message.RecipientType.TO, InternetAddress.parse(sendTo));
		msg.setSubject(subject);

		BodyPart msgBodyPart = new MimeBodyPart();
		msgBodyPart.setText("Here is your schedule " + employeeName + "!");

		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(msgBodyPart);

		//Attach all fileNames
		for(String fileName : fileNames)
			if (!fileName.isEmpty())
				addAttachment(multipart, fileName);

		//Set content of message to contain all parts including attachment
		msg.setContent(multipart);

		//Send message
		Transport.send(msg);

		System.out.println("\nYour message was sent successfully...");
	}

	//Add attachments to attachment messages
	private void addAttachment(Multipart multipart, String fileName) throws MessagingException
	{
		DataSource source = new FileDataSource(fileName);
		BodyPart messageBodyPart = new MimeBodyPart();
		messageBodyPart.setDataHandler(new DataHandler(source));
		messageBodyPart.setFileName(fileName);
		multipart.addBodyPart(messageBodyPart);
	}
}
