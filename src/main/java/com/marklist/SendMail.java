package com.marklist;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class SendMail {

	public void send() throws IOException {

		String to = "******@gmail.com";

		String from = "nickyjacobsam@gmail.com";

		String host = "smtp.gmail.com";

		Properties properties = System.getProperties();

		properties.put("mail.smtp.host", host);
		properties.put("mail.smtp.port", "465");
		properties.put("mail.smtp.ssl.enable", "true");
		properties.put("mail.smtp.auth", "true");

		Session session = Session.getInstance(properties, new javax.mail.Authenticator() {
			@Override
			protected PasswordAuthentication getPasswordAuthentication() {

				return new PasswordAuthentication("*******@gmail.com", "**********");

			}

		});

		try {

			MimeMessage message = new MimeMessage(session);

			message.setFrom(new InternetAddress(from));

			message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

			message.setSubject("Mark list of all students");

			Multipart multipart = new MimeMultipart();

			MimeBodyPart attachmentPart = new MimeBodyPart();
			MimeBodyPart attachmentPart2 = new MimeBodyPart();
			MimeBodyPart attachmentPart3 = new MimeBodyPart();

			MimeBodyPart textPart = new MimeBodyPart();

			
				String f1 = "C:\\Users\\nicky\\eclipse-workspace\\com.njs.markList\\Mark_list.xlsx";
				String f2 = "C:\\Users\\nicky\\eclipse-workspace\\com.njs.markList\\passed_students_list.xlsx";
				String f3 = "C:\\Users\\nicky\\eclipse-workspace\\com.njs.markList\\failed_students_list.xlsx";

				File filel = new File(f1);
				File file2 = new File(f2);
				File file3 = new File(f3);

				attachmentPart.attachFile(filel);
				attachmentPart2.attachFile(file2);
				attachmentPart3.attachFile(file3);
				textPart.setText("Pass percentage of division A: 100% \n Pass percentage of division B: 37.5%");
				multipart.addBodyPart(textPart);
				multipart.addBodyPart(attachmentPart);
				multipart.addBodyPart(attachmentPart2);
				multipart.addBodyPart(attachmentPart3);

			message.setContent(multipart);

			System.out.println("sending...");
			Transport.send(message);
			System.out.println("Sent message successfully....");
		} catch (MessagingException mex) {
			mex.printStackTrace();
		}

	}

}
