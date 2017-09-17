package com.coe.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.mail.Authenticator;
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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.coe.model.User;

public class Utility {

	public  SimpleDateFormat sdf;
	public  File file;

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		Utility utility = new Utility();
		if (args.length == 2) {
			try {
				utility.sdf = new SimpleDateFormat(args[0]);
				utility.file = new File(args[1]);
			} catch (Exception e) {
				System.out.println("Incorrect parameters!!!");
				System.out.println("arg[0]: Simple Date Format");
				System.out.println("arg[1]: Absolute File Path");
			}
		}else if(args.length==1 || args.length>2){
			System.out.println("Incorrect parameters!!!");
			System.out.println("arg[0]: Simple Date Format");
			System.out.println("arg[1]: Absolute File Path");
		}else if (args.length==0){
			System.out.println("No arguments passed. Using default parameters as follows:");
			System.out.println("arg[0]: Simple Data Format = MM-dd-yyyy");
			System.out.println("arg[1]: Absolute File Path = c:/data/birthdays.xls");
			utility.sdf = new SimpleDateFormat("MM-dd-yyyy");
			utility.file = new File("c:/data/birthdays.xls");
		}

		InputStream inputStream = null;
		Workbook myWorkBook = null;
		List<User> userList = new ArrayList<User>();

		try {
			inputStream = new FileInputStream(utility.file);
			myWorkBook = new HSSFWorkbook(inputStream);
			Sheet mySheet = myWorkBook.getSheetAt(0);

			Iterator<Row> rowIter = mySheet.rowIterator();
			int row = 0;
			while (rowIter.hasNext()) {
				Row myRow = rowIter.next();

				if (myRow.getRowNum() == 0) {
					continue;
				}

				Cell myCell = null;

				Calendar today = Calendar.getInstance();
				myCell = myRow.getCell(2, MissingCellPolicy.RETURN_BLANK_AS_NULL);
				if (myCell == null || myCell.getStringCellValue().equals("")) {
					continue;
				}
				Calendar birthdate = Calendar.getInstance();
				birthdate.setTime(utility.sdf.parse(myCell.getStringCellValue()));

				if (today.get(Calendar.MONTH) == birthdate.get(Calendar.MONTH)
						&& today.get(Calendar.DATE) == birthdate.get(Calendar.DATE)) {

					User user = new User();

					myCell = myRow.getCell(1, MissingCellPolicy.RETURN_BLANK_AS_NULL);
					if (myCell == null || myCell.getStringCellValue().equals("")) {
						continue;
					}
					user.setName(myCell.getStringCellValue());

					myCell = myRow.getCell(2, MissingCellPolicy.RETURN_BLANK_AS_NULL);
					if (myCell == null || myCell.getStringCellValue().equals("")) {
						continue;
					}
					user.setBirthday(utility.sdf.parse(myCell.getStringCellValue()));

					myCell = myRow.getCell(3, MissingCellPolicy.RETURN_BLANK_AS_NULL);
					if (myCell == null || myCell.getStringCellValue().equals("")) {
						continue;
					}
					user.setEmail(myCell.getStringCellValue());

					userList.add(user);
				}
			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		finally {
			inputStream.close();
		}

		for (User user : userList) {
			System.out.println(user);
			try {
				utility.sendEmail(user);
			} catch (MessagingException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	public void sendEmail(User user) throws MessagingException, IOException {

		final String userName = "user@gmail.com";
		final String password = "password";

		// sets SMTP server properties
		Properties properties = new Properties();
		properties.put("mail.smtp.host", "smtp.gmail.com");
		properties.put("mail.smtp.port", "587");
		properties.put("mail.smtp.auth", "true");
		properties.put("mail.smtp.starttls.enable", "true");
		properties.put("mail.user", userName);
		properties.put("mail.password", password);

		Authenticator auth = new Authenticator() {
			public PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(userName, password);
			}
		};

		Session session = Session.getInstance(properties, auth);

		// ContentID is used by both parts
		String cid1 = ContentIdGenerator.getContentId();
		String cid2 = ContentIdGenerator.getContentId();

		// creates a new e-mail message
		Message msg = new MimeMessage(session);

		msg.setFrom(new InternetAddress(userName));
		InternetAddress[] toAddresses = { new InternetAddress(user.getEmail()) };
		msg.setRecipients(Message.RecipientType.TO, toAddresses);
		msg.setSubject("Birthday Greetings");
		msg.setSentDate(new Date());

		// creates message part
		MimeBodyPart messageBodyPart = new MimeBodyPart();
		String content = getTemplate(cid1, cid2).replace("[$Name$]", user.getName());
		messageBodyPart.setContent(content, "text/html");

		// creates multi-part
		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(messageBodyPart);

		// Image part
		MimeBodyPart imagePart = new MimeBodyPart();
		//imagePart.attachFile("resources/logo.jpg");
		
		InputStream logoStream = getClass().getClassLoader().getResource("resource/logo.jpg").openStream();
		byte[] buffer = new byte[logoStream.available()];
		logoStream.read(buffer);
		
		File targetFile = new File("logo.jpg");
	    OutputStream outStream = new FileOutputStream(targetFile);
	    outStream.write(buffer);
		outStream.close();
		imagePart.attachFile("logo.jpg");
		
		imagePart.setContentID("<" + cid1 + ">");
		imagePart.setDisposition(MimeBodyPart.INLINE);
		multipart.addBodyPart(imagePart);

		MimeBodyPart imagePart2 = new MimeBodyPart();
		//imagePart2.attachFile("resources/greeting.jpg");
		
		InputStream greetingStream = getClass().getClassLoader().getResource("resource/greeting.jpg").openStream();
		byte[] buffer2 = new byte[greetingStream.available()];
		greetingStream.read(buffer2);
		
		File targetFile2 = new File("greeting.jpg");
	    OutputStream outStream2 = new FileOutputStream(targetFile2);
	    outStream2.write(buffer2);
		outStream2.close();
		imagePart2.attachFile("greeting.jpg");

		
		imagePart2.setContentID("<" + cid2 + ">");
		imagePart2.setDisposition(MimeBodyPart.INLINE);
		multipart.addBodyPart(imagePart2);

		// sets the multi-part as e-mail's content
		msg.setContent(multipart);

		// sends the e-mail
		Transport.send(msg);

	}

	private  String getTemplate(String cid1, String cid2) {
		return "<style>.logo{	float:right;}.content{	text-align: center;	color:#ff471a;	clear: both;	margin: auto;"
				+ "font-size: 26px;}.heading{	text-align: left;	color:#ff471a;	clear: both;	margin: auto;"
				+ "font-size: 30px;}.greeting{	display: block;" + "margin: auto;}</style><img src=\"cid:" + cid1
				+ "\" alt=\"Harsha Abakus Solar\" class=\"logo\"><div style=\"clear: both;\"><div class=\"heading\">"
				+ "Dearest [$Name$],</div><br><br><div class=\"content\">May your birthday greets you with abundance "
				+ "of joy, good fortune, health, wealth and happiness.</div><br><br>"
				+ "<div class=\"content\">May you continue to achieve success. The birthday greetings are flowing your "
				+ "way, we wish you a cheerful and a blessed birthday.</div>" + "<img src=\"cid:" + cid2
				+ "\" alt=\"Happy Birthday\" class=\"greeting\"><div class=\"content\">Regards,</div><div class=\"content\">Team HR</div></div>";
	}
}
