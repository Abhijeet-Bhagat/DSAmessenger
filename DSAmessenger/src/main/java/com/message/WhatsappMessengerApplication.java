package com.message;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;

import com.twilio.Twilio;
import com.twilio.rest.api.v2010.account.Message;
import com.twilio.type.PhoneNumber;

@SpringBootApplication
@EnableScheduling

public class WhatsappMessengerApplication {

	public static final String ACCOUNT_SID = "ACa96469a5301b44b7d434742372ecb046";
	public static final String AUTH_TOKEN = "7042aa089c22cf75869315b317c57229";

	// @Scheduled(cron = "0 0 8 * * ?") // Execute daily at 8:00 AM
	public static void main(String[] args) throws FileNotFoundException, IOException {
		SpringApplication.run(WhatsappMessengerApplication.class, args);
		QuestionInfo qinfo = new QuestionInfo();
		try {

			System.err.println("Method Started :: -------------------{}");
			readQuestionfromSheet();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// @Scheduled(fixedRate = 3000)
	@Scheduled(cron = "0 0 8 * * ?") // Execute daily at 8:00 AM
	public static void readQuestionfromSheet() throws FileNotFoundException, IOException {
		System.err.println("Sending Started :: -------------------{}");
		
        File file = new File("src/main/resources/Competitive_programming.xlsx");

		FileInputStream fileInputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(fileInputStream);

		Sheet sheet = workbook.getSheetAt(0);
		QuestionInfo qinfo = new QuestionInfo();
		ArrayList<String> question = new ArrayList<>();
		ArrayList<String> link = new ArrayList<>();
		// function to check which sent column has last yes
		int startingRowNum = 0;
		for (Row row : sheet) {

			System.out.println(row.getCell(3));
			if (row.getCell(3) == null || row.getCell(3).equals("")) {
				System.out.println("This is our row");
				break;
			}
			startingRowNum++;
		}

		for (int i = startingRowNum; i < startingRowNum + 4; i++) {
			Row row = sheet.getRow(i);
			Cell cell = (Cell) row.getCell(2);
			question.add(cell.getStringCellValue());
			link.add(cell.getHyperlink().getAddress());
		}

		qinfo.setQuestion(question);
		qinfo.setLinks(link);
		for (String q : question) {
			System.out.println(q);
		}

		sendMessage(qinfo.getQuestion(), qinfo.getLinks());
		fileInputStream.close();
		// update sheet with Sent column as yes
		 // Update sheet with "Sent" column as "Yes"
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        for (int i = startingRowNum; i < startingRowNum + 4; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            Cell sentCell = row.getCell(3); // Assuming "Sent" column is the fourth column
            if (sentCell == null) {
                sentCell = row.createCell(3);
            }
            sentCell.setCellValue("Yes");
        }

        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();

		// for testing
		for (int i = startingRowNum; i < startingRowNum + 4; i++) {
			Row row = sheet.getRow(i);
			Cell cell = (Cell) row.getCell(3);
			System.out.println(cell.getStringCellValue());
		}

	}

	public static void sendMessage(ArrayList question, ArrayList link) {
		Twilio.init(ACCOUNT_SID, AUTH_TOKEN);
		String toNumber = "whatsapp:+917667434147"; // recipient's WhatsApp number
		String fromNumber = "whatsapp:+14155238886";
		StringBuilder msg = new StringBuilder();

		msg.append("Good Morning! Here are your daily questions. Let's solve these.\n");
		for (int k = 0; k < question.size(); k++) {
			msg.append((k + 1) + ". " + question.get(k) + "\n");
			msg.append("   Link: " + link.get(k) + "\n\n");

		}
		Message message = Message.creator(new PhoneNumber(toNumber), new PhoneNumber(fromNumber), msg.toString())
				.create();

		System.out.println(message.getSid());

	}

	// To-do
	/**
	 * Exception handling - if message sending failed Scheduling Updating the rows
	 * with sent.
	 */
}
