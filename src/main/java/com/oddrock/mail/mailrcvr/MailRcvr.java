package com.oddrock.mail.mailrcvr;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.xmlbeans.XmlException;

import com.oddrock.common.file.FileUtils;
import com.oddrock.common.mail.MailRecv;
import com.oddrock.common.mail.MailRecvAttach;
import com.oddrock.common.mail.pop3.Pop3Config;
import com.oddrock.common.mail.pop3.Pop3MailRcvr;
import com.oddrock.common.office.excel.ExcelUtils;
import com.oddrock.common.office.ppt.PptUtils;
import com.oddrock.common.office.word.WordUtils;
import com.oddrock.common.pdf.PdfUtils;
import com.oddrock.common.prop.Prop;

public class MailRcvr {
	// 将邮件附件内容解析出来保存到同名txt文件
	private static void parseMailAttachFile2TxtFile(MailRecv mail) throws IOException, XmlException, OpenXML4JException{
		for(MailRecvAttach attach : mail.getAttachments()){
			parseFile2TxtFile(attach.getLocalFilePath());
		}
	}
	
	// 将文件内容解析出来保存到同名txt文件
	private static void parseFile2TxtFile(String filePath) throws IOException, XmlException, OpenXML4JException{
		if(!FileUtils.fileExists(filePath)){
			return;
		}
		File file = new File(filePath);
		String suffix = FileUtils.getFileNameSuffix(file);
		if(suffix.equalsIgnoreCase("doc") || suffix.equalsIgnoreCase("docx")){
			WordUtils.createTxtFileFromWord(file);
		}else if(suffix.equalsIgnoreCase("xls") || suffix.equalsIgnoreCase("xlsx")){
			ExcelUtils.createTxtFileFromExcel(file);
		}else if(suffix.equalsIgnoreCase("ppt") || suffix.equalsIgnoreCase("pptx")){
			PptUtils.createTxtFileFromPpt(file);
		}else if(suffix.equalsIgnoreCase("pdf")){
			PdfUtils.createTxtFileFromPdf(file);
		}
	}

	public static void main(String[] args) throws Exception {
		String server = Prop.get("mail.popserver");
		String account = Prop.get("mail.account");
		String passwd = Prop.get("mail.passwd");
		String localAttachDirPath = Prop.get("mail.savefolder");
		String rejectAddresses = Prop.get("mail.rejectAddresses");
		Pop3Config pop3Config = new Pop3Config(server, account, passwd, localAttachDirPath, rejectAddresses);
		List<MailRecv> mailList = new Pop3MailRcvr().rcvMail(pop3Config);
		for(MailRecv mail : mailList){
			System.out.println(mail.getSubject());
			parseMailAttachFile2TxtFile(mail);
		}
	}

}
