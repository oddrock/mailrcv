package com.oddrock.mail.mailrcvr;

import com.oddrock.common.mail.pop3.Pop3Config;
import com.oddrock.common.mail.pop3.Pop3MailRcvr;
import com.oddrock.common.prop.Prop;

public class MailRcvr {

	public static void main(String[] args) throws Exception {
		String server = Prop.get("mail.popserver");
		String account = Prop.get("mail.account");
		String passwd = Prop.get("mail.passwd");
		String localAttachDirPath = Prop.get("mail.savefolder");
		String rejectAddresses = Prop.get("mail.rejectAddresses");
		Pop3Config pop3Config = new Pop3Config(server, account, passwd, localAttachDirPath, rejectAddresses);
		new Pop3MailRcvr().rcvMail(pop3Config);
	}

}
