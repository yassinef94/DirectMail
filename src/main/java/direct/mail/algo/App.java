package direct.mail.algo;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

import org.apache.log4j.Logger;

public class App {

	private static final Logger logger = Logger.getLogger(App.class.getName());

	public static final String pathFileDocx = "D:\\template.docx";
	public static final String outputFile = "D:\\";
	public static final String fileName = "test.docx";

	public static void main(String[] args) {
		try {
			if (checkExistsInuputFileDocx()) {
				new DirectMailStatic().directMailStatic(pathFileDocx, outputFile, fileName, initData());
				System.out.println("End of treatment with success");
			} else 
				System.out.println("File not exists in System");
		} catch (Exception e) {
			logger.error("Error in Class App in method main :: " + e.toString());
		}
	}

	public static Boolean checkExistsInuputFileDocx() {
		if (new File(pathFileDocx).exists())
			return true;
		return false;
	}

	private static HashMap<String, String> initData() {
		HashMap<String, String> map = new HashMap<String, String>();
		map.put("&name&", "Yassine");
		map.put("&date&",  new SimpleDateFormat("dd/MM/YYYY").format(new Date()));
		return map;
	}
}
