package direct.mail.algo;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;

import org.apache.log4j.Logger;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Text;

public class DirectMailStatic {

	private static final Logger logger = Logger.getLogger(DirectMailStatic.class.getName());

	public File directMailStatic(String inputPathFile, String outPutPathFile, String fileName,
			HashMap<String, String> listData) {
		try {
			WordprocessingMLPackage document = importDocx(inputPathFile);

			document = treatmentWords(document, listData);
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			document.save(baos);

			File file = new File(outPutPathFile + File.separator + fileName);
			file.createNewFile();

			FileOutputStream fos = new FileOutputStream(file);
			baos.writeTo(fos);
		} catch (Exception e) {
			logger.error("Error in Class DirectMailStatic in method directMailStatic :: " + e.toString());
		}
		return null;
	}

	protected WordprocessingMLPackage treatmentWords(WordprocessingMLPackage document,
			Map<String, String> coordoonnee) {
		List<Object> texts = getAllElementFromObject(document.getMainDocumentPart(), Text.class);
		for (Object text : texts) {
			Text t = (Text) text;
			if (t.getValue().startsWith("&") || t.getValue().endsWith("&")) {
				if (countOccurrences(t.getValue(), '&') == 2) {
					System.out.println("Read key word :: "+t.getValue());
					for (Map.Entry<String, String> coordonnee : coordoonnee.entrySet()) {
						if (t.getValue().contains(coordonnee.getKey()) || t.getValue().equals(coordonnee.getKey())) {
							if (coordonnee.getValue() != null)
								t.setValue(t.getValue().replace(coordonnee.getKey(), coordonnee.getValue()));
							else
								t.setValue(t.getValue().replace(coordonnee.getKey(), ""));
							break;
						}
					}
				}
			}
		}
		return document;
	}

	private int countOccurrences(String text, char key) {
		int count = 0;
		for (int index = 0; index < text.length(); index++)
			if (text.charAt(index) == key)
				count++;
		return count;
	}

	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
		List<Object> result = new ArrayList<Object>();
		if (obj instanceof JAXBElement)
			obj = ((JAXBElement<?>) obj).getValue();

		if (obj.getClass().equals(toSearch))
			result.add(obj);
		else if (obj instanceof ContentAccessor) {
			List<?> children = ((ContentAccessor) obj).getContent();
			for (Object child : children) {
				result.addAll(getAllElementFromObject(child, toSearch));
			}
		}
		return result;
	}

	private WordprocessingMLPackage importDocx(String filePath) throws Exception {
		File file = new File(filePath);
		return WordprocessingMLPackage.load(file);
	}
}
