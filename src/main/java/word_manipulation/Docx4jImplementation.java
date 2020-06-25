package word_manipulation;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.docx4j.Docx4J;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Docx4jImplementation {

	public static InputStream fillTemplate(byte[] wordDocArray, Map<String, String> documentDataMap) throws Exception {
		WordprocessingMLPackage wordMLPackage = byteArrayToWPMLPackage(wordDocArray);		

		VariablePrepare.prepare(wordMLPackage);
		wordMLPackage.getMainDocumentPart().variableReplace(documentDataMap);
		
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		Docx4J.save(wordMLPackage, outStream, Docx4J.FLAG_NONE);
		
		return outputStreamToInputStream(outStream);
	}

	public static InputStream convertPDF(byte[] wordDocArray) throws Docx4JException, IOException {
		WordprocessingMLPackage wordMLPackage = byteArrayToWPMLPackage(wordDocArray);
		
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		Docx4J.toPDF(wordMLPackage, outStream);
		
		return outputStreamToInputStream(outStream);
	}
	
	private static WordprocessingMLPackage byteArrayToWPMLPackage(byte[] wordDocArray) throws IOException, Docx4JException {
		InputStream wordDocStream = new ByteArrayInputStream(wordDocArray);
		WordprocessingMLPackage wordMLPackage = Docx4J.load(wordDocStream);
		wordDocStream.close();
		return wordMLPackage;
	}
	
	private static InputStream outputStreamToInputStream(ByteArrayOutputStream outStream) throws IOException {
		byte[] fileByteArray = outStream.toByteArray();
		outStream.close();
		//Check for possible memory leak from passing this inputstream
		return new ByteArrayInputStream(fileByteArray);
	}
}
