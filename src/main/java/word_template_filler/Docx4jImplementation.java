package word_template_filler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Map;

import org.docx4j.Docx4J;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Docx4jImplementation {

	public static InputStream fillTemplate(byte[] wordDocArray, Map<String, String> documentDataMap) throws Exception {

		InputStream wordDocStream = new ByteArrayInputStream(wordDocArray);
		WordprocessingMLPackage wordMLPackage = Docx4J.load(wordDocStream);
		wordDocStream.close();		

		VariablePrepare.prepare(wordMLPackage);
		wordMLPackage.getMainDocumentPart().variableReplace(documentDataMap);
		
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		Docx4J.save(wordMLPackage, outStream, Docx4J.FLAG_NONE);
		byte[] fileByteArray = outStream.toByteArray();
		outStream.close();
		//Check for possible memory leak from passing this inputstream
		return new ByteArrayInputStream(fileByteArray);
	}
}
