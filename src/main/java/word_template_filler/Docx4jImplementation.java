package word_template_filler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.bind.JAXBException;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Docx4jImplementation {

	public static InputStream fillTemplate(byte[] wordDocArray) throws Docx4JException, IOException, JAXBException {
		
		InputStream wordDocStream = new ByteArrayInputStream(wordDocArray);
		WordprocessingMLPackage wordMLPackage = Docx4J.load(wordDocStream);
		wordDocStream.close();		
				
		String xml = "<data><name>Tony</name></data>";
        Docx4J.bind(wordMLPackage, xml, Docx4J.FLAG_NONE);

		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		Docx4J.save(wordMLPackage, outStream, Docx4J.FLAG_NONE);
		byte[] fileByteArray = outStream.toByteArray();
		outStream.close();
		//Check for possible memory leak from passing this inputstream
		return new ByteArrayInputStream(fileByteArray);
	}
}
