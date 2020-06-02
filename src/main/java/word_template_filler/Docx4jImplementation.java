package word_template_filler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import javax.xml.bind.JAXBException;

import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.model.datastorage.BindingHandler;
import org.docx4j.model.datastorage.CustomXmlDataStorage;
import org.docx4j.model.datastorage.CustomXmlDataStorageImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePart;
import org.docx4j.openpackaging.parts.CustomXmlPart;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Docx4jImplementation {

	public static InputStream fillTemplate(byte[] wordDocArray) throws Docx4JException, IOException, JAXBException {
		
		//Dynamically populate
		Map<String, String> documentDataMap = new HashMap<String, String>();
		documentDataMap.put("name", "Tony Pastrami");
		InputStream wordDocStream = new ByteArrayInputStream(wordDocArray);
		WordprocessingMLPackage wordMLPackage = Docx4J.load(wordDocStream);
		wordDocStream.close();		

		for (CustomXmlPart part : wordMLPackage.getCustomXmlDataStorageParts().values()) {
			CustomXmlDataStoragePart customPart = (CustomXmlDataStoragePart)part;
			CustomXmlDataStorage customXmlDataStorage = customPart.getData();
			//Change from data so something more specific and move to variable
			if (customXmlDataStorage.xpathGetNodes("/data", "") != null) {
				for (Node dataNode : customXmlDataStorage.xpathGetNodes("/data", "")) {
					NodeList nList = dataNode.getChildNodes();
					for (int i = 0; i < nList.getLength(); i++) {
					    Node attributeNode = nList.item(i);
						String nodeXPath = "/" + dataNode.getNodeName() + "/" + attributeNode.getNodeName().replace("#", "");
						((CustomXmlDataStorageImpl)customXmlDataStorage).setNodeValueAtXPath(nodeXPath, documentDataMap.get(attributeNode.getNodeName()), "");
					}
					System.out.println("XML String: " + customXmlDataStorage.getXML());
				}
			}
		}

		BindingHandler bh = new BindingHandler(wordMLPackage);
		bh.applyBindings(wordMLPackage.getMainDocumentPart());
		System.out.println(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));

		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		Docx4J.save(wordMLPackage, outStream, Docx4J.FLAG_NONE);
		byte[] fileByteArray = outStream.toByteArray();
		outStream.close();
		//Check for possible memory leak from passing this inputstream
		return new ByteArrayInputStream(fileByteArray);
	}
}
