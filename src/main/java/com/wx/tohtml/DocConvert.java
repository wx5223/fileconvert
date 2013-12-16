package com.wx.tohtml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.w3c.dom.Document;

/** 
 * @Description 
 * @date 2013年12月5日
 * @author WangXin
 */
public class DocConvert implements IHtmlConvert {
	public void convert(File srcFile, String desPath) {
		final String realName = srcFile.getName();
		final String htmlPath = desPath;
		File imageDir = new File(htmlPath + realName);  
        final String suggestDirName = imageDir.getName();  
        if (!imageDir.isDirectory()) {  
            imageDir.mkdirs();
        }  
		InputStream input = null;
		try {
			input = new FileInputStream(srcFile);
		} catch (FileNotFoundException e2) {
			e2.printStackTrace();
		}
		HWPFDocument wordDocument = null;
		try {
			wordDocument = new HWPFDocument(input);
		} catch (IOException e) {
			e.printStackTrace();
		}
		Document doc = null;
		try {
			doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		}
		WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(doc);
		wordToHtmlConverter.setPicturesManager(new PicturesManager() {
			public String savePicture(byte[] content, PictureType pictureType, String suggestedName, 
					float widthInches, float heightInches) {
				String imgagePath = htmlPath + realName + File.separator + suggestedName;  
                File file = new File(imgagePath);  
                FileOutputStream fos = null;  
                try {  
                    fos = new FileOutputStream(file);  
                    fos.write(content);  
                    fos.close();  
                } catch (Exception e) {  
                    e.printStackTrace();  
                }  
                return suggestDirName + File.separator + suggestedName;  
			}
		});
		wordToHtmlConverter.processDocument(wordDocument);
		Document htmlDocument = wordToHtmlConverter.getDocument();
		DOMSource domSource = new DOMSource(htmlDocument);
		try {
			FileOutputStream fos = new FileOutputStream(htmlPath + realName + ".html");
			StreamResult streamResult = new StreamResult(fos);
			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.setOutputProperty(OutputKeys.MEDIA_TYPE, "yes" );
			serializer.transform(domSource, streamResult);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (TransformerConfigurationException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (TransformerFactoryConfigurationError e) {
			e.printStackTrace();
		} catch (TransformerException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

