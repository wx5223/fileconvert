package com.wx.tohtml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.core.XWPFConverterException;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/** 
 * @Description 
 * @date 2013年12月5日
 * @author WangXin
 */
public class DocxConvert implements IHtmlConvert {
	public void convert(File srcFile, String desPath) {
		final String realName = srcFile.getName();
		final String htmlPath = desPath;
		XWPFDocument document = null;
		try {
			InputStream in = new FileInputStream(srcFile);
			document = new XWPFDocument(in);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(new File(htmlPath + realName)));
		options.setExtractor(new FileImageExtractor(new File(htmlPath + realName)));
		try {
			OutputStream out = new FileOutputStream(htmlPath + realName + ".html");
			XHTMLConverter.getInstance().convert(document, out, options);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (XWPFConverterException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
}

