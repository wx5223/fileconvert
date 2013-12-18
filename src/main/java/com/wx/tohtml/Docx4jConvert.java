package com.wx.tohtml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/** 
 * @Description 
 * @date 2013年12月5日
 * @author WangXin
 */
public class Docx4jConvert implements IHtmlConvert {
	public void convert(File srcFile, String desPath) {
		final String realName = srcFile.getName();
		final String htmlPath = desPath;
		WordprocessingMLPackage wordMLPackage = null;
		try {
			wordMLPackage = Docx4J.load(new FileInputStream(srcFile));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		}
		HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
		htmlSettings.setImageDirPath(htmlPath + realName);
		htmlSettings.setImageTargetUri(realName);
		htmlSettings.setWmlPackage(wordMLPackage);
		try {
			OutputStream os = new FileOutputStream(htmlPath + realName + ".html");
			Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML",	true);
			Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		}
	}
}

