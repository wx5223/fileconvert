package com.wx.tohtml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Formatter;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.examples.html.ToHtml;

/** 
 * @Description 
 * @date 2013年12月6日
 * @author WangXin
 */
public class ExcelConvert implements IHtmlConvert {
	/*
	 * xls & xlsx 
	 * but now just convert the first sheet
	 * @see com.wx.tohtml.IHtmlConvert#convert(java.io.File, java.lang.String)
	 */
	public void convert(File srcFile, String desPath) {
		final String realName = srcFile.getName();
		final String htmlPath = desPath;
		InputStream is = null;
		try {
			is = new FileInputStream(srcFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		StringBuffer sb = new StringBuffer();
		Formatter out = null;
		try {
			ToHtml tohtml = ToHtml.create(is, sb);
			tohtml.setCompleteHTML(true);
			tohtml.printPage();
			FileUtils.writeStringToFile(new File(htmlPath + realName + ".html"), sb.toString(), "utf-8");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
            if (out != null)
                out.close();
        }
	}
}

