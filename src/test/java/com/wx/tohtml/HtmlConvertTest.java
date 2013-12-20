package com.wx.tohtml;

import static org.junit.Assert.*;

import java.io.File;

import org.junit.Test;

/** 
 * @Description 
 * @date 2013年12月16日
 * @author WangXin
 */
public class HtmlConvertTest {

	@Test
	public void doc() {
		File f = new File("src/test/resources/test1.doc");
		IHtmlConvert c = new DocConvert();
		c.convert(f, "src/test/resources/html/");
	}
	
	@Test
	public void docx4j() {
		File f = new File("src/test/resources/test1.docx");
		IHtmlConvert c = new Docx4jConvert();
		File dic = new File("src/test/resources/html/");
		if(!dic.isDirectory()) {
			dic.mkdirs();
		}
		c.convert(f, "src/test/resources/html/");
	}
	@Test
	public void docx() {
		File f = new File("src/test/resources/test2.docx");
		IHtmlConvert c = new DocxConvert();
		File dic = new File("src/test/resources/html/");
		if(!dic.isDirectory()) {
			dic.mkdirs();
		}
		c.convert(f, "src/test/resources/html/");
	}
	@Test
	public void excel() {
		File f = new File("src/test/resources/test3.xlsx");
		IHtmlConvert c = new ExcelConvert();
		File dic = new File("src/test/resources/html/");
		if(!dic.isDirectory()) {
			dic.mkdirs();
		}
		c.convert(f, "src/test/resources/html/");
	}
	@Test
	public void ppt() {
		File f = new File("src/test/resources/test4.ppt");
		IHtmlConvert c = new PptConvert();
		File dic = new File("src/test/resources/html/");
		if(!dic.isDirectory()) {
			dic.mkdirs();
		}
		c.convert(f, "src/test/resources/html/");
	}
	@Test
	public void pptx() {
		File f = new File("src/test/resources/test5.pptx");
		IHtmlConvert c = new PptxConvert();
		File dic = new File("src/test/resources/html/");
		if(!dic.isDirectory()) {
			dic.mkdirs();
		}
		c.convert(f, "src/test/resources/html/");
	}
	@Test
	public void pdf() {
		File f = new File("src/test/resources/test6.pdf");
		IHtmlConvert c = new PdfConvert();
		File dic = new File("src/test/resources/html/");
		if(!dic.isDirectory()) {
			dic.mkdirs();
		}
		c.convert(f, "src/test/resources/html/");
	}

}

