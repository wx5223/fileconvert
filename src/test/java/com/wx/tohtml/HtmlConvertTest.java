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

}

