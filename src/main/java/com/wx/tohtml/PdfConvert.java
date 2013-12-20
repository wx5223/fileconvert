package com.wx.tohtml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.Formatter;

import net.timendum.pdf.Images2HTML;
import net.timendum.pdf.PDFText2HTML;
import net.timendum.pdf.StatisticParser;

import org.apache.commons.io.FileUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.examples.html.ToHtml;

/** 
 * @Description 
 * @date 2013年12月6日
 * @author WangXin
 */
public class PdfConvert implements IHtmlConvert {
	
	public void convert(File srcFile, String desPath) {
		final String realName = srcFile.getName();
		final String htmlPath = desPath;
		PDDocument document = null;
		Writer output = null;
		try {
			document = PDDocument.load(srcFile);
			output = new OutputStreamWriter(new FileOutputStream(htmlPath + realName + ".html"), "UTF-8");
			StatisticParser statisticParser = new StatisticParser();
			statisticParser.writeText(document, new Writer() {
				@Override
				public void write(char[] cbuf, int off, int len) throws IOException {
				}
				@Override
				public void flush() throws IOException {
				}
				@Override
				public void close() throws IOException {
				}
			});
			Images2HTML image = null;
			image = new Images2HTML();
			File imgDir = new File(htmlPath + realName + File.separator);
			if(!imgDir.exists() && !imgDir.isDirectory()) imgDir.mkdirs();
			image.setBasePath(imgDir);
			image.setRelativePath(realName);
			image.processDocument(document);
			PDFText2HTML stripper = new PDFText2HTML("UTF-8", statisticParser);
			stripper.setForceParsing(true);
			stripper.setSortByPosition(true);
			stripper.setImageStripper(image);
			stripper.writeText(document, output);
		} catch (UnsupportedEncodingException e1) {
			e1.printStackTrace();
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		} finally {
			if (output != null) {
				try {
					output.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (document != null) {
				try {
					document.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}

