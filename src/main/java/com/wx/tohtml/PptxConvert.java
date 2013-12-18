package com.wx.tohtml;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.Closeable;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Formatter;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.util.PPTX2PNG;

/** 
 * @Description 
 * @date 2013年12月7日
 * @author WangXin
 */
public class PptxConvert implements IHtmlConvert {
	private String htmlPath = "";
	private String realName = "";
	private Appendable output;
	private Formatter out;
	
	public void convert(File srcFile, String desPath) {
		realName = srcFile.getName();
		htmlPath = desPath;
		XMLSlideShow slideShow = null;
		try {
			slideShow = new XMLSlideShow(OPCPackage.open(srcFile));
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
		File file = new File(htmlPath + realName + ".html");
		file.getParentFile().mkdirs();
		try {
			output = new PrintWriter(new FileWriter(file));
		} catch (IOException e) {
			e.printStackTrace();
		}
		print(slideShow);
	}
	
	private void print(XMLSlideShow slideShow) {
		out = new Formatter(output);
		out.format("<?xml version=\"1.0\" encoding=\"utf-8\" ?>%n");
		out.format("<html>%n");
		out.format("<head>%n");
		out.format("</head>%n");
		out.format("<body>%n");
		out.format("<table>%n");
		getPics(slideShow);
		out.format("</table>");
		out.format("</body>%n");
		out.format("</html>%n");
		out.flush();
		out.close();
	}

	private void printImage(String scr) {
		out.format("<tr>%n");
		out.format("<td>%n");
		out.format("<img src='" + scr + "'/>");
		out.format("</td>%n");
		out.format("</tr>%n");
	}

	private void getPics(XMLSlideShow slideShow) {
		Dimension pgsize = slideShow.getPageSize();
		XSLFSlide[] slides = slideShow.getSlides();
		for (int i = 0; i < slides.length; i++) {
			XSLFTextShape[] xShape = slides[i].getPlaceholders();
            for(int j=0; j< xShape.length;j++){
              List<XSLFTextParagraph> xTexts = xShape[j].getTextParagraphs();
               for(int k=0;k<xTexts.size();k++){
                    List<XSLFTextRun> list = xTexts.get(k).getTextRuns();
                     for(XSLFTextRun xtr : list){
                          xtr.setFontFamily("宋体");
                    }
              }
            }
            BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB );
            Graphics2D graphics = img.createGraphics();
            graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON );
            graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY );
            graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC );
            graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON );
            graphics.setColor(Color. white);
            graphics.clearRect(0, 0, pgsize.width, pgsize.height);
            slides[i].draw(graphics);
			String imgSrcPath = htmlPath + realName + File.separator + i
					+ ".png";
			File ff = new File(imgSrcPath);
			ff.getParentFile().mkdirs();
			FileOutputStream fos = null;
			try {
				fos = new FileOutputStream(ff);
				ImageIO.write(img, "png", fos);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				try {
					if(fos != null)
						fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			printImage(realName + File.separator + i + ".png");
		}
	}
}

