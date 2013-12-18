package com.wx.tohtml;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Formatter;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.RichTextRun;
import org.apache.poi.hslf.usermodel.SlideShow;

/** 
 * @Description 
 * @date 2013年12月7日
 * @author WangXin
 */
public class PptConvert implements IHtmlConvert {
	private String htmlPath = "";
	private String realName = "";
	private Appendable output;
	private Formatter out;
	
	public void convert(File srcFile, String desPath) {
		realName = srcFile.getName();
		htmlPath = desPath;
		SlideShow slideShow = null;
		try {
			slideShow = new SlideShow(new FileInputStream(srcFile));
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
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
	
	private void print(SlideShow slideShow) {
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

	private void getPics(SlideShow slideShow) {
		Dimension pgsize = slideShow.getPageSize();
		Slide[] slides = slideShow.getSlides();
		for (int i = 0; i < slides.length; i++) {
			TextRun[] truns = slides[i].getTextRuns();
			for (int k = 0; k < truns.length; k++) {
				RichTextRun[] rtruns = truns[k].getRichTextRuns();
				for (int j = 0; j < rtruns.length; j++) {
					rtruns[j].setFontIndex(1);
					rtruns[j].setFontName("宋体");
				}
			}
			BufferedImage img = new BufferedImage(pgsize.width, pgsize.height,
					BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = img.createGraphics();
			graphics.setPaint(Color.WHITE);
			graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width,
					pgsize.height));
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

