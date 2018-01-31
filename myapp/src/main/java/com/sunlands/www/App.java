package com.sunlands.www;

/**
 * Hello world!
 *
 */ 
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.util.List;

import javax.imageio.ImageIO;

//import org.apache.poi.hslf.HSLFSlideShow;
//import org.apache.poi.hslf.model.Slide;
//import org.apache.poi.hslf.model.TextRun;
//import org.apache.poi.hslf.usermodel.RichTextRun;
//import org.apache.poi.hslf.usermodel.SlideShow;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
 
public class App 
{
    public static void main( String[] args )
    {
		try {
			File raw_ppt_f = new File("1.pptx");
			Long filelength = raw_ppt_f.length();
			byte[] filecontent = new byte[filelength.intValue()];
			
			FileInputStream in = new FileInputStream(raw_ppt_f);  
            in.read(filecontent);  
            in.close();
			
			//ByteArrayInputStream pptInput = Utility.readNetFile(url);
			ByteArrayInputStream pptInput = new ByteArrayInputStream(filecontent);
						
			ByteArrayOutputStream result = convert(".pptx", pptInput);
			
			File img_f = new File("1.png");
			FileOutputStream imgOutput = new FileOutputStream(img_f);
			result.writeTo(imgOutput);
			imgOutput.close();
			
			//System.out.println(result.getTotalImage());
			//System.out.println(result.getImageByteList().size());
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		System.out.println( "Hello World!" );
    }
	
	public static ByteArrayOutputStream convert(String suffix, ByteArrayInputStream pptInput) {
		if (pptInput != null) {
			if (suffix.endsWith(".pptx")) {
				return ppt2007Img(pptInput);
			}
			/*
			else if (suffix.endsWith(".ppt")) {
				//return APP.ppt2003Img(pptInput);
				;
			}
			*/
		}
		
		return null;
	}
	
	private static void setFont(XSLFSlide slide) {
		for (XSLFShape shape : slide.getShapes()) {
			if (shape instanceof XSLFTextShape) {
				XSLFTextShape txtshape = (XSLFTextShape) shape;
				
				for (XSLFTextParagraph paragraph : txtshape.getTextParagraphs()) {
					List<XSLFTextRun> truns = paragraph.getTextRuns();
					
					for (XSLFTextRun trun : truns) {
						trun.setFontFamily("宋体");
						double currentFontSize = trun.getFontSize();
						
						if((currentFontSize <= 0)||(currentFontSize >= 26040)){
							trun.setFontSize((double)30);
						}
					}
				}
			}

		}
	}
	
	public static ByteArrayOutputStream ppt2007Img(ByteArrayInputStream bais) {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();		
		XMLSlideShow ppt = null;
		
		try {
			ppt = new XMLSlideShow(bais);
			Dimension pgsize = ppt.getPageSize();
			List<XSLFSlide> slides = ppt.getSlides();			
			
			for (int i = 0; i < slides.size(); i++) {				
				CTSlide xmlObject = slides.get(i).getXmlObject();
				setFont(slides.get(i));
				
				CTGroupShape spTree = xmlObject.getCSld().getSpTree();
				CTShape[] spArray = spTree.getSpArray();
				
				for (CTShape shape : spArray) {
					CTTextBody txBody = shape.getTxBody();

					if (txBody == null) {
						continue;
					}
					
					CTTextParagraph[] pArray = txBody.getPArray();
					//CTTextFont font = CTTextFont.Factory.parse("");
								
					for (CTTextParagraph textParagraph : pArray) {
						CTRegularTextRun[] textRuns = textParagraph.getRArray();
						
						for (CTRegularTextRun textRun : textRuns) {
							CTTextCharacterProperties properties = textRun.getRPr();
							//properties.setLatin(font);
						}
					}
				}
				
				BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();
				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
				slides.get(i).draw(graphics);
				
				ImageIO.write(img, "png", baos);
				baos.flush();
				baos.close();				
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (bais != null) {			
				try {
					bais.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	
		return baos;
	}
	
	/*
	private static void set2003Font(Slide slide) {
		TextRun[] truns = slide.getTextRuns();
		if(Utility.isNotEmpty(truns)){
			for (TextRun trun: truns) {
				RichTextRun[] rtruns = trun.getRichTextRuns();
				if(Utility.isNotEmpty(rtruns)){
					for (RichTextRun rtrun : rtruns) {
						rtrun.setFontIndex(1);
						rtrun.setFontName("宋体");
					}
				}
			}
		}
	}
	*/

	/*
	public static PptImageResult ppt2003Img(ByteArrayInputStream bais) {
		PptImageResult result = new PptImageResult();
		
		try {
			SlideShow ppt = new SlideShow(new HSLFSlideShow(bais));
			Dimension pgsize = ppt.getPageSize();
			Slide[] slides = ppt.getSlides();
			result.setTotalImage(slides.length);
			
			for (int i = 0; i < slides.length; i++) {
				set2003Font(slides[i]);
				BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();
				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
				slides[i].draw(graphics);
				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				ImageIO.write(img, "png", baos);
				baos.flush();
				baos.close();
				result.getImageByteList().add(baos.toByteArray());
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (bais != null) {
				try {
					bais.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
		return result;
	}
	*/
}
