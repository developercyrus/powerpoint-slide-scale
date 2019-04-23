package foo.bar;

import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.common.usermodel.fonts.FontGroup;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;

/*
 * 	reference: https://stackoverflow.com/a/45664920
 */

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class scale {
	static String path = "C:/TEMP/h";
	static String path4l3h = "4l3h";
	static String path16l9h = "16l9h";
	static double size54 = 54;
	static double size64 = 64;
	static String fontTypeMingLiU = "細明體";
	static String fontTypeDFKaiSB = "標楷體";
	
	
	public static void main(String[] args) throws IOException {
		File folder = new File(path + "/" + path4l3h);
		File[] listOfFiles = folder.listFiles();

		for (int i = 0; i < listOfFiles.length; i++) {
			if (listOfFiles[i].isFile()) {
				System.out.println(listOfFiles[i].getAbsolutePath());
				detect(listOfFiles[i].getAbsolutePath());
			} 
		}
	}
	
    public static void detect(String filename) throws IOException {
    	SlideShow<?,?> ppt = SlideShowFactory.create(new FileInputStream(filename));
        Slide<?,?> slide = ppt.getSlides().get(0);

        if (slide instanceof XSLFSlide) {
        	System.out.println("pptx");
            convertPptx(filename);    
        } else {
        	System.out.println("ppt");
        	convertPpt(filename);
        }

    }
    
    
    public static void convertPptx(String filename) throws IOException {
    	XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(filename));
    	
    	double sourcePageWidth = ppt.getPageSize().getWidth();
		double sourcePageHeight = ppt.getPageSize().getHeight(); 				
		
		double targetPageWidth = sourcePageWidth * 4 / 3;
		ppt.setPageSize(new Dimension((new Double(targetPageWidth)).intValue(), (new Double(sourcePageHeight)).intValue()));

		List<XSLFSlide> slides = ppt.getSlides();
        for (int i = 0; i < slides.size(); i++) {
        	
        	XSLFSlide slide = slides.get(i);

            List<XSLFShape> shapes = slide.getShapes();
            for (int j=0; j < shapes.size(); j++) {
            	XSLFShape sh = (XSLFShape) shapes.get(j);
	        	
            	
            	if (sh instanceof XSLFPictureShape) {
        			
        			XSLFPictureShape shape = (XSLFPictureShape) sh;
        			        			
        			Rectangle2D rectangle = shape.getAnchor();	         
        			
                    double sourceSlideX = rectangle.getX();
                    double sourceSlideY = rectangle.getY();
                    double sourceSlideWidth = rectangle.getWidth();
                    double sourceSlideHeight = rectangle.getHeight();
                            			
		            System.out.println("before XSLFPictureShape shape X=" + sourceSlideX);
		            System.out.println("before XSLFPictureShape shape Y=" + sourceSlideY);
		            System.out.println("before XSLFPictureShape shape Width=" + sourceSlideWidth);
		            System.out.println("before XSLFPictureShape shape Height=" + sourceSlideHeight);	  
                    
                    shape.setAnchor(new Rectangle(
                    		(new Double(0)).intValue(), 
                    		(new Double(sourceSlideY)).intValue(), 
                    		(new Double(targetPageWidth)).intValue(), 
                    		(new Double(sourceSlideHeight)).intValue()));
                    
		            System.out.println("after XSLFPictureShape shape X=" + shape.getAnchor().getX());
		            System.out.println("after XSLFPictureShape shape Y=" + shape.getAnchor().getY());
		            System.out.println("after XSLFPictureShape shape Width=" + shape.getAnchor().getWidth());
		            System.out.println("after XSLFPictureShape shape Height=" + shape.getAnchor().getHeight());	   

        		}
        		else if (sh instanceof XSLFTextShape) {
	                XSLFTextShape shape = (XSLFTextShape) sh;
	                
	                Rectangle2D rectangle = shape.getAnchor();
	                
	                double sourceSlideX = rectangle.getX();
	                double sourceSlideY = rectangle.getY();
	                double sourceSlideWidth = rectangle.getWidth();
	                double sourceSlideHeight = rectangle.getHeight();
	                            	 	
		            System.out.println("before XSLFTextShape shape X=" + sourceSlideX);
		            System.out.println("before XSLFTextShape shape Y=" + sourceSlideY);
		            System.out.println("before XSLFTextShape shape Width=" + sourceSlideWidth);
		            System.out.println("before XSLFTextShape shape Height=" + sourceSlideHeight);	                
	                

	                shape.setAnchor(new Rectangle(
	                		(new Double(0)).intValue(), 
	                		(new Double(sourceSlideY)).intValue(), 
	                		(new Double(targetPageWidth)).intValue(), 
	                		(new Double(sourceSlideHeight)).intValue()));	                

		            System.out.println("after XSLFTextShape shape X=" + shape.getAnchor().getX());
		            System.out.println("after XSLFTextShape shape Y=" + shape.getAnchor().getY());
		            System.out.println("after XSLFTextShape shape Width=" + shape.getAnchor().getWidth());
		            System.out.println("after XSLFTextShape shape Height=" + shape.getAnchor().getHeight());	                
		            
	                         	                    	
                	List<XSLFTextParagraph> paragraphs = shape.getTextParagraphs();
                	for (int k=0; k<paragraphs.size(); k++) {
                		XSLFTextParagraph paragraph = paragraphs.get(k);     
                		
                		List<XSLFTextRun> textruns = paragraph.getTextRuns();
                		for (int l=0; l<textruns.size(); l++) {
                			XSLFTextRun textrun = textruns.get(l); 
                			
                			if (textrun.getFontSize() == size54) {            
                			
	                			System.out.println("before XSLFTextShape fontsize=" + textrun.getFontSize());
	                			System.out.println("before XSLFTextShape fonttype=" + textrun.getFontFamily());

	                			
	                			textrun.setFontSize(size64);                      			                    			
	                			textrun.setFontFamily(fontTypeDFKaiSB, FontGroup.EAST_ASIAN);
	                			
	                			System.out.println("after XSLFTextShape fontsize=" + textrun.getFontSize());
	                			System.out.println("after XSLFTextShape fonttype=" + textrun.getFontFamily());
	                			System.out.println(textrun.getRawText());
	                			
	                			textruns.set(l, textrun);
                			}
                		}
                		paragraphs.set(k, paragraph);
                	}                    	
                    
	            } 
	        }
            
            
	    }
        
        File oldFile = new File(filename);
		FileOutputStream out = new FileOutputStream(new File(path + "/" + path16l9h + "/" + oldFile.getName()));
        ppt.write(out);
        
        out.close();		
    }
    
    
    
    public static void convertPpt(String filename) throws IOException {
		HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(filename));
		
		double sourcePageWidth = ppt.getPageSize().getWidth();
		double sourcePageHeight = ppt.getPageSize().getHeight(); 
				
		System.out.println(ppt.getPageSize());
		
		double targetPageWidth = sourcePageWidth * 4 / 3;
		ppt.setPageSize(new Dimension((new Double(targetPageWidth)).intValue(), (new Double(sourcePageHeight)).intValue()));
		
		List<HSLFSlide> slides = ppt.getSlides();
        for (int i=0; i<slides.size(); i++) {
        	
        	System.out.println("slide:" + i);
        	
        	HSLFSlide slide = slides.get(i);        	       
        	
            List<HSLFShape> shapes = slide.getShapes();
            for (int j=0; j<shapes.size(); j++) {
            	HSLFShape sh = (HSLFShape) shapes.get(j);

            	System.out.println("shape:" + j);            
            		
        		if (sh instanceof HSLFPictureShape) {
        			
        			HSLFPictureShape shape = (HSLFPictureShape) sh;
        			        			
        			Rectangle2D rectangle = shape.getAnchor();	                
                    double sourceSlideX = rectangle.getX();
                    double sourceSlideY = rectangle.getY();
                    double sourceSlideWidth = rectangle.getWidth();
                    double sourceSlideHeight = rectangle.getHeight();
                    
                    shape.setAnchor(new Rectangle(
                    		(new Double(0)).intValue(), 
                    		(new Double(sourceSlideY)).intValue(), 
                    		(new Double(targetPageWidth)).intValue(), 
                    		(new Double(sourceSlideHeight)).intValue()));

                    System.out.println(sourceSlideX);
                    System.out.println(sourceSlideY);
                    System.out.println(sourceSlideWidth);
                    System.out.println(sourceSlideHeight);	                

                    shapes.set(j, shape);
        			
        		}
        		else if (sh instanceof HSLFTextShape) {
        	
        			HSLFTextShape shape = (HSLFTextShape) sh;
        			
        			
        			Rectangle2D rectangle = shape.getAnchor();	                
                    double sourceSlideX = rectangle.getX();
                    double sourceSlideY = rectangle.getY();
                    double sourceSlideWidth = rectangle.getWidth();
                    double sourceSlideHeight = rectangle.getHeight();
                    
                    shape.setAnchor(new Rectangle(
                    		(new Double(0)).intValue(), 
                    		(new Double(sourceSlideY)).intValue(), 
                    		(new Double(targetPageWidth)).intValue(), 
                    		(new Double(sourceSlideHeight)).intValue()));
                    	 
                    //shape.setHorizontalCentered(Boolean.TRUE);
                    
                    System.out.println(sourceSlideX);
                    System.out.println(sourceSlideY);
                    System.out.println(sourceSlideWidth);
                    System.out.println(sourceSlideHeight);	                
                    System.out.println(shape.getText());
                    
                    
                    	                    	
                	List<HSLFTextParagraph> paragraphs = shape.getTextParagraphs();
                	for (int k=0; k<paragraphs.size(); k++) {
                		HSLFTextParagraph paragraph = paragraphs.get(k);     
                		
                		List<HSLFTextRun> textruns = paragraph.getTextRuns();
                		for (int l=0; l<textruns.size(); l++) {
                			HSLFTextRun textrun = textruns.get(l); 

                			if (textrun.getFontSize() == size54) {                                                			
	                			System.out.println("HSLFTextShape fontsize befor=" + textrun.getFontSize());
	                			System.out.println("HSLFTextShape fonttype before=" + textrun.getFontFamily());
	                			
	                			textrun.setFontSize(size64);    
	                			textrun.setFontFamily(fontTypeDFKaiSB, FontGroup.EAST_ASIAN);
	                			
	                			System.out.println("HSLFTextShape fontsize after=" + textrun.getFontSize());
	                			System.out.println("HSLFTextShape fonttype after=" + textrun.getFontFamily());
	                			
	                			textruns.set(l, textrun);
                			}
                		}                		                    		
                		paragraphs.set(k, paragraph);
                	}                    	
                    shapes.set(j, shape);
        		}
        		
            }
        			
	    }
		
        File oldFile = new File(filename);
        FileOutputStream out = new FileOutputStream(new File(path + "/" + path16l9h + "/" + oldFile.getName()));
        ppt.write(out);
        
        out.close();		
	}
    
    public static boolean containsHan(String s) {
        for (int i = 0; i < s.length(); ) {
            int codepoint = s.codePointAt(i);
            i += Character.charCount(codepoint);
            if (Character.UnicodeScript.of(codepoint) == Character.UnicodeScript.HAN) {
                return true;
            }
        }
        return false;
    }
}