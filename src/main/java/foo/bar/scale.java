package foo.bar;

import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hslf.usermodel.HSLFAutoShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;

/*
 * 	reference: https://stackoverflow.com/a/45664920
 */

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class scale {
	public static void main(String[] args) throws IOException {
		File folder = new File("C:/TEMP/16lh9");
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
            convertPptx(filename);    
        } else {
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
	        	
	            if (sh instanceof XSLFTextShape) {
	                XSLFTextShape shape = (XSLFTextShape) sh;
	                
	                Rectangle2D rectangle = shape.getAnchor();
	                
	                double sourceSlideX = rectangle.getX();
	                double sourceSlideY = rectangle.getY();
	                double sourceSlideWidth = rectangle.getWidth();
	                double sourceSlideHeight = rectangle.getHeight();
	                
	                if (shape.getText().contains("0")) {
		                shape.setAnchor(new Rectangle(
		                		(new Double(0)).intValue(), 
		                		(new Double(sourceSlideY)).intValue(), 
		                		(new Double(targetPageWidth)).intValue(), 
		                		(new Double(sourceSlideHeight)).intValue()));	                
	                }
	                else {
	                	shape.setAnchor(new Rectangle(
		                		(new Double(sourceSlideX)).intValue(), 
		                		(new Double(sourceSlideY)).intValue(), 
		                		(new Double(targetPageWidth)).intValue(), 
		                		(new Double(sourceSlideHeight)).intValue()));
	                }
	                	 	                              
	                System.out.println(sourceSlideX);
	                System.out.println(sourceSlideY);
	                System.out.println(sourceSlideWidth);
	                System.out.println(sourceSlideHeight);	                
	                System.out.println(shape.getText());
	            } 
	        }
	    }
        
        FileOutputStream out = new FileOutputStream(new File(filename));
        ppt.write(out);
        
        out.close();		
    }
    
    public static void convertPpt(String filename) throws IOException {
		HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl(filename));
		
		double sourcePageWidth = ppt.getPageSize().getWidth();
		double sourcePageHeight = ppt.getPageSize().getHeight(); 
				
		double targetPageWidth = sourcePageWidth * 4 / 3;
		ppt.setPageSize(new Dimension((new Double(targetPageWidth)).intValue(), (new Double(sourcePageHeight)).intValue()));

		
		List<HSLFSlide> slides = ppt.getSlides();
        for (int i = 0; i < slides.size(); i++) {
        	
        	HSLFSlide slide = slides.get(i);
            List<HSLFShape> shapes = slide.getShapes();
            for (int j=0; j < shapes.size(); j++) {
            	HSLFShape sh = (HSLFShape) shapes.get(j);
            	
	        	if (sh instanceof HSLFAutoShape) {
	                HSLFAutoShape shape = (HSLFAutoShape) sh;
	                
	                Rectangle2D rectangle = shape.getAnchor();	                
	                double sourceSlideX = rectangle.getX();
	                double sourceSlideY = rectangle.getY();
	                double sourceSlideWidth = rectangle.getWidth();
	                double sourceSlideHeight = rectangle.getHeight();
	                
	                shape.setAnchor(new Rectangle(
	                		(new Double(sourceSlideX)).intValue(), 
	                		(new Double(sourceSlideY)).intValue(), 
	                		(new Double(targetPageWidth)).intValue(), 
	                		(new Double(sourceSlideHeight)).intValue()));
	                	 
	                System.out.println(sourceSlideX);
	                System.out.println(sourceSlideY);
	                System.out.println(sourceSlideWidth);
	                System.out.println(sourceSlideHeight);	                
	                System.out.println(shape.getText());
	                
	                shapes.set(j, shape);
	                	                
	            } else if (sh instanceof HSLFTextBox) {
	                HSLFTextBox textbox = (HSLFTextBox) sh;
	                
	                Rectangle2D rectangle = textbox.getAnchor();	                
	                double sourceSlideX = rectangle.getX();
	                double sourceSlideY = rectangle.getY();
	                double sourceSlideWidth = rectangle.getWidth();
	                double sourceSlideHeight = rectangle.getHeight();
	                
	                textbox.setAnchor(new Rectangle(
	                		(new Double(0)).intValue(), 
	                		(new Double(sourceSlideY)).intValue(), 
	                		(new Double(targetPageWidth)).intValue(), 
	                		(new Double(sourceSlideHeight)).intValue()));
	                	 
	                System.out.println(sourceSlideX);
	                System.out.println(sourceSlideY);
	                System.out.println(sourceSlideWidth);
	                System.out.println(sourceSlideHeight);	                
	                System.out.println(textbox.getText());
	                
	                shapes.set(j, textbox);
	            }
	        }
	    }
		
		FileOutputStream out = new FileOutputStream(new File(filename));
        ppt.write(out);
        
        out.close();		
	}
}