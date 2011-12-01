package com.criticalmass.powerpoint_poc;

import java.io.FileInputStream;

import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextShape;
import org.apache.poi.hslf.usermodel.SlideShow;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception 
    {
//        if(args.length == 0)  {
//            System.out.println("Input file is required");
//            return;
//        }
//
    	SlideShow ppt = new SlideShow(new FileInputStream("/Users/mminella/CM/STS_Workspace/powerpoint_poc/src/main/resources/ppt_presentation_simple_blue.ppt"));
    	
    	for(Slide slide : ppt.getSlides()) {
    		System.out.println("Title: " + slide.getTitle());
    		
	          for(Shape shape : slide.getShapes()){
	          if(shape instanceof TextShape) {
	              TextShape tsh = (TextShape)shape;
                  System.out.println(tsh.getText());
//	                      System.out.println("  bold: " + tsh.);
//	                      System.out.println("  italic: " + tsh.isItalic());
//	                      System.out.println("  underline: " + tsh.isUnderline());
//	                      System.out.println("  font.family: " + tsh.getFontFamily());
//	                      System.out.println("  font.size: " + tsh.getFontSize());
//	                      System.out.println("  font.color: " + tsh.get);
	          }
	      }
    	}
    	
    	
    	
//        for( slide : ppt.getSlides()){
//            System.out.println("Title: " + slide.getTitle());
//
//            for(XSLFShape shape : slide.getShapes()){
//                if(shape instanceof XSLFTextShape) {
//                    XSLFTextShape tsh = (XSLFTextShape)shape;
//                    for(XSLFTextParagraph p : tsh.getTextParagraphs()){
//                        System.out.println("Paragraph level: " + p.getLevel());
//                        for(XSLFTextRun r : p){
//                            System.out.println(r.getText());
//                            System.out.println("  bold: " + r.isBold());
//                            System.out.println("  italic: " + r.isItalic());
//                            System.out.println("  underline: " + r.isUnderline());
//                            System.out.println("  font.family: " + r.getFontFamily());
//                            System.out.println("  font.size: " + r.getFontSize());
//                        }
//                    }
//                }
//            }
//        }
    }
}
