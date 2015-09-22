package utils;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

public class WordHandler {

		public void launchWord(){
			try {
			     if (Desktop.isDesktopSupported()) {
			       Desktop.getDesktop().open(new File("HelloWord.docx"));
			     //  Desktop.getDesktop().edit(new File("rapport.doc"));
			     }
			   } catch (IOException ioe) {
			     ioe.printStackTrace();
			  }
		}
		
		public void createWord(){
			try{
				WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
			wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "Hello Word!");
			wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Subtitle",
			"This is a subtitle!");
			try {wordMLPackage.save(new java.io.File("HelloWord.docx"));}
			catch(Docx4JException e){}
			}
			catch(InvalidFormatException e){}
		}
		
		public void modifyWord(){
			try{WordprocessingMLPackage wordMLPackage = 
				WordprocessingMLPackage.load(new java.io.File("HelloWord.docx"));
				 MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

				 documentPart.addStyledParagraphOfText("Title", "Hello Word2!");
				 documentPart.addStyledParagraphOfText("Subtitle",
			"This is a subtitle2!");
			wordMLPackage.save(new java.io.File("HelloWord.docx"));}
			catch(Docx4JException e){}
		}
	
	/** to create a doc
	 *     ObjectFactory factory = Context.getWmlObjectFactory();

    org.docx4j.wml.Body body = factory .createBody();

    org.docx4j.wml.Document wmlDocumentEl = factory .createDocument();

    wmlDocumentEl.setBody(body);

    

    // Put the content in the part

    wordDocumentPart.setJaxbElement(wmlDocumentEl);
    
    
    
    
        ObjectFactory factory = Context.getWmlObjectFactory();

    // Create the paragraph

    org.docx4j.wml.P para = factory.createP();

    // Create the text element

    org.docx4j.wml.Text t = factory.createText();

    t.setValue(simpleText);

  

    // Create the run

    org.docx4j.wml.R run = factory.createR();

    run.getRunContent().add(t);    

      

    para.getParagraphContent().add(run);

    // Now add our paragraph to the document body

    Body body = this.jaxbElement.getBody();

    Body.getEGBlockLevelElts().add(para)
    
	 */
	 
}
