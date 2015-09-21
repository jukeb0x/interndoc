package utils;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

public class WordHandler {

		public void launchWord(){
			try {
			     if (Desktop.isDesktopSupported()) {
			       Desktop.getDesktop().open(new File("ballons.docx"));
			     //  Desktop.getDesktop().edit(new File("rapport.doc"));
			     }
			   } catch (IOException ioe) {
			     ioe.printStackTrace();
			  }
		}
	
	
}
