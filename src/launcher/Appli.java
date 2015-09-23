package launcher;
import utils.ExcelHandler;
import utils.XMLParser;
import utils.WordHandler;

public class Appli {
	 public static void main(String argv[]) {
		 XMLParser parser=new XMLParser();
		 //parser.XML_parse();
		 //parser.modifyMXL();
		 WordHandler handler=new WordHandler();
		 
		/* handler.createWord();
		 handler.modifyWord();
		handler.launchWord();*/
		ExcelHandler handler2=new ExcelHandler();
		handler2.makeExcel();
		 
	 }
	 

}
