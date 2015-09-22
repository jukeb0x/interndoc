package utils;



import java.io.File;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.xlsx4j.exceptions.Xlsx4jException;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.SheetData;



public class ExcelHandler {

	
	public void createExcel(){
		try{
			try{
				String outputfilepath = "test.xlsx";
		
				SpreadsheetMLPackage pkg = SpreadsheetMLPackage.createPackage();
				System.out.println("right");
				PartName partN=new PartName("sheet1.xml");
				System.out.println("now");
				WorksheetPart sheet = pkg.createWorksheetPart(partN, "Sheet1", 1);
			//	new PartName();
				//pkg.getWorkbookPart().addTargetPart(sheet);
			
				System.out.println("there");
				pkg.save(new File(outputfilepath));
				System.out.println("here");
			}
			catch(Docx4JException i){}
		}
		catch(JAXBException e){}
				

	}
	
	public void modifyExcel(){
		try{
		SpreadsheetMLPackage wordMLPackage = SpreadsheetMLPackage.createPackage();
		try{
			WorksheetPart sheet=wordMLPackage.getWorkbookPart().getWorksheet(0);
		// Minimal content already present
				SheetData sheetData = sheet.getJaxbElement().getSheetData();
						
				// Now add
				Row row = Context.getsmlObjectFactory().createRow();
				Cell cell = Context.getsmlObjectFactory().createCell();
				cell.setV("1234");
				
				// Note: if you are trying to add characters, not a number,
				// the easiest approach is to use inline strings (as opposed to the shared string table).
				// See http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2011/11/22/screen-cast-write-simpler-spreadsheetml-when-generating-spreadsheets.aspx
				// and http://www.docx4java.org/forums/xlsx-java-f15/cells-with-character-values-t874.html
				
				row.getC().add(cell);
				sheetData.getRow().add(row);}
		catch(Xlsx4jException e){}
		}catch(InvalidFormatException e2){}
	}
	
}
