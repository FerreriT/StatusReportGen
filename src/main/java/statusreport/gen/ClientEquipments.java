package statusreport.gen;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.swing.JFileChooser;


public class ClientEquipments {

	private String custommerName;

	private List<Equipment> kSheet = new ArrayList<Equipment>();

	private List<Equipment> newSheet = new ArrayList<Equipment>();
	
	private List<Equipment> outingEqpt = new ArrayList<Equipment>();
 
	public ClientEquipments() {

	}

	public ClientEquipments(String cName, String path1, String path2) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		this.custommerName = cName;
		
		Workbook wbk1 = WorkbookFactory.create(new File(path1));
		Sheet sheet1 = wbk1.getSheet("Eqt list");
		Workbook wbk2 = WorkbookFactory.create(new File(path2));
		Sheet sheet2 = wbk2.getSheet("Eqt list");

		DataFormatter dataFormatter = new DataFormatter();

		List<Attribute> standard1 = new ArrayList<Attribute>();
        
        for(Row row:sheet1) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard1.add(std);
        		}break;
        	}
        }for (Row row:sheet1) {
        	Cell customer = row.getCell(2);
        	if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
        		int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
        		for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard1.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
        		}newstand.add(new Attribute("Has Changed","No"));
        		Equipment equip = new Equipment(sheet1,n,newstand);
        		this.kSheet.add(equip);
        	}
        }
        
        List<Attribute> standard2 = new ArrayList<Attribute>();
        
        for(Row row:sheet2) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard2.add(std);
        		}break;
        	}
        }for (Row row:sheet2) {
			Cell customer = row.getCell(2);
			if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
				int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
				for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard2.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
				}newstand.add(new Attribute("Has Changed","No"));
        		newstand.add(new Attribute("Last State of this Data",""));
				Equipment equip = new Equipment(sheet2,n,newstand);
				this.newSheet.add(equip);
			}
		}
		wbk1.close();
		wbk2.close();
	}

	public ClientEquipments(String cName, File kWorkbook, File newWbk) throws InvalidFormatException, IOException {
		
		this.custommerName = cName;
		
		Workbook wbk1 = WorkbookFactory.create(kWorkbook);
		Sheet sheet1 = wbk1.getSheet("Eqt list");
		Workbook wbk2 = WorkbookFactory.create(newWbk);
		Sheet sheet2 = wbk2.getSheet("Eqt list");
		
		DataFormatter dataFormatter = new DataFormatter();

		List<Attribute> standard1 = new ArrayList<Attribute>();
        
        for(Row row:sheet1) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard1.add(std);
        		}break;
        	}
        }for (Row row:sheet1) {
        	Cell customer = row.getCell(2);
        	if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
        		int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
        		for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard1.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
        		}newstand.add(new Attribute("Has Changed","No"));
        		Equipment equip = new Equipment(sheet1,n,newstand);
        		this.kSheet.add(equip);
        	}
        }
        
        List<Attribute> standard2 = new ArrayList<Attribute>();
        
        for(Row row:sheet2) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard2.add(std);
        		}break;
        	}
        }for (Row row:sheet2) {
			Cell customer = row.getCell(2);
			if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
				int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
				for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard2.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
				}newstand.add(new Attribute("Has Changed","No"));
        		newstand.add(new Attribute("Last State of this Data",""));
				Equipment equip = new Equipment(sheet2,n,newstand);
				this.newSheet.add(equip);
			}
		}
		wbk1.close();
		wbk2.close();
	}
	
	public ClientEquipments(String cName, Sheet sheet1, Sheet sheet2, Sheet sorties1, Sheet sorties2) {
		
		this.custommerName = cName;
		
		DataFormatter dataFormatter = new DataFormatter();

		List<Attribute> standard1 = new ArrayList<Attribute>();
        
        for(Row row:sheet1) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard1.add(std);
        		}break;
        	}
        }for (Row row:sheet1) {
        	Cell customer = row.getCell(2);
        	if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
        		int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
        		for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard1.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
        		}newstand.add(new Attribute("Has Changed","No"));
        		Equipment equip = new Equipment(sheet1,n,newstand);
        		this.kSheet.add(equip);
        	}
        }
        
        List<Attribute> standard2 = new ArrayList<Attribute>();
        
        for(Row row:sheet2) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard2.add(std);
        		}break;
        	}
        }for (Row row:sheet2) {
			Cell customer = row.getCell(2);
			if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
				int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
				for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard2.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
				}newstand.add(new Attribute("Has Changed","No"));
        		newstand.add(new Attribute("Last State of this Data",""));
				Equipment equip = new Equipment(sheet2,n,newstand);
				this.newSheet.add(equip);
			}
		}
        
        
	}
	
	public void buildLists(Sheet sheet) {
		
		DataFormatter dataFormatter = new DataFormatter();

		List<Attribute> standard1 = new ArrayList<Attribute>();
        
        for(Row row:sheet) {
        	if(!row.getCell(1).getStringCellValue().isEmpty()) {
        		for(Cell cell:row) {
        			Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
        			standard1.add(std);
        		}break;
        	}
        }for (Row row:sheet) {
        	Cell customer = row.getCell(2);
        	if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
        		int n = row.getRowNum();
        		List<Attribute> newstand = new ArrayList<Attribute>();
        		for(int i = 0; i<row.getLastCellNum(); i++) {
        			newstand.add(new Attribute(standard1.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i))));
        		}newstand.add(new Attribute("Has Changed","No"));
        		Equipment equip = new Equipment(sheet,n,newstand);
        		this.kSheet.add(equip);
        	}
        }
        
	}
	
	public List<Equipment> getkSheet() {
		return kSheet;
	}

	public void setkSheet(List<Equipment> kSheet) {
		this.kSheet = kSheet;
	}

	public List<Equipment> getNewSheet() {
		return newSheet;
	}

	public void setNewSheet(List<Equipment> newSheet) {
		this.newSheet = newSheet;
	}

	public String getCustommerName() {
		return custommerName;
	}

	public void setCustommerName(String custommerName) {
		this.custommerName = custommerName;
	}

	public List<Equipment> getOutingEqpt() {
		return outingEqpt;
	}

	public void setOutingEqpt(List<Equipment> outingEqpt) {
		this.outingEqpt = outingEqpt;
	}

}
