package statusreport.gen;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class ClientEquipments {

	private String custommerName;

	private List<Equipment> previousSheet = new ArrayList<Equipment>();

	private List<Equipment> newSheet = new ArrayList<Equipment>();
	
	private List<Equipment> previousEqpt = new ArrayList<Equipment>();
	
	private List<Equipment> newEqpt = new ArrayList<Equipment>();
 
	public ClientEquipments() {

	}

	public ClientEquipments(String cName, String path1, String path2) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		this.custommerName = cName;
		
		Workbook wbk1 = WorkbookFactory.create(new File(path1));
		Sheet sheet1 = wbk1.getSheet("Eqt list");
		Workbook wbk2 = WorkbookFactory.create(new File(path2));
		Sheet sheet2 = wbk2.getSheet("Eqt list");

		this.buildLists(sheet1, 1);
		this.buildLists(sheet2, 2);
		
		wbk1.close();
		wbk2.close();
	}

	public ClientEquipments(String cName, File kWorkbook, File newWbk) throws InvalidFormatException, IOException {
		
		this.custommerName = cName;
		
		Workbook wbk1 = WorkbookFactory.create(kWorkbook);
		Sheet sheet1 = wbk1.getSheet("Eqt list");
		Workbook wbk2 = WorkbookFactory.create(newWbk);
		Sheet sheet2 = wbk2.getSheet("Eqt list");
		
		this.buildLists(sheet1, 1);
		this.buildLists(sheet2, 2);
		
		wbk1.close();
		wbk2.close();
	}
	
	public ClientEquipments(String cName, Sheet sheet1, Sheet sheet2, Sheet sorties1, Sheet sorties2) {
		
		this.custommerName = cName;
		
		this.buildLists(sheet1, 1);
		this.buildLists(sheet2, 2);
		this.buildLists(sorties1, 3);
		this.buildLists(sorties2, 4);
        
	}
	
	public void buildLists(Sheet sheet, int type) {

		DataFormatter dataFormatter = new DataFormatter();

		List<Attribute> standard1 = new ArrayList<Attribute>();

		int nbFirstRow=0;

		for(Row row:sheet) {
			if(!row.getCell(row.getFirstCellNum()).getStringCellValue().isEmpty()) {
				nbFirstRow = row.getRowNum();
				for(Cell cell:row) {
					if(!(cell == null || cell.getCellTypeEnum()==CellType.BLANK || cell.toString().isEmpty())) {
						Attribute std = new Attribute(dataFormatter.formatCellValue(cell));
						standard1.add(std);
					}
				}break;
			}
		}for (Row row:sheet) {
			Cell customer = row.getCell(row.getFirstCellNum()+2);
			if (dataFormatter.formatCellValue(customer).toString().equals(this.custommerName)) {
				int n = row.getRowNum();
				List<Attribute> newstand = new ArrayList<Attribute>();
				int j=0;
				for(int i = 0; i<standard1.size(); i++) {
					while((sheet.getRow(nbFirstRow).getCell(i+j) == null || sheet.getRow(nbFirstRow).getCell(i+j).
							getCellTypeEnum()==CellType.BLANK || sheet.getRow(nbFirstRow).getCell(i+j).toString().isEmpty())) j++;
					newstand.add(new Attribute(standard1.get(i).getName(),dataFormatter.formatCellValue(row.getCell(i+j))));
				}Equipment equip;
				switch(type) {
				case 1:
					newstand.add(new Attribute("Has Changed","No"));
	        		equip = new Equipment(sheet,n,newstand);
	        		this.previousSheet.add(equip);
					break;
				case 2:
					newstand.add(new Attribute("Has Changed","No"));
	        		newstand.add(new Attribute("Last State of this Data",""));
					equip = new Equipment(sheet,n,newstand);
					this.newSheet.add(equip);
					break;
				case 3:
					equip = new Equipment(sheet,n,newstand);
					this.previousEqpt.add(equip);
					break;
				case 4:
					equip = new Equipment(sheet,n,newstand);
					this.newEqpt.add(equip);
				}
			}
			
		}

	}
	
	public List<Equipment> getpreviousSheet() {
		return previousSheet;
	}

	public void setpreviousSheet(List<Equipment> previousSheet) {
		this.previousSheet = previousSheet;
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

	public List<Equipment> getpreviousEqpt() {
		return previousEqpt;
	}

	public void setpreviousEqpt(List<Equipment> previousEqpt) {
		this.previousEqpt = previousEqpt;
	}

	public List<Equipment> getNewEqpt() {
		return newEqpt;
	}

	public void setNewEqpt(List<Equipment> newEqpt) {
		this.newEqpt = newEqpt;
	}

}
