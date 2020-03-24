package statusreport.gen;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Runner {

	private String path1;
	private String path2;
	private String path3;
	private String path4;
	
	private ClientEquipments getAllFields;
	
	private List<Equipment> listAFI1;
	private List<Equipment> listAFI2;

	private int nbOutingsFromFirstComp;
	private List<Equipment> wipOutings;
	private List<Equipment> longOutings;
	
	@SuppressWarnings("resource")
	public Runner() throws EncryptedDocumentException, InvalidFormatException, IOException {
	
		this.path1 = "H:/Copy of CPT - ASI - Workload NEW 030320 (ProjetAutoStatusReport).xlsx";
		this.path2 = "H:/Copy of CPT - ASI - Workload NEW 040320.xls";
		this.path3 = "H:/Copy of CPT - RPM - Sorties 070320.xls";
		this.path4 = "H:/Copy of CPT - RPM - Sorties 190320.xls";

		Workbook previousWbk = new XSSFWorkbook();
		previousWbk = WorkbookFactory.create(new File(path1));
		Sheet previousSht = previousWbk.getSheet("Eqt list");
		
		Workbook todayWbk = new XSSFWorkbook();
		todayWbk = WorkbookFactory.create(new File(path2));
		Sheet todaySht = todayWbk.getSheet("Eqt list");
		
		Workbook previousOutings = new XSSFWorkbook();
		previousOutings = WorkbookFactory.create(new File(path3));
		Sheet previousOutSht = previousOutings.getSheetAt(0);
		
		Workbook todayOutings = new XSSFWorkbook();
		todayOutings = WorkbookFactory.create(new File(path4));
		Sheet todayOutSht = todayOutings.getSheetAt(0);
		
		this.getAllFields = new ClientEquipments("AIR FRANCE INDUSTRIES",previousSht,todaySht, previousOutSht, todayOutSht);
		this.listAFI1 = getAllFields.getpreviousSheet();
		this.listAFI2 = getAllFields.getNewSheet();

		this.wipOutings = new ArrayList<Equipment>();
		this.longOutings = new ArrayList<Equipment>();
	}
	
	@SuppressWarnings("resource")
	public Runner(String path1, String path2, String path3, String path4, String sheetName, String custommerName) 
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		super();
		this.path1 = path1;
		this.path2 = path2;
		this.path3 = path3;
		this.path4 = path4;
		
		Workbook previousWbk = new XSSFWorkbook();
		previousWbk = WorkbookFactory.create(new File(this.path1));
		Sheet previousSht = previousWbk.getSheet(sheetName);
		
		Workbook todayWbk = new XSSFWorkbook();
		todayWbk = WorkbookFactory.create(new File(this.path2));
		Sheet todaySht = todayWbk.getSheet(sheetName);

		Workbook previousOutings = new XSSFWorkbook();
		previousOutings = WorkbookFactory.create(new File(this.path3));
		Sheet previousOutSht = previousOutings.getSheetAt(0);
		
		Workbook todayOutings = new XSSFWorkbook();
		todayOutings = WorkbookFactory.create(new File(this.path4));
		Sheet todayOutSht = todayOutings.getSheetAt(0);
		
		this.getAllFields = new ClientEquipments(custommerName, previousSht,todaySht, previousOutSht, todayOutSht);
		this.listAFI1 = getAllFields.getpreviousSheet();
		this.listAFI2 = getAllFields.getNewSheet();

		this.wipOutings = new ArrayList<Equipment>();
		this.longOutings = new ArrayList<Equipment>();
	}

	public void compareFields(List<Equipment> list1, List<Equipment> list2, int i) {
		switch(i) {
		case 1:
		int count=0;
        int nbTot = list2.size();
        for(Equipment equip1:list1) {
        	String MCO1 = equip1.getAttributes().get(3).getValue();
        	for(Equipment equip2:list2) {
            	String MCO2 = equip2.getAttributes().get(3).getValue();
            	if(MCO1.equals(MCO2)) {
            		int countn1 = 0;
            		for(String name1:equip1.getNomsChamps()) {
            			if(name1.contains("COMP")||name1.equals("WO Status Quote")
            			||name1.equals("Total Quotation (Q3)")||name1.equals("Nb Badged Hours")) {
            				int countn2 = 0;
            				for(String name2:equip2.getNomsChamps()) {
            					if(name1.equals(name2)) {
            						if(!equip1.getAttribute(countn1).getValue()
            								.equals(equip2.getAttribute(countn2).getValue())) {
            							equip1.setHasChanged(true);
            							equip1.addPreviousState(countn1);
            							equip2.setHasChanged(true);
            							equip2.addPreviousState(countn2);
            						}
            					}countn2++;
            				}
            			}countn1++;
            		}
            		break;
            	}count++;
        	}if(count==nbTot) this.wipOutings.add(equip1); count=0;
        }this.nbOutingsFromFirstComp = this.wipOutings.size();
        break;
		case 2:
			List<Equipment> further; List<Equipment> closer;
			if(list1.size()<list2.size()) {further = list1; closer = list2; }
			else{further = list2; closer = list1; }
			boolean test=false;
			for(int eq2 = 0; eq2 < closer.size(); eq2++) {
				if(test) eq2--;test=false;
				String MCO2="";List<Attribute> attrs2 = closer.get(eq2).getAttributes();
				for(int attr2=0; attr2<attrs2.size(); attr2++) {
					if(attrs2.get(attr2).getName().trim().equals("MCO")) { MCO2 = attrs2.get(attr2).getValue();break;}
				}
				for(int eq1 = 0; eq1 < further.size(); eq1++) {
					String MCO1="";List<Attribute> attrs1 = further.get(eq1).getAttributes();
					for(int attr1=0; attr1<attrs1.size(); attr1++) {
						if(attrs1.get(attr1).getName().trim().equals("MCO")) {MCO1 = attrs1.get(attr1).getValue();break;}
					}
					if(MCO2.equals(MCO1)) {closer.remove(eq2);test=true;break;}
				}
			}
			this.longOutings = closer;
			test=false;
			for(int eq2 = 0; eq2 < this.wipOutings.size(); eq2++) {
				if(test) eq2--;test=false;
				String MCO2="";List<Attribute> attrs2 = this.wipOutings.get(eq2).getAttributes();
				for(int attr2=0; attr2<attrs2.size(); attr2++) {
					if(attrs2.get(attr2).getName().trim().equals("MCO")) { MCO2 = attrs2.get(attr2).getValue();break;}
				}
				for(int eq1 = 0; eq1 < this.longOutings.size(); eq1++) {
					String MCO1="";List<Attribute> attrs1 = this.longOutings.get(eq1).getAttributes();
					for(int attr1=0; attr1<attrs1.size(); attr1++) {
						if(attrs1.get(attr1).getName().trim().equals("MCO")) {MCO1 = attrs1.get(attr1).getValue();break;}
					}
					if(MCO2.equals(MCO1)) {this.longOutings.remove(eq1);test=true;break;}
				}
			}for(Equipment eq:this.longOutings) this.wipOutings.add(eq);
		}
	}

	public static void main(String[] args) throws InvalidFormatException, IOException {

		String path1 = "H:/Copy of CPT - ASI - Workload NEW 030320 (ProjetAutoStatusReport).xlsx";
		String path2 = "H:/Copy of CPT - ASI - Workload NEW 090320.xls";
		String path3 = "H:/Copy of CPT - RPM - Sorties 070320.xls";
		String path4 = "H:/Copy of CPT - RPM - Sorties 190320.xls";
		String sheetName = "Eqt list";
		String custommerName = "AIR FRANCE INDUSTRIES";
		Runner itWorks = new Runner(path1,path2,path3,path4,sheetName, custommerName);
		
        //Maintenant ==> On repere les equipement manquants (avec MCO) et compare les champs
    	//qui nous interessent
        
		itWorks.compareFields(itWorks.getListAFI1(),itWorks.getListAFI2(),1);
		itWorks.compareFields(itWorks.getGetAllFields().getpreviousEqpt(), itWorks.getGetAllFields().getNewEqpt(), 2);
        
        //Creation d'un workbook test pour lecture du resultat
        
		ComparisonWriter wrtr = new ComparisonWriter();
		wrtr.setListAFI1(itWorks.getListAFI1());
		wrtr.setListAFI2(itWorks.getListAFI2());
		wrtr.setOutlet(itWorks.getOutlet(1));
		
        // Ecriture du resultat de la comparaison pour le premier workbook

		wrtr.wkbk1Writing();
		
    	// Ecriture du resultat de la comparaison du deuxieme workbook
    	
		wrtr.wkbk2Writing(itWorks.getNbOutingsFromFirstComp());
    	
    	// Sauvegarde et fermeture des workbooks generes
        
        wrtr.saveWkbk1("poi-generated-file.xlsx");
        wrtr.closeWkbk1();
        
        wrtr.saveWkbk2("poi-generated-file-bis.xlsx");
        wrtr.closeWkbk2();
        	
	}

	public String getPath1() {
		return path1;
	}

	public void setPath1(String path1) {
		this.path1 = path1;
	}

	public String getPath2() {
		return path2;
	}

	public void setPath2(String path2) {
		this.path2 = path2;
	}

	public String getPath3() {
		return path3;
	}

	public void setPath3(String path3) {
		this.path3 = path3;
	}

/*
 * 
	
	private Workbook previousWbk;
	private Workbook todayWbk;

	private Sheet previousSht;
	private Sheet todaySht;
	
	public Workbook getPreviousWbk() {
		return previousWbk;
	}

	public void setPreviousWbk(Workbook previousWbk) {
		this.previousWbk = previousWbk;
	}

	public Workbook getTodayWbk() {
		return todayWbk;
	}

	public void setTodayWbk(Workbook todayWbk) {
		this.todayWbk = todayWbk;
	}

	public Sheet getPreviousSht() {
		return previousSht;
	}

	public void setPreviousSht(Sheet previousSht) {
		this.previousSht = previousSht;
	}

	public Sheet getTodaySht() {
		return todaySht;
	}

	public void setTodaySht(Sheet todaySht) {
		this.todaySht = todaySht;
	}
*/
	public ClientEquipments getComparison() {
		return getAllFields;
	}

	public void setComparison(ClientEquipments comparison) {
		this.getAllFields = comparison;
	}

	public List<Equipment> getListAFI1() {
		return listAFI1;
	}

	public void setListAFI1(List<Equipment> listAFI1) {
		this.listAFI1 = listAFI1;
	}

	public List<Equipment> getListAFI2() {
		return listAFI2;
	}

	public void setListAFI2(List<Equipment> listAFI2) {
		this.listAFI2 = listAFI2;
	}

	public List<Equipment> getOutlet(int i) {
		switch(i) {
		case 1:
			return this.wipOutings;
		case 2:
			return this.longOutings;
		default:
			return this.wipOutings;
		}
	}

	public void setOutlet(List<Equipment> outlet, int i) {
		switch(i) {
		case 1:
			this.wipOutings = outlet;
			break;
		case 2:
			this.longOutings = outlet;
			break;
		default:
			this.wipOutings = outlet;
			break;
		}
	}

	public String getPath4() {
		return path4;
	}

	public void setPath4(String path4) {
		this.path4 = path4;
	}

	public ClientEquipments getGetAllFields() {
		return getAllFields;
	}

	public void setGetAllFields(ClientEquipments getAllFields) {
		this.getAllFields = getAllFields;
	}

	public List<Equipment> getWipOutings() {
		return wipOutings;
	}

	public void setWipOutings(List<Equipment> wipOutings) {
		this.wipOutings = wipOutings;
	}

	public List<Equipment> getLongOutings() {
		return longOutings;
	}

	public void setLongOutings(List<Equipment> longOutings) {
		this.longOutings = longOutings;
	}

	public int getNbOutingsFromFirstComp() {
		return nbOutingsFromFirstComp;
	}

	public void setNbOutingsFromFirstComp(int nbOutingsFromFirstComp) {
		this.nbOutingsFromFirstComp = nbOutingsFromFirstComp;
	}
}
