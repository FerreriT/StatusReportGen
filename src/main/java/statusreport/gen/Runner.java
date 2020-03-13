package statusreport.gen;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import javax.swing.JFileChooser;
import javax.swing.JFrame;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Runner {

	public static void main(String[] args) throws InvalidFormatException, IOException {
/*
		JFrame frame = new JFrame("");
		DemoJFileChooser panel = new DemoJFileChooser();
		frame.addWindowListener(
				new WindowAdapter() {
					public void windowClosing(WindowEvent e) {
						System.exit(0);
					}
				}
				);
		frame.getContentPane().add(panel,"Center");
		frame.setSize(panel.getPreferredSize());
		frame.setVisible(true);
		*/
		
		String path1 = "H:/Copy of CPT - ASI - Workload NEW 030320 (ProjetAutoStatusReport).xlsx";
		String path2 = "H:/Copy of CPT - ASI - Workload NEW 040320.xls";
		String path3 = "H:/Copy of CPT - RPM - Sorties.xls";
		
		Workbook previousWbk = WorkbookFactory.create(new File(path1));
		Workbook todayWbk = WorkbookFactory.create(new File(path2));
		
		Sheet previousSht = previousWbk.getSheet("Eqt list");
		Sheet todaySht = todayWbk.getSheet("Eqt list");
		
        DataFormatter dataFormatter = new DataFormatter();

        ClientEquipments comparison = new ClientEquipments(path1,path2);
        
        List<Equipment> listAFI1 = comparison.getkSheet();
        List<Equipment> listAFI2 = comparison.getNewSheet();
        
        List<Equipment> outlet = new ArrayList<Equipment>();
        
        //Maintenant ==> On repere les equipement manquants (avec MCO) et compare les champs
    	//qui nous interessent
        
        int count=0;
        int nbTot = listAFI2.size();
        for(Equipment equip1:listAFI1) {
        	String MCO1 = equip1.getAttributes().get(3).getValue();
        	for(Equipment equip2:listAFI2) {
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
        	}if(count==nbTot) outlet.add(equip1);count=0;
        }
        
        //Creation d'un workbook test pour lecture du resultat
        Workbook workbook1 = new XSSFWorkbook();
        Workbook workbook2 = new XSSFWorkbook();
        Sheet sheet1 = workbook1.createSheet("AFI equipments");
        Sheet sheet2 = workbook2.createSheet("AFI equipments");
        
        Font headerFont1 = workbook1.createFont();
        headerFont1.setBold(true);
        Font headerFont2 = workbook2.createFont();
        headerFont2.setBold(true);
        
        CellStyle headerCellStyle1 = workbook1.createCellStyle();
        headerCellStyle1.setFont(headerFont1);
        headerCellStyle1.setFillBackgroundColor((short)200);
        CellStyle headerCellStyle2 = workbook2.createCellStyle();
        headerCellStyle2.setFont(headerFont2);
        headerCellStyle2.setFillBackgroundColor((short)200);
        CellStyle changedCellStyle1 = workbook1.createCellStyle();
        changedCellStyle1.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        changedCellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        CellStyle changedCellStyle2 = workbook2.createCellStyle();
        changedCellStyle2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        changedCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        

        int nRow = 0;
        Row headerRow1 = sheet1.createRow(nRow);
        Equipment eqp1 = listAFI1.get(nRow);
        for(int i=0;i<eqp1.getAttributes().size();i++) {
        	Cell cell = headerRow1.createCell(i);
        	cell.setCellStyle(headerCellStyle1);
        	cell.setCellValue(eqp1.getNomChamp(i));
        }
        nRow++;
        List<Equipment> eqpChangedAttr = new ArrayList<Equipment>();
        for(Equipment eqp:listAFI1) {
        	Row row = sheet1.createRow(nRow++);
        	List<Attribute> attr = eqp.getAttributes();
        	int nc = 0;
        	for(int i=0; i<attr.size()-1; i++) {
        		int ic = 0;
        		if(eqp.hasChanged()&&nc<eqp.getPreviousStates().size()) {
        			ic = eqp.getPreviousState(nc);
        		}
        		Cell cell = row.createCell(i);
        		if(i==ic&&eqp.hasChanged()) {
        			cell.setCellStyle(changedCellStyle1);
        			eqpChangedAttr.add(eqp);
        			nc++;
        		}
        		cell.setCellValue(attr.get(i).getValue());
        	}if(eqp.hasChanged()) {
        		row.createCell(attr.size()-1).setCellValue("Yes");
        		
        	}
        }List<Attribute> attr1 = eqp1.getAttributes();
    	for(int i=0; i<attr1.size(); i++) {
    		sheet1.autoSizeColumn(i);;
    	}
    	nRow = 0;
        Row headerRow2 = sheet2.createRow(nRow);
        Equipment eqp2 = listAFI2.get(nRow);
        for(int i=0;i<eqp2.getAttributes().size();i++) {
        	Cell cell = headerRow2.createCell(i);
        	cell.setCellStyle(headerCellStyle2);
        	cell.setCellValue(eqp2.getNomChamp(i));
        }
        nRow++;
        for(Equipment eqp:listAFI2) {
        	Row row = sheet2.createRow(nRow++);
        	List<Attribute> attr = eqp.getAttributes();
        	int nc = 0;
        	for(int i=0; i<attr.size()-2; i++) {
        		int ic = 0;
        		if(eqp.hasChanged()&&nc<eqp.getPreviousStates().size()) {
        			ic = eqp.getPreviousState(nc);
        		}
        		Cell cell = row.createCell(i);
        		if(i==ic&&eqp.hasChanged()) {
        			cell.setCellStyle(changedCellStyle2);
        			nc++;
        		}
        		cell.setCellValue(attr.get(i).getValue());
        	}if(eqp.hasChanged()) {
        		row.createCell(attr.size()-2).setCellValue("Yes");
        		for(Equipment eqpch:eqpChangedAttr) {
        			if(eqpch.getAttribute(3).getValue().equals(eqp.getAttribute(3).getValue())) {
        				for(int i = 0; i<eqpch.getPreviousStates().size(); i++)	row.createCell(attr.size()-1+i)
        				.setCellValue(eqpch.getAttribute(eqpch.getPreviousState(i)).getValue());
        			}
        		}
        	}
        }List<Attribute> attr2 = eqp2.getAttributes();
    	for(int i=0; i<attr2.size(); i++) {
    		sheet2.autoSizeColumn(i);
    	}
    	
    	nRow+=2;
    	Row secHeaderRow = sheet2.createRow(nRow++);
    	Cell headerCell = secHeaderRow.createCell(0);
    	headerCell.setCellValue("Below, equipment outings");
    	
    	for(Equipment eqp:outlet) {
    		Row row = sheet2.createRow(nRow);
    		List<Attribute> attr = eqp.getAttributes();
    		for(int i=0; i<attr.size(); i++) {
        		row.createCell(i).setCellValue(attr.get(i).getValue());
        	}
    	}
    	
    	////
    	
    	Workbook wbks = WorkbookFactory.create(new File(path3));
    	Sheet sorties = wbks.getSheetAt(0);
    	int firstRown = sorties.getFirstRowNum();
    	int firstCelln = sorties.getRow(firstRown).getFirstCellNum();
    	//while(sorties.getRow(firstRown).getCell(cellnum)
    	
    	FileOutputStream fileOut1 = new FileOutputStream("poi-generated-file.xlsx");
        workbook1.write(fileOut1);
        fileOut1.close();
        FileOutputStream fileOut2 = new FileOutputStream("poi-generated-file-bis.xlsx");
        workbook2.write(fileOut2);
        fileOut2.close();
        workbook1.close();
        workbook2.close();
        
        
        	
	}

}
