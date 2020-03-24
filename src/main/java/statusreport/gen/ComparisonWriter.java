package statusreport.gen;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ComparisonWriter {

	private Workbook workbook1;
    private Workbook workbook2;
    private Sheet sheet1;
    private Sheet sheet2;
    
    private Font headerFont1;
    private Font headerFont2;
    
    private CellStyle headerCellStyle1;
    private CellStyle headerCellStyle2;
    
    private CellStyle changedCellStyle1;
    private CellStyle changedCellStyle2;
	
    private List<Equipment> listAFI1;
    private List<Equipment> listAFI2;
    
    private List<Equipment> eqpChangedAttr;
	private List<Equipment> outlet;
	
	private String path;
    
    public ComparisonWriter() {
    	
    	this.workbook1 = new XSSFWorkbook();
    	
        this.sheet1 = this.workbook1.createSheet("AFI equipments");
        
        this.headerFont1 = this.workbook1.createFont();
        this.headerFont1.setBold(true);
        
        this.headerCellStyle1 = this.workbook1.createCellStyle();
        this.headerCellStyle1.setFont(headerFont1);
        this.headerCellStyle1.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        this.headerCellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        this.changedCellStyle1 = this.workbook1.createCellStyle();
        this.changedCellStyle1.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        this.changedCellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        this.listAFI1 = new ArrayList<Equipment>();
        
        
        this.workbook2 = new XSSFWorkbook();
        
        this.sheet2 = this.workbook2.createSheet("AFI equipments");
        
        this.headerFont2 = this.workbook2.createFont();
        this.headerFont2.setBold(true);
        
        this.headerCellStyle2 = this.workbook2.createCellStyle();
        this.headerCellStyle2.setFont(headerFont2);
        this.headerCellStyle2.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        this.headerCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        this.changedCellStyle2 = this.workbook2.createCellStyle();
        this.changedCellStyle2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        this.changedCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        this.listAFI2 = new ArrayList<Equipment>();
        
        this.eqpChangedAttr = new ArrayList<Equipment>();
        this.outlet = new ArrayList<Equipment>();
        
        this.path = "";
    }
    
    public void buildOutlet(List<Equipment> outlet1, List<Equipment> outlet2) {
    	
    }

	public void wkbk1Writing() {
    	int nRow = 0;
        Row headerRow1 = this.sheet1.createRow(nRow);
        Equipment eqp1 = this.listAFI1.get(nRow);
        for(int i=0;i<eqp1.getAttributes().size();i++) {
        	Cell cell = headerRow1.createCell(i);
        	cell.setCellStyle(this.headerCellStyle1);
        	cell.setCellValue(eqp1.getNomChamp(i));
        }
        nRow++;
        for(Equipment eqp:this.listAFI1) {
        	Row row = this.sheet1.createRow(nRow++);
        	List<Attribute> attr = eqp.getAttributes();
        	int nc = 0;
        	for(int i=0; i<attr.size()-1; i++) {
        		int ic = 0;
        		if(eqp.hasChanged()&&nc<eqp.getPreviousStates().size()) {
        			ic = eqp.getPreviousState(nc);
        		}
        		Cell cell = row.createCell(i);
        		if(i==ic&&eqp.hasChanged()) {
        			cell.setCellStyle(this.changedCellStyle1);
        			this.eqpChangedAttr.add(eqp);
        			nc++;
        		}
        		cell.setCellValue(attr.get(i).getValue());
        	}if(eqp.hasChanged()) {
        		row.createCell(attr.size()-1).setCellValue("Yes");
        		
        	}
        }List<Attribute> attr1 = eqp1.getAttributes();
    	for(int i=0; i<attr1.size(); i++) {
    		this.sheet1.autoSizeColumn(i);;
    	}
    }
    
    public void wkbk2Writing(int n) {
    	int nRow = 0;
        Row headerRow2 = this.sheet2.createRow(nRow);
        Equipment eqp2 = this.listAFI2.get(nRow);
        for(int i=0;i<eqp2.getAttributes().size();i++) {
        	Cell cell = headerRow2.createCell(i);
        	cell.setCellStyle(this.headerCellStyle2);
        	cell.setCellValue(eqp2.getNomChamp(i));
        }
        nRow++;
        for(Equipment eqp:this.listAFI2) {
        	Row row = this.sheet2.createRow(nRow++);
        	List<Attribute> attr = eqp.getAttributes();
        	int nc = 0;
        	for(int i=0; i<attr.size()-2; i++) {
        		int ic = 0;
        		if(eqp.hasChanged()&&nc<eqp.getPreviousStates().size()) {
        			ic = eqp.getPreviousState(nc);
        		}
        		Cell cell = row.createCell(i);
        		if(i==ic&&eqp.hasChanged()) {
        			cell.setCellStyle(this.changedCellStyle2);
        			nc++;
        		}
        		cell.setCellValue(attr.get(i).getValue());
        	}if(eqp.hasChanged()) {
        		row.createCell(attr.size()-2).setCellValue("Yes");
        		for(Equipment eqpch:this.eqpChangedAttr) {
        			if(eqpch.getAttribute(3).getValue().equals(eqp.getAttribute(3).getValue())) {
        				for(int i = 0; i<eqpch.getPreviousStates().size(); i++) {
        					if(eqpch.getAttribute(eqpch.getPreviousState(i)).getValue().trim().isEmpty())
        						row.createCell(attr.size()-1+i).setCellValue("Empty");
        					else row.createCell(attr.size()-1+i).setCellValue(eqpch.getAttribute(eqpch.getPreviousState(i))
        							.getValue());
        				}
        			}
        		}
        	}
        }List<Attribute> attr2 = eqp2.getAttributes();
    	for(int i=0; i<attr2.size(); i++) {
    		this.sheet2.autoSizeColumn(i);
    	}
    	
    	nRow+=2;
    	Row secHeaderRow = this.sheet2.createRow(nRow++);
    	Cell headerCell = secHeaderRow.createCell(0);
    	headerCell.setCellValue("Below, equipment outings from last time");
    	int compt = 0;Row row;
    	for(Equipment eqp:this.outlet) {
    		row = this.sheet2.createRow(nRow++);
    		if(compt==n) {
    			Row secHeaderRow2 = this.sheet2.createRow(nRow++);
    	    	Equipment eqpo = this.outlet.get(n);
    	    	for(int i=0;i<eqpo.getAttributes().size();i++) {
    	        	Cell cell = secHeaderRow2.createCell(i);
    	        	cell.setCellStyle(this.headerCellStyle2);
    	        	cell.setCellValue(eqpo.getNomChamp(i));
    	        }row = this.sheet2.createRow(nRow++);
    		}
    		List<Attribute> attr = eqp.getAttributes();
    		for(int i=0; i<attr.size(); i++) {
        		row.createCell(i).setCellValue(attr.get(i).getValue());
        	}compt++;
    	}List<Attribute> attr3 = this.outlet.get(n).getAttributes();
    	for(int i=0; i<attr3.size(); i++) {
    		this.sheet2.autoSizeColumn(i);
    	}
    }
    
    public void saveWkbk1(String fileName) throws IOException {
    	FileOutputStream fileOut1 = new FileOutputStream(new File(fileName));
        this.workbook1.write(fileOut1);
        fileOut1.close();
    }
    
    public void saveWkbk2(String fileName) throws IOException {
    	FileOutputStream fileOut2 = new FileOutputStream(new File(fileName));
        this.workbook2.write(fileOut2);
        fileOut2.close();
    }
    
    public void closeWkbk1() throws IOException {
    	this.workbook1.close();
    }
    
    public void closeWkbk2() throws IOException {
    	this.workbook2.close();
    }
    
    public Workbook getWorkbook1() {
		return workbook1;
	}

	public void setWorkbook1(Workbook workbook1) {
		this.workbook1 = workbook1;
	}

	public Workbook getWorkbook2() {
		return workbook2;
	}

	public void setWorkbook2(Workbook workbook2) {
		this.workbook2 = workbook2;
	}

	public Sheet getSheet1() {
		return sheet1;
	}

	public void setSheet1(Sheet sheet1) {
		this.sheet1 = sheet1;
	}

	public Sheet getSheet2() {
		return sheet2;
	}

	public void setSheet2(Sheet sheet2) {
		this.sheet2 = sheet2;
	}

	public Font getHeaderFont1() {
		return headerFont1;
	}

	public void setHeaderFont1(Font headerFont1) {
		this.headerFont1 = headerFont1;
	}

	public Font getHeaderFont2() {
		return headerFont2;
	}

	public void setHeaderFont2(Font headerFont2) {
		this.headerFont2 = headerFont2;
	}

	public CellStyle getHeaderCellStyle1() {
		return headerCellStyle1;
	}

	public void setHeaderCellStyle1(CellStyle headerCellStyle1) {
		this.headerCellStyle1 = headerCellStyle1;
	}

	public CellStyle getHeaderCellStyle2() {
		return headerCellStyle2;
	}

	public void setHeaderCellStyle2(CellStyle headerCellStyle2) {
		this.headerCellStyle2 = headerCellStyle2;
	}

	public CellStyle getChangedCellStyle1() {
		return changedCellStyle1;
	}

	public void setChangedCellStyle1(CellStyle changedCellStyle1) {
		this.changedCellStyle1 = changedCellStyle1;
	}

	public CellStyle getChangedCellStyle2() {
		return changedCellStyle2;
	}

	public void setChangedCellStyle2(CellStyle changedCellStyle2) {
		this.changedCellStyle2 = changedCellStyle2;
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
    
    public List<Equipment> getOutlet() {
		return outlet;
	}

	public void setOutlet(List<Equipment> outlet) {
		this.outlet = outlet;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}
    
}
