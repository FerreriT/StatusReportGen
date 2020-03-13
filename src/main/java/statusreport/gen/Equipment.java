package statusreport.gen;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;


public class Equipment {

	private Sheet originalSheet;
	
	private int line;
	
	private List<Attribute> attributes;
	
	private List<String> kNomsChamps;
	
	private boolean hasChanged;
	
	private List<Integer> previousState;

	public Equipment() {}

	public Equipment(int line) {
		this.line=line;
	}
	
	public Equipment(Sheet originalSheet, int line) {
		this.originalSheet=originalSheet;
		this.line=line;
	}
	
	public Equipment(Sheet originalSheet, int line, List<Attribute> attributes) {
		super();
		this.originalSheet = originalSheet;
		this.line = line;
		this.attributes = attributes;
		this.kNomsChamps = new ArrayList<String>();
		for(Attribute attr:this.attributes) {
			this.kNomsChamps.add(attr.getName());
		}
		this.previousState = new ArrayList<Integer>();
	}

	public Sheet getOriginalSheet() {
		return originalSheet;
	}

	public void setOriginalSheet(Sheet originalSheet) {
		this.originalSheet = originalSheet;
	}

	public int getLine() {
		return line;
	}

	public void setLine(int line) {
		this.line = line;
	}

	public List<Attribute> getAttributes() {
		return attributes;
	}

	public void setAttributes(List<Attribute> attributes) {
		this.attributes = attributes;
	}
	
	public Attribute getAttribute(int i) {
		return attributes.get(i);
	}
	
	public void setAttribute(int i, Attribute attribute) {
		attributes.add(i, attribute);
	}
	
	public String getNomChamp(int i) {
		return kNomsChamps.get(i);
	}
	
	public void setNomChamp(int i, String nomChamp) {
		kNomsChamps.add(i, nomChamp);
	}
	
	public List<String> getNomsChamps() {
		return kNomsChamps;
	}
	
	public boolean hasChanged() {
		return hasChanged;
	}

	public void setHasChanged(boolean hasChanged) {
		this.hasChanged = hasChanged;
	}

	public List<Integer> getPreviousStates() {
		return previousState;
	}
	
	public int getPreviousState(int i) {
		return previousState.get(i);
	}

	public void addPreviousState(Integer i) {
		this.previousState.add(i);
	}
	
	public void setPreviousState(int index,int previousState) {
		this.previousState.set(index, previousState);
	}
}
