package statusreport.gen;

public class Attribute {

	private String name;
	
	private String value;
	
	public Attribute(String name) {
		this.name=name;
	}

	public Attribute(String name, String valeur) {
		super();
		this.name = name;
		this.value = valeur;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String valeur) {
		this.value = valeur;
	}
	
}
