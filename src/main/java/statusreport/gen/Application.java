package statusreport.gen;

import java.awt.EventQueue;

import javax.swing.JFrame;
import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.layout.RowSpec;
import com.jgoodies.forms.layout.FormSpecs;
import javax.swing.JLabel;
import javax.swing.JTextField;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.Box;
import javax.swing.JTabbedPane;
import java.awt.BorderLayout;
import javax.swing.JPanel;
import com.jgoodies.forms.layout.Sizes;
import javax.swing.JScrollPane;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.awt.ComponentOrientation;
import java.awt.Frame;
import java.awt.Component;
import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.Insets;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.Color;
import java.beans.PropertyChangeListener;
import java.beans.PropertyChangeEvent;

public class Application {

	private JFrame frmStatusReportHelper;
	private JTextField path1;
	private JTextField path2;
	private JTextField path3;
	private JTextField path4;
	private final JFileChooser fc = new JFileChooser();

	private Runner itWorks;
	private JTextField sheetName;
	private JTextField customerName;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Application window = new Application();
					window.frmStatusReportHelper.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Application() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmStatusReportHelper = new JFrame();
		frmStatusReportHelper.setComponentOrientation(ComponentOrientation.LEFT_TO_RIGHT);
		frmStatusReportHelper.setAutoRequestFocus(false);
		frmStatusReportHelper.setTitle("Status Report Helper");
		frmStatusReportHelper.setBounds(400, 100, 1200, 1000);
		frmStatusReportHelper.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		GridBagLayout gridBagLayout = new GridBagLayout();
		gridBagLayout.columnWidths = new int[]{1166, 0};
		gridBagLayout.rowHeights = new int[]{912, 0};
		gridBagLayout.columnWeights = new double[]{0.0, Double.MIN_VALUE};
		gridBagLayout.rowWeights = new double[]{0.0, Double.MIN_VALUE};
		frmStatusReportHelper.getContentPane().setLayout(gridBagLayout);
		
		JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane.setMinimumSize(new Dimension(0, 0));
		tabbedPane.setMaximumSize(new Dimension(0, 0));
		tabbedPane.setPreferredSize(new Dimension(0, 0));
		GridBagConstraints gbc_tabbedPane = new GridBagConstraints();
		gbc_tabbedPane.anchor = GridBagConstraints.BASELINE;
		gbc_tabbedPane.fill = GridBagConstraints.BOTH;
		gbc_tabbedPane.gridx = 0;
		gbc_tabbedPane.gridy = 0;
		frmStatusReportHelper.getContentPane().add(tabbedPane, gbc_tabbedPane);
		
		JScrollPane scrollPane = new JScrollPane();
		tabbedPane.addTab("WIP & Outings Analyzise", null, scrollPane, null);
		tabbedPane.setEnabledAt(0, true);
		
		JPanel panelAnalyze = new JPanel();
		scrollPane.setRowHeaderView(panelAnalyze);
		GridBagLayout gbl_panelAnalyze = new GridBagLayout();
		gbl_panelAnalyze.columnWidths = new int[]{139, 481, 109, 77, 171, 0};
		gbl_panelAnalyze.rowHeights = new int[]{197, 33, 41, 33, 41, 33, 41, 33, 41, 43, 0, 0, 0, 42, 136, 41, 0};
		gbl_panelAnalyze.columnWeights = new double[]{0.0, 1.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		gbl_panelAnalyze.rowWeights = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		panelAnalyze.setLayout(gbl_panelAnalyze);
		
		JButton getPath2 = new JButton("Browser");
		getPath2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				int returnVal = fc.showOpenDialog(Application.this.frmStatusReportHelper);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
		            File file = fc.getSelectedFile();
		            path2.setText(file.getPath());
		            
		        }
			}
		});
		
		JLabel labelPath1 = new JLabel("Workload Last Data Path");
		GridBagConstraints gbc_labelPath1 = new GridBagConstraints();
		gbc_labelPath1.anchor = GridBagConstraints.NORTH;
		gbc_labelPath1.fill = GridBagConstraints.HORIZONTAL;
		gbc_labelPath1.insets = new Insets(0, 0, 5, 5);
		gbc_labelPath1.gridx = 1;
		gbc_labelPath1.gridy = 1;
		panelAnalyze.add(labelPath1, gbc_labelPath1);
		
		path1 = new JTextField();
		path1.setEnabled(false);
		path1.setEditable(false);
		path1.setFont(new Font("Times New Roman", Font.ITALIC, 20));
		path1.setText("Please set file path name");
		GridBagConstraints gbc_path1 = new GridBagConstraints();
		gbc_path1.fill = GridBagConstraints.HORIZONTAL;
		gbc_path1.insets = new Insets(0, 0, 5, 5);
		gbc_path1.gridwidth = 2;
		gbc_path1.gridx = 1;
		gbc_path1.gridy = 2;
		panelAnalyze.add(path1, gbc_path1);
		path1.setColumns(10);
		
		JButton getPath1 = new JButton("Browser");
		getPath1.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				int returnVal = fc.showOpenDialog(Application.this.frmStatusReportHelper);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
		            File file = fc.getSelectedFile();
		            path1.setText(file.getPath());
		            
		        }
			}
		});
		GridBagConstraints gbc_getPath1 = new GridBagConstraints();
		gbc_getPath1.anchor = GridBagConstraints.NORTH;
		gbc_getPath1.fill = GridBagConstraints.HORIZONTAL;
		gbc_getPath1.insets = new Insets(0, 0, 5, 0);
		gbc_getPath1.gridx = 4;
		gbc_getPath1.gridy = 2;
		panelAnalyze.add(getPath1, gbc_getPath1);
		
		JLabel labelPath2 = new JLabel("Workload Today Data Path");
		GridBagConstraints gbc_labelPath2 = new GridBagConstraints();
		gbc_labelPath2.anchor = GridBagConstraints.NORTH;
		gbc_labelPath2.fill = GridBagConstraints.HORIZONTAL;
		gbc_labelPath2.insets = new Insets(0, 0, 5, 5);
		gbc_labelPath2.gridx = 1;
		gbc_labelPath2.gridy = 3;
		panelAnalyze.add(labelPath2, gbc_labelPath2);
		
		path2 = new JTextField();
		path2.setEnabled(false);
		path2.setEditable(false);
		path2.setFont(new Font("Times New Roman", Font.ITALIC, 20));
		path2.setText("Please set file path name");
		GridBagConstraints gbc_path2 = new GridBagConstraints();
		gbc_path2.fill = GridBagConstraints.HORIZONTAL;
		gbc_path2.insets = new Insets(0, 0, 5, 5);
		gbc_path2.gridwidth = 2;
		gbc_path2.gridx = 1;
		gbc_path2.gridy = 4;
		panelAnalyze.add(path2, gbc_path2);
		path2.setColumns(10);
		GridBagConstraints gbc_getPath2 = new GridBagConstraints();
		gbc_getPath2.anchor = GridBagConstraints.NORTH;
		gbc_getPath2.fill = GridBagConstraints.HORIZONTAL;
		gbc_getPath2.insets = new Insets(0, 0, 5, 0);
		gbc_getPath2.gridx = 4;
		gbc_getPath2.gridy = 4;
		panelAnalyze.add(getPath2, gbc_getPath2);
		
		JLabel labelPath3 = new JLabel("Outings Last Data Path");
		GridBagConstraints gbc_labelPath3 = new GridBagConstraints();
		gbc_labelPath3.anchor = GridBagConstraints.NORTH;
		gbc_labelPath3.fill = GridBagConstraints.HORIZONTAL;
		gbc_labelPath3.insets = new Insets(0, 0, 5, 5);
		gbc_labelPath3.gridx = 1;
		gbc_labelPath3.gridy = 5;
		panelAnalyze.add(labelPath3, gbc_labelPath3);
		
		JButton getPath3 = new JButton("Browser");
		getPath3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				int returnVal = fc.showOpenDialog(Application.this.frmStatusReportHelper);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
		            File file = fc.getSelectedFile();
		            path3.setText(file.getPath());
		            
		        }
			}
		});
		
		path3 = new JTextField();
		path3.setEnabled(false);
		path3.setEditable(false);
		path3.setFont(new Font("Times New Roman", Font.ITALIC, 20));
		path3.setText("Please set file path name");
		GridBagConstraints gbc_path3 = new GridBagConstraints();
		gbc_path3.fill = GridBagConstraints.HORIZONTAL;
		gbc_path3.insets = new Insets(0, 0, 5, 5);
		gbc_path3.gridwidth = 2;
		gbc_path3.gridx = 1;
		gbc_path3.gridy = 6;
		panelAnalyze.add(path3, gbc_path3);
		path3.setColumns(10);
		GridBagConstraints gbc_getPath3 = new GridBagConstraints();
		gbc_getPath3.anchor = GridBagConstraints.NORTH;
		gbc_getPath3.fill = GridBagConstraints.HORIZONTAL;
		gbc_getPath3.insets = new Insets(0, 0, 5, 0);
		gbc_getPath3.gridx = 4;
		gbc_getPath3.gridy = 6;
		panelAnalyze.add(getPath3, gbc_getPath3);
		
		JLabel labelPath4 = new JLabel("Outings Today Data Path");
		GridBagConstraints gbc_labelPath4 = new GridBagConstraints();
		gbc_labelPath4.anchor = GridBagConstraints.NORTH;
		gbc_labelPath4.fill = GridBagConstraints.HORIZONTAL;
		gbc_labelPath4.insets = new Insets(0, 0, 5, 5);
		gbc_labelPath4.gridx = 1;
		gbc_labelPath4.gridy = 7;
		panelAnalyze.add(labelPath4, gbc_labelPath4);
		
		path4 = new JTextField();
		path4.setEnabled(false);
		path4.setEditable(false);
		path4.setFont(new Font("Times New Roman", Font.ITALIC, 20));
		path4.setText("Please set file path name");
		GridBagConstraints gbc_path4 = new GridBagConstraints();
		gbc_path4.fill = GridBagConstraints.HORIZONTAL;
		gbc_path4.insets = new Insets(0, 0, 5, 5);
		gbc_path4.gridwidth = 2;
		gbc_path4.gridx = 1;
		gbc_path4.gridy = 8;
		panelAnalyze.add(path4, gbc_path4);
		path4.setColumns(10);
		
		JButton getPath4 = new JButton("Browser");
		getPath4.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				int returnVal = fc.showOpenDialog(Application.this.frmStatusReportHelper);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
		            File file = fc.getSelectedFile();
		            path4.setText(file.getPath());
		            
		        }
			}
		});
		GridBagConstraints gbc_getPath4 = new GridBagConstraints();
		gbc_getPath4.anchor = GridBagConstraints.NORTH;
		gbc_getPath4.fill = GridBagConstraints.HORIZONTAL;
		gbc_getPath4.insets = new Insets(0, 0, 5, 0);
		gbc_getPath4.gridx = 4;
		gbc_getPath4.gridy = 8;
		panelAnalyze.add(getPath4, gbc_getPath4);
		
		JButton launcher = new JButton("Launch Analyzise");
		launcher.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				try {
					if(!(path1.getText()=="Please set file path name".trim()||path2.getText()=="Please set file path name".trim()||
							path3.getText()=="Please set file path name".trim()||path4.getText()=="Please set file path name".trim())) {
						itWorks = new Runner(path1.getText(),path2.getText(),path3.getText(),path4.getText(),sheetName.getText(),
								customerName.getText());
						//On repere les equipement manquants (avec MCO) et compare les champs
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
				        int ext = 3;
				        if(path1.getText().substring(path1.getText().length()-ext)!="xls") ext=4;
				        wrtr.saveWkbk1(path1.getText().substring(0, path1.getText().length()-ext)+"-changesHighlight.xlsx");
				        wrtr.closeWkbk1();
				        
				        if(ext==3&&path2.getText().substring(path2.getText().length()-ext)!="xls") ext=4;
				        else if(ext==4&&path2.getText().substring(path2.getText().length()-ext)!="xlsx") ext=3;
				        wrtr.saveWkbk2(path2.getText().substring(0, path2.getText().length()-ext)+"-changesHighlight-outingEquipment.xlsx");
				        wrtr.closeWkbk2();
					}
				} catch (EncryptedDocumentException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (InvalidFormatException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		
		JLabel labelSheetName = new JLabel("SheetName in workload workbook");
		GridBagConstraints gbc_labelSheetName = new GridBagConstraints();
		gbc_labelSheetName.fill = GridBagConstraints.HORIZONTAL;
		gbc_labelSheetName.anchor = GridBagConstraints.NORTH;
		gbc_labelSheetName.insets = new Insets(0, 0, 5, 5);
		gbc_labelSheetName.gridx = 1;
		gbc_labelSheetName.gridy = 10;
		panelAnalyze.add(labelSheetName, gbc_labelSheetName);
		
		sheetName = new JTextField();
		sheetName.setForeground(Color.GRAY);
		sheetName.setFont(new Font("Times New Roman", Font.ITALIC, 20));
		sheetName.setText("Eqt list");
		GridBagConstraints gbc_sheetName = new GridBagConstraints();
		gbc_sheetName.gridwidth = 2;
		gbc_sheetName.insets = new Insets(0, 0, 5, 5);
		gbc_sheetName.fill = GridBagConstraints.HORIZONTAL;
		gbc_sheetName.gridx = 1;
		gbc_sheetName.gridy = 11;
		panelAnalyze.add(sheetName, gbc_sheetName);
		sheetName.setColumns(10);
		
		JLabel labelCustomerName = new JLabel("Customer's Name");
		GridBagConstraints gbc_labelCustomerName = new GridBagConstraints();
		gbc_labelCustomerName.fill = GridBagConstraints.HORIZONTAL;
		gbc_labelCustomerName.anchor = GridBagConstraints.NORTH;
		gbc_labelCustomerName.insets = new Insets(0, 0, 5, 5);
		gbc_labelCustomerName.gridx = 1;
		gbc_labelCustomerName.gridy = 12;
		panelAnalyze.add(labelCustomerName, gbc_labelCustomerName);
		
		customerName = new JTextField();
		customerName.setForeground(Color.GRAY);
		customerName.setText("AIR FRANCE INDUSTRIES");
		customerName.setFont(new Font("Times New Roman", Font.ITALIC, 20));
		GridBagConstraints gbc_customerName = new GridBagConstraints();
		gbc_customerName.gridwidth = 2;
		gbc_customerName.insets = new Insets(0, 0, 5, 5);
		gbc_customerName.fill = GridBagConstraints.HORIZONTAL;
		gbc_customerName.gridx = 1;
		gbc_customerName.gridy = 13;
		panelAnalyze.add(customerName, gbc_customerName);
		customerName.setColumns(10);
		GridBagConstraints gbc_launcher = new GridBagConstraints();
		gbc_launcher.anchor = GridBagConstraints.NORTHWEST;
		gbc_launcher.gridwidth = 3;
		gbc_launcher.gridx = 2;
		gbc_launcher.gridy = 15;
		panelAnalyze.add(launcher, gbc_launcher);
		
		Application.this.frmStatusReportHelper.pack();
	}

}
