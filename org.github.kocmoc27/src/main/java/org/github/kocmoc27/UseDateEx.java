package org.github.kocmoc27;

import java.awt.EventQueue;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JFormattedTextField.AbstractFormatter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.UtilDateModel;

import net.miginfocom.swing.MigLayout;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.awt.event.ActionEvent;

public class UseDateEx {

	private JFrame frame;
	
	private JLabel useDateLabel;
	private JLabel reqNumLabel;
	private JLabel pathFolderLabel;
	
	private JDatePickerImpl datePickerUseDate; 
	private JTextField reqNumTxtBox;
	
	private JButton setUseDate;
	private JButton selectFileButton;
	private JFileChooser selectFile;
	
	private JButton selectFolderButton;
	private JFileChooser selectFolder;
	public static String pathFolder;
	public static String pathFile;
	
	public static ArrayList<String> CadNums = new ArrayList<>();
	public static ArrayList<String> dates = new ArrayList<>();
	public static ArrayList<String> cadNumChanges =  new ArrayList<>();
	public static String reqNum;
	public static String useDate;
	
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UseDateEx window = new UseDateEx();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public UseDateEx() {
		initialize();
	}

	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 250, 180);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		frame.getContentPane().setLayout(new MigLayout("", "[][grow][][]", "[][][][][][][][]"));
		reqNumLabel = new JLabel("Номер заявки:");
		reqNumTxtBox = new JTextField();
		
		pathFolderLabel = new JLabel("...");
		
		setUseDate = new JButton("Установить дату утверждения");
		setUseDate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				//System.out.println(pathFolder);				
				File file = new File(pathFolder);
				//System.out.println("");
				File[] folderEntries = file.listFiles();
			    for (File entry : folderEntries)
			    {
			    	System.out.println("Обход"); 
				    	if(entry.getName().contains(".xls"));{
				    	try {
							FileInputStream fis = new FileInputStream(entry.getPath());
							HSSFWorkbook wb = new HSSFWorkbook(fis);	
							HSSFSheet sheet = wb.getSheetAt(0);
					        Iterator<Row> rowIterator = sheet.iterator();
					        int flag = 1;
					        
					        while (rowIterator.hasNext()) { 
					         Row row = rowIterator.next();
				        	 if (flag == 1) {
				        		flag = 0; 
				        		row = rowIterator.next();
				        	 }
				        	  if(row.getCell(0) == null) CadNums.add("");
				        	  else { CadNums.add(getValueOfCell(row.getCell(0)));  }
						        		        	  
						      if(row.getCell(1) == null) dates.add(null);
						      else {
						    	  dates.add(row.getCell(1).getStringCellValue());
						      }
					        }  		        
						} catch (FileNotFoundException e1) {e1.printStackTrace();} 
							catch (IOException e1) {e1.printStackTrace();}	
			    	}				    					    
			    }   
			    
  
			   	reqNum = reqNumTxtBox.getText();

			   	
			   Connection con = connect("","" ,"" ,"" ,"" );				       
				        String sql = "SELECT DISTINCT o.cad_num "+
				        		"FROM zkoks.obj o, zkoks.reg r, request.request rq " +
				        		"WHERE o.id = r.obj_id AND r.request_id = rq.id " +
				        		"and rq.request_number = '"+ reqNum +"'"; 
				         
				        Statement stmt = null;
				       
				        try {
							stmt = con.createStatement();
							ResultSet rs = stmt.executeQuery(sql);
							
							while (rs.next())
							{  
								cadNumChanges.add(rs.getString(1)); 						
							}
						} catch (SQLException e1) {e1.printStackTrace();}
						
				        System.out.println("После"); 
				        
						Statement stmt2 = null; String sql2 = "";
						String log = "КН \t Дата применения \t Дата определения";
						
							
						try {
							stmt2 = con.createStatement();
						} catch (SQLException e2) {
							e2.printStackTrace();
						}
						
						for(int ind = 0;ind < cadNumChanges.size(); ind++) {
							System.out.println(cadNumChanges.get(ind) + " "+  searchDateEgroks(CadNums, dates, cadNumChanges.get(ind)));

							
							sql2 = "UPDATE zkoks.payment pmt "+    
								" set pmt.payment_date = null "+
							  ", PMT.USE_DATE =  '01.01.2020' " + // to_date('"+ searchDateEgroks(CadNums, dates, cadNumChanges.get(ind)) +"', 'dd.mm.yyyy'), " + //
								", pmt.definition_date = TO_DATE('"+ searchDateEgroks(CadNums, dates, cadNumChanges.get(ind)) +"', 'dd.mm.yyyy') "+   //searchDateEgroks(,cadNumChanges.get(ind)) 
								" where pmt.id in (select  pu.id from ZKOKS.OBJ o, zkoks.reg r, zkoks.payment pu, request.request rq "+
								                  " where O.CAD_NUM = '" + cadNumChanges.get(ind)
								                  + "' and o.id = r.obj_id and r.request_id = rq.id " 
								                  + " and rq.request_number = '" + reqNum
								                  +"' and r.id = Pu.REG_ID AND pu.CODE = 016010000000)";	
											
							try {
								
								stmt2.executeQuery(sql2);	

							} catch (SQLException e1) {e1.printStackTrace(); 
								
							}	
						}
						
			    }
		});
		
		selectFolderButton = new JButton("Выбрать папку");	
		selectFolderButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				selectFolder = new JFileChooser();    
				selectFolder.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int ret = selectFolder.showDialog(null, "Выбрать папку");                
                if (ret == JFileChooser.APPROVE_OPTION) {
                    File file = selectFolder.getSelectedFile();
                    pathFolder = file.getPath();
                   // System.out.println(pathFolder);
                    pathFolderLabel.setText(pathFolder);
                }
			}
		});
					
		frame.getContentPane().add(reqNumLabel,"cell 0 0,grow");
		frame.getContentPane().add(reqNumTxtBox,"cell 0 1,grow");
		
		//frame.getContentPane().add(selectFileButton, "cell 0 3,grow");
		frame.getContentPane().add(setUseDate, "cell 0 3,grow");
		frame.getContentPane().add(selectFolderButton, "cell 0 2,grow");
		frame.getContentPane().add(pathFolderLabel,"cell 0 4");
	}

	private static Connection connect(String address, String port, String SID,
			String login, String password) 
		{
			Connection con = null;
			try
			{
				Class.forName("oracle.jdbc.driver.OracleDriver");
				con = DriverManager.getConnection(
					"jdbc:oracle:thin:@" + address + ":" + port + ":" + SID, login,
					password);
			}
			catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return con;
		}
	
	private static String searchDateEgroks(ArrayList<String> arr, ArrayList<String> arr2, String CadNum) {
		String dateEgroks = "";
		for(int i = 0; i < arr.size(); i++) {
			if(arr.get(i).equals(CadNum)) {
				dateEgroks = arr2.get(i);
				StringBuffer buf = new StringBuffer(dateEgroks);
				buf.insert(7, "20");
				dateEgroks = buf.toString();
				String month = dateEgroks.substring(3, 6);	
				String buf2 = "";		
				switch(month) {
					case("JAN"): buf2 = dateEgroks.substring(0, 2)+ ".01." +dateEgroks.substring(7, 11); month = "01"; break;
					case("FEB"): buf2 = dateEgroks.substring(0, 2)+ ".02." +dateEgroks.substring(7, 11); month = "01"; break;
					case("MAR"): buf2 = dateEgroks.substring(0, 2)+ ".03." +dateEgroks.substring(7, 11); month = "01"; break;
					case("APR"): buf2 = dateEgroks.substring(0, 2)+ ".04." +dateEgroks.substring(7, 11); month = "01"; break;
					case("MAY"): buf2 = dateEgroks.substring(0, 2)+ ".05." +dateEgroks.substring(7, 11); month = "01"; break;
					case("JUN"): buf2 = dateEgroks.substring(0, 2)+ ".06." +dateEgroks.substring(7, 11); month = "01"; break;
					case("JUL"): buf2 = dateEgroks.substring(0, 2)+ ".07." +dateEgroks.substring(7, 11); month = "01"; break;
					case("AUG"): buf2 = dateEgroks.substring(0, 2)+ ".08." +dateEgroks.substring(7, 11); month = "01"; break;
					case("SEP"): buf2 = dateEgroks.substring(0, 2)+ ".09." +dateEgroks.substring(7, 11); month = "01"; break;
					case("OCT"): buf2 = dateEgroks.substring(0, 2)+ ".10." +dateEgroks.substring(7, 11); month = "01"; break;
					case("NOV"): buf2 = dateEgroks.substring(0, 2)+ ".11." +dateEgroks.substring(7, 11); month = "01"; break;
					case("DEC"): buf2 = dateEgroks.substring(0, 2)+ ".12." +dateEgroks.substring(7, 11); month = "01"; break;
					
					default: buf2 = dateEgroks.substring(0, 1)+ ".00." +dateEgroks.substring(7, 10); break;					
				}

				
				return buf2;
			}
		}
		System.out.println(dateEgroks);
		return dateEgroks;
	}
	
	private static String getValueOfCell(Cell cell) {               
		String res = "";
		CellType cellType = cell.getCellTypeEnum();

		  switch (cellType) {
		    case _NONE: res = ""; break;		        
		    case BOOLEAN:	res = String.valueOf(cell.getBooleanCellValue());break;      
		    case BLANK: System.out.print("");    
		    case NUMERIC:
		    	res = String.valueOf(cell.getNumericCellValue()); 
    	
		    	break;	       

		    case STRING: res = cell.getStringCellValue();break;	        
		    case ERROR: res = "!";break;
		  }
		return res;      	
	}
}