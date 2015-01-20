package Cross;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

import jxl.Cell;
import jxl.CellType;
import jxl.FormulaCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.biff.formula.FormulaException;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.nio.file.CopyOption;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

public class QTPDataTable 
{
	public static int CurrentRow = 1;
	public static String XLS = "Global.xls"; 
	public static String XST = "Global";
	public QTPDataTable() throws IOException, WriteException{
	  WritableWorkbook wworkbook;
	  wworkbook = Workbook.createWorkbook(new File(XLS));
      WritableSheet wsheet = wworkbook.createSheet(XST, 0);
      wworkbook.write();
      wworkbook.close();
	}
  /* public static void main(String[] args) throws BiffException, IOException, WriteException {   }*/
   
   public void Import(String xls) throws IOException {//Over
		Path FROM = Paths.get(xls);
		Path TO = Paths.get(XLS);
		CopyOption[] options = new CopyOption[]{
		  StandardCopyOption.REPLACE_EXISTING,
		  StandardCopyOption.COPY_ATTRIBUTES
		};  
		Files.copy(FROM, TO, options);
   }
   // Fix this in C:\Softwares\eclipse\plugins\jexcelapi\jxl.jar -  C:\Softwares\eclipse\plugins\jexcelapi\src\jxl\write\WritableSheet.java
   public void ImportSheet(String xls,String xst) throws BiffException, IOException, WriteException{//importing specified sheet from sourse
		WritableWorkbook wworkbook;
		WritableWorkbook wworkbook1;
		wworkbook1 = Workbook.createWorkbook(new File("tmp.xls"));
	    WritableSheet wsheet = wworkbook1.createSheet(xst, 1);
	    wworkbook1.write();
	    wworkbook1.close();
		System.out.println(Workbook.getWorkbook(new File(xls)).getSheet(xst).getRows());

	    Workbook W;
		W = Workbook.getWorkbook(new File(XLS));
	    wworkbook = Workbook.createWorkbook(new File(XLS),W);
	    int found = 0;
	    for(int i=0;i<W.getNumberOfSheets();i++){
	    	System.out.println(W.getSheetNames()[i]);
	    	if(W.getSheetNames()[i].equalsIgnoreCase(xst)){
	    		found = 1;
	    		break;
	    	}
	    }
		System.out.println(wworkbook.getSheets().length+1);
		System.out.println(found);
		if(found == 0){
			wworkbook.importSheet(xst,wworkbook.getSheets().length+1 , Workbook.getWorkbook(new File(xls)).getSheet(xst));
		}
	    wworkbook.write();
	    wworkbook.close(); 	   
   }

   public void Export(String xls) throws IOException{//over
		Path FROM = Paths.get(XLS);
		Path TO = Paths.get(xls);
		CopyOption[] options = new CopyOption[]{
		  StandardCopyOption.REPLACE_EXISTING,
		  StandardCopyOption.COPY_ATTRIBUTES
		};  
		Files.copy(FROM, TO, options);
   }
   // Fix this in C:\Softwares\eclipse\plugins\jexcelapi\jxl.jar -  C:\Softwares\eclipse\plugins\jexcelapi\src\jxl\write\WritableSheet.java
   public void ExportSheet(String xls, String xst) throws BiffException, IOException, WriteException{
		WritableWorkbook wworkbook = null;
	    Workbook W;
		W = Workbook.getWorkbook(new File(XLS));

	    int found = 0;
	    for(int i=0;i<W.getNumberOfSheets();i++){
	    	System.out.println(W.getSheetNames()[i]);
	    	if(W.getSheetNames()[i].equalsIgnoreCase(xst)){
	    		found = 1;
	    		break;
	    	}
	    }
	    System.out.println(found);
		if(found == 0){
		    wworkbook = Workbook.createWorkbook(new File(xls),Workbook.getWorkbook(new File(xls)));
			wworkbook.importSheet(xst,wworkbook.getSheets().length+1 ,W.getSheet(xst));
		}
	    wworkbook.write();
	    wworkbook.close();
	 
   }
   
   public void AddSheet(String xst)throws BiffException, IOException, WriteException {//Over //adding empty sheet
	WritableWorkbook wworkbook;
	WritableWorkbook wworkbook1;
    Workbook W;
	wworkbook1 = Workbook.createWorkbook(new File("tmp.xls"));
    WritableSheet wsheet = wworkbook1.createSheet(xst, 1);
    wworkbook1.write();
    wworkbook1.close();	
	W = Workbook.getWorkbook(new File(XLS));
    wworkbook = Workbook.createWorkbook(new File(XLS),W);
    int found = 0;
    for(int i=0;i<W.getNumberOfSheets();i++){
    	System.out.println(W.getSheetNames()[i]);
    	if(W.getSheetNames()[i].equalsIgnoreCase(xst)){
    		found = 1;
    		break;
    	}
    }
	if(found == 0){
		wworkbook.importSheet(xst,wworkbook.getSheets().length+1 , Workbook.getWorkbook(new File("tmp.xls")).getSheet(xst));
	}
    wworkbook.write();
    wworkbook.close();    
   }
   
   public void DeleteSheet(String xst) throws IOException, WriteException{//Over
	   WritableWorkbook wworkbook;
		int intC = 0;
	   wworkbook = Workbook.createWorkbook(new File(XLS));
	    intC = wworkbook.getSheetNames().length;	    
	    for( int intX=0;intX<intC; intX++){
	    	 System.out.println( wworkbook.getSheetNames()[intX]);
	    	 if( wworkbook.getSheetNames()[intX].equalsIgnoreCase(xst))  {
	    		 wworkbook.removeSheet(intX);
	    		 break;
	    	 }
	    }
	    wworkbook.write();
	    wworkbook.close();
   }   

   public void GetSheet(String xst){//Over
	   XST = xst;
   }
   
   public int GetCurrentRow(){//Over
		return CurrentRow;		   
   }
   
   public int GetRowCount() throws BiffException, IOException{//Over	   
	Workbook workbook = Workbook.getWorkbook(new File(XLS));
	if(XST.length()>0){
		Sheet sheet = workbook.getSheet(XST);
		//System.out.println(sheet.getRows());
		return (sheet.getRows());
	}
	return -1;
   }
   
   public int GetRowCount(String xst) throws BiffException, IOException{	  //Over 
	   Workbook workbook = Workbook.getWorkbook(new File(XLS));
		Sheet sheet = workbook.getSheet(xst);
		//System.out.println(sheet.getRows());
		return (sheet.getRows());	   
   }
   
   public int CurrentRow(){//over
	return CurrentRow;	   
   }
   
   public void SetCurrentRow(int Row){//over
	 CurrentRow = Row;	   
   }
   
   public void SetNextRow(){//Over
	 CurrentRow = CurrentRow + 1;	   
   }
   
   public void SetPreviousRow(){//Over
	 CurrentRow = CurrentRow - 1;	   
   }
   
   public String GetValue(String Col) throws BiffException, IOException{//Get Sheet + Set Row
	  Workbook workbook = Workbook.getWorkbook(new File(XLS));
	  Sheet sheet = workbook.getSheet(XST);
	  int rows=0;
	  int cols=0;
	  int R=0;
	  rows=sheet.getRows();
	  cols=sheet.getColumns();
	  for(R=0; R<cols; R++){
		  Cell cellx = sheet.getCell(R, 0);
		  if(cellx.getContents().equalsIgnoreCase(Col)){
			break;
		  }
	  }
	  Cell cell1 = sheet.getCell(R,CurrentRow-1);
	  String x = cell1.getContents();
	  workbook.close();
	  return x;
   }   
  
   public String GetValue(int Ro, int Col, String xst) throws BiffException, IOException{
		  Workbook workbook = Workbook.getWorkbook(new File(XLS));
		  Sheet sheet = workbook.getSheet(xst);
		  int rows=0;
		  int cols=0;
		  Cell cell1 = sheet.getCell(Col-1,Ro-1);
		  String x = cell1.getContents();
		  workbook.close();
		  return x;
   }   

   public String RawValue(int Ro,int Col, String xst) throws BiffException, IOException, FormulaException{
	Workbook workbook = Workbook.getWorkbook(new File(XLS));
	Sheet sheet = workbook.getSheet(xst);
	int rows=0;
	int cols=0;
	String x = null;
	
	try {
		FormulaCell formula5 = (FormulaCell) sheet.getCell(Col-1,Ro-1);
		x = formula5.getFormula();
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	  workbook.close();
	  return x;	  
   }
   
   public void SetValue(String val, int Ro, int Col, String xst) throws BiffException, IOException, WriteException{
	  Workbook w1 = Workbook.getWorkbook(new File("NSE.xls"));
	  WritableWorkbook w2 = Workbook.createWorkbook(new File("NSE.xls"), w1);
	  Cell cell = w2.getSheet(xst).getWritableCell(Ro-1, Col-1);	  
	  //Number
	    if (cell.getType() == CellType.NUMBER)
		{
	    	Number n = (Number) cell;
			n.setValue(Integer.parseInt(val));
		}	  
	  //Label
	  if (cell.getType() == CellType.LABEL){
			 Label lc = (Label) cell;
			 lc.setString(val);
		}	  
	  
	  w2.write();
	  w2.close();
   }
      
   public void SetValue(String val, String Col) throws BiffException, IOException, WriteException{//Get Sheet + Set Row
		  Workbook w1 = Workbook.getWorkbook(new File(XLS));
		  WritableWorkbook w2 = Workbook.createWorkbook(new File(XLS), w1);
		  int R ;
		  Sheet sheet = w1.getSheet(XST);
		  int cols=sheet.getColumns();
		  for(R = 0; R<cols; R++){
			  Cell cellx = sheet.getCell(R, 0);
			  if(cellx.getContents().equalsIgnoreCase(Col)){
				break;
			  }
		  }
		  
		  Cell cell = w2.getSheet(XST).getWritableCell(CurrentRow, R-1);	  
		  //Number
		    if (cell.getType() == CellType.NUMBER)
			{
		    	Number n = (Number) cell;
				n.setValue(Integer.parseInt(val));
			}	  
		  //Label
		    if (cell.getType() == CellType.LABEL){
		    		Label lc = (Label) cell;
					lc.setString(val);
			}	  
  
		  w2.write();
		  w2.close(); 
   }
   
   public void AddParameter(String val, String Col) throws BiffException, IOException, WriteException{
	      Workbook w1 = Workbook.getWorkbook(new File(XLS));
		  WritableWorkbook w2 = Workbook.createWorkbook(new File(XLS), w1);
		  
		  Cell cell = w2.getSheet(XST).getWritableCell(CurrentRow, 0);	  
		  //Number
		    if (cell.getType() == CellType.NUMBER)
			{
		    	Number n = (Number) cell;
				n.setValue(Integer.parseInt(val));
			}	  
		  //Label
		    if (cell.getType() == CellType.LABEL){
		    		Label lc = (Label) cell;
					lc.setString(val);
			}	  

		  w2.write();
		  w2.close(); 
	 
   }
   
   public void AddParameter( String xst, String col, String val) throws BiffException, IOException, WriteException{
	      Workbook w1 = Workbook.getWorkbook(new File(XLS));
		  WritableWorkbook w2 = Workbook.createWorkbook(new File(XLS), w1);
		  
		  Cell cell = w2.getSheet(xst).getWritableCell(CurrentRow, 0);	  
		  //Number
		    if (cell.getType() == CellType.NUMBER)
			{
		    	Number n = (Number) cell;
				n.setValue(Integer.parseInt(val));
			}	  
		  //Label
		    if (cell.getType() == CellType.LABEL){
		    		Label lc = (Label) cell;
					lc.setString(val);
			}	  

		  w2.write();
		  w2.close(); 
   }
     
}

