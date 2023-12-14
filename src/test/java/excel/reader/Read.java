package excel.reader;


	
	
	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.IOException;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class Read {
		
		public static FileInputStream fis;
		XSSFWorkbook wb ;
		XSSFSheet sheet;
		
		
		
		public void readdata() {
		
		 try {
			fis = new FileInputStream("C:\\Users\\santh\\OneDrive\\Desktop\\Automation\\w2a\\src\\testproperties\\tesdata.xlsx");
			try {
				wb = new XSSFWorkbook(fis);
				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
		 XSSFSheet sheetname = wb.getSheet("Sheet1");
		 int rownum =sheetname.getLastRowNum();
		 
		 
		 for (int i= 0 ; i<=rownum ; i++) {
			 
			 XSSFRow row =sheetname.getRow(i);
			 XSSFCell cell = row.getCell(i);
			 String value = cell.getStringCellValue();
			 System.out.println(value);
			 
			 
			 
		 }
		 
		 
		 
		}
		
		
		public static void main(String[] args) {
			
			Read  ex = new Read ();
			ex.readdata();
			
		}

	}







