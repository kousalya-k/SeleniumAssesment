import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class yahoo {
	public static String vSearch;
	public static int xlCols, xlRows;
	public static String xData[][];

	public static void main(String[] args) throws Exception {

		companyData();

	}
    
	
	
	//To read from the Excel Sheet
	public static void xlRead(String sPath) throws Exception {
		File myFile = new File(sPath);
		FileInputStream myStream = new FileInputStream(myFile);
		HSSFWorkbook myworkbook = new HSSFWorkbook(myStream);
		HSSFSheet mySheet = myworkbook.getSheetAt(0);
		xlRows = mySheet.getLastRowNum() + 1;
		xlCols = mySheet.getRow(0).getLastCellNum();
		xData = new String[xlRows][xlCols];
		for (int i = 0; i < xlRows; i++) {
			HSSFRow row = mySheet.getRow(i);
			for (short j = 0; j < xlCols; j++) {
				HSSFCell cell = row.getCell(j);
				String value = cellToString(cell);
				xData[i][j] = value;
				System.out.print("-" + xData[i][j]);
			}
			System.out.println();
		}
	}

	public static String cellToString(HSSFCell cell) {
		int type = cell.getCellType();
		Object result;
		switch (type) {
		case HSSFCell.CELL_TYPE_NUMERIC:
			result = cell.getNumericCellValue();
			break;
		case HSSFCell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			throw new RuntimeException("We cannot evaluate formula");
		case HSSFCell.CELL_TYPE_BLANK:
			result = "-";
		case HSSFCell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue();
		case HSSFCell.CELL_TYPE_ERROR:
			result = "This cell has some error";
		default:
			throw new RuntimeException("We do not support this cell type");
		}
		return result.toString();

	}
   
	
	
	//To write in the Excel Sheet
	public static void xlwrite(String xlpath1, String[][] xData) throws Exception {
		System.out.println("Inside XL Write");
		File myFile1 = new File(xlpath1);
		FileOutputStream fout = new FileOutputStream(myFile1);
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet mySheet1 = wb.createSheet("TestResults");
		for (int i = 0; i < xlRows; i++) {
			HSSFRow row1 = mySheet1.createRow(i);
			for (short j = 0; j < xlCols; j++) {
				HSSFCell cell1 = row1.createCell(j);
				cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell1.setCellValue(xData[i][j]);
			}
		}
		wb.write(fout);
		fout.flush();
		fout.close();
	}

	public static void companyData() throws Exception {
		
		//calling the Read method
		xlRead("C:\\Users\\kk98\\Documents\\Googlexls.xls");

		for (int i = 0; i < xlRows; i++) {
                
			
			   //To execute for flag of value Y
			    if (xData[i][1].equalsIgnoreCase("Y")) {
                
			    	
			    //Assigning value in excel sheet into search box
				vSearch = xData[i][0];
				System.setProperty("webdriver.chrome.driver", "C:\\Users\\kk98\\SeleniumJar\\chromedriver.exe");
				WebDriver driver = new ChromeDriver();
				driver.manage().window().maximize();
				driver.get("https://www.yahoo.com/");
				Thread.sleep(5000);
				driver.findElement(By.id("uh-search-box")).sendKeys(vSearch, Keys.ENTER);
				Thread.sleep(5000);
				String title = driver.getTitle();
				xData[i][2] = title;
				driver.close();
			}
		}
		
		//calling the write method
		xlwrite("C:\\Users\\kk98\\Documents\\Yahoo.xls", xData);
	}
}