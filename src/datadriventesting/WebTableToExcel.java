package datadriventesting;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class WebTableToExcel {

	public static void main(String[] args) throws IOException {
		
		System.setProperty("webdriver.chrome.driver", "D:\\Software\\Drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		
		driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");
		
		String path = ".\\dataFiles\\population.xlsx";
		XLUtility xlutil = new XLUtility(path);
		
		//Write headers in Excel Sheet
		xlutil.setCellData("Sheet1", 0, 0, "Country");
		xlutil.setCellData("Sheet1", 0, 1, "Region");
		xlutil.setCellData("Sheet1", 0, 2, "Population");
		xlutil.setCellData("Sheet1", 0, 3, "% of world");
		xlutil.setCellData("Sheet1", 0, 4, "Date");
		xlutil.setCellData("Sheet1", 0, 5, "Source");
		xlutil.setCellData("Sheet1", 0, 6, "Notes");
		
		//Capture table rows
		WebElement table = driver.findElement(By.xpath("//table[@class = 'wikitable sortable jquery-tablesorter']/tbody"));
		
		int rows = table.findElements(By.xpath("tr")).size();	//rows present in web table
		
		for(int r=1; r<=rows; r++)
		{
			String country = table.findElement(By.xpath("tr["+r+"]/td[1]")).getText();
			String region = table.findElement(By.xpath("tr["+r+"]/td[2]")).getText();
			String population = table.findElement(By.xpath("tr["+r+"]/td[3]")).getText();
			String perOfWorld = table.findElement(By.xpath("tr["+r+"]/td[4]")).getText();
			String date = table.findElement(By.xpath("tr["+r+"]/td[5]")).getText();
			String source = table.findElement(By.xpath("tr["+r+"]/td[6]")).getText();
			String notes = table.findElement(By.xpath("tr["+r+"]/td[7]")).getText();
			
			System.out.println(country+region+population+perOfWorld+date+source+notes);
			
			//Writing the data in Excel Sheet
			xlutil.setCellData("Sheet1", r, 0, country);
			xlutil.setCellData("Sheet1", r, 1, region);
			xlutil.setCellData("Sheet1", r, 2, population);
			xlutil.setCellData("Sheet1", r, 3, perOfWorld);
			xlutil.setCellData("Sheet1", r, 4, date);
			xlutil.setCellData("Sheet1", r, 5, source);
			xlutil.setCellData("Sheet1", r, 6, notes);
		}
		System.out.println("Web scrapping is done successfully...");
		driver.close();

	}

}
