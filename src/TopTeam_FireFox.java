import java.io.FileInputStream;
import java.io.FileNotFoundException;

import javax.swing.JOptionPane;

import jxl.Sheet;
import jxl.Workbook;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TopTeam_FireFox
{
	static WebDriver driver = null;
	static Sheet s;
	static FileInputStream fi = null;

	public static void main(String[] args) throws Throwable
	{
		driver = new FirefoxDriver();
		driver.get("http://topteam.quintiles.net/");

		String a1 = new String("EDTUSERNAME");
		String a2 = new String("BTNOKDLGPROJECTSELECTION");
		String a3 = new String("mniRepositoryanchorId");

		try
		{
			fi = new FileInputStream("C:\\Selenium\\TopTeam_ConfigFile.xls");
		}catch(FileNotFoundException e)
		{
			JOptionPane.showMessageDialog(null,"TopTeam Configuration File is not available in the given path.... Please check and try again","ERROR",JOptionPane.ERROR_MESSAGE);
			driver.quit();
			System.exit(1);
		}  

		Workbook W = Workbook.getWorkbook(fi);
		s = W.getSheet(0); 
		String Username = s.getCell(1,0).getContents();//1 refers column and 0 refers row

		clickandwaitbyid(a1);

		driver.findElement(By.id("EDTUSERNAME")).clear();

		driver.findElement(By.id("EDTUSERNAME")).sendKeys(Username);
		String password= s.getCell(1, 1).getContents();
		driver.findElement(By.id("EDTPASSWORD")).sendKeys(password);
		WebElement dropDownListBox1 = driver.findElement(By.id("CMBTIMEZONE"));
		Select clickThis1 = new Select(dropDownListBox1);
		clickThis1.selectByVisibleText("(GMT+05:30) India Standard Time");
		driver.findElement(By.id("BTNSUBMIT")).click();

		try
		{
			boolean Password1=driver.findElement(By.id("LBLMSGFRAMFEEDBACKDLG")).isDisplayed();

			if(Password1)
			{
				JOptionPane.showMessageDialog(null,"Sorry, your username and password are incorrect, Please correct in the config excel and try again","ERROR",JOptionPane.ERROR_MESSAGE);
				driver.quit();
				System.exit(1);
			}
		}catch(NoSuchElementException c1)
		{
			clickandwaitbyid(a2);

			String Project = s.getCell(1,3).getContents();

			driver.findElement(By.xpath("//span[contains(text(),'" + Project + "')]")).click();

			driver.findElement(By.id("BTNOKDLGPROJECTSELECTION")).click();

			clickandwaitbyid(a3);

			driver.findElement(By.id("mniRepositoryanchorId")).click();

			driver.findElement(By.xpath("//a/img[@src='/ttmISAPI_Ora.dll/Files/imgLst_RecordTypes27.gif']/preceding-sibling::img[@src='/ttmISAPI_Ora.dll/Files/tree_plus_image.gif']")).click();
		}
		System.exit(0);
	}

	public static void clickandwaitbyid(String id1)
	{
		WebDriverWait wait = new WebDriverWait(driver, 120);
		wait.until(ExpectedConditions.elementToBeClickable(By.id(id1)));
	}
}
