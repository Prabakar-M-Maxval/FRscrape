/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package frscrape;

import java.io.File;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

/**
 *
 * @author prabakar
 */
public class DriverSetup {

    protected static WebDriver sfDriver;
    public static String baseUrl;

    @BeforeClass
    public static void setUp() {
        GetDriver(2);
        sfDriver.manage().deleteAllCookies();
        sfDriver.manage().window().maximize();
        sfDriver.manage().timeouts().implicitlyWait(90, TimeUnit.SECONDS);
        
    }

    public DriverSetup navigateToSearchPage() {
        baseUrl = "https://bases-brevets.inpi.fr/en/recherche-avanceeen.html";
        sfDriver.navigate().to(baseUrl);
        return new DriverSetup();
    }

    public static void GetDriver(int browserId) {
        {
            switch (browserId) {
                case 1:
                    // For Firefox Driver
                    sfDriver = new FirefoxDriver();
                    break;

                case 2:
                    // For Chrome Driver
                   File file = new File("src/main/resources/chromedriver.exe");
                    String absolutePath = file.getAbsolutePath();
                    ChromeOptions options = new ChromeOptions();
                    options.addArguments("chrome.switches", "--disable-extensions");
                    options.addArguments("--disable-notifications");
                    System.setProperty("webdriver.chrome.driver",absolutePath);
                   
                     // System.setProperty("webdriver.chrome.sfDriver",
                     //       "D:\\NetBeansProjects\\Max-IDS4.0\\src\\main\\resources\\chromedriver.exe");
                    
                    
                    sfDriver = new ChromeDriver(options);
                    break;

                case 3:
                    // For IE Driver
                     File file1 = new File ("src/main/resources/IEDriverServer.exe");
                    String absolutePath1 = file1.getAbsolutePath();
                    System.setProperty("webdriver.ie.driver",
                            absolutePath1);
                    sfDriver = new InternetExplorerDriver();
            }
        }
    }

    @AfterClass
    public static void tearDown() {
        sfDriver.quit();
    }
}
