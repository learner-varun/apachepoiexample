import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Driver {
   private static WebDriver seleniumWebDriver;
    public static void initializeDriver()
    {
        WebDriverManager.chromedriver().version("2.40").setup();
seleniumWebDriver = new ChromeDriver();
    }
    public static WebDriver getDriver()
    {
        return seleniumWebDriver;
    }
}
