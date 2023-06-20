using BotRpaChallenge;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

class Program
{
    static readonly string downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
    static readonly string filePath = Path.Combine(downloadsFolderPath, "challenge.xlsx");
    static void Main(string[] args)
    {
        IWebDriver driver = new ChromeDriver
        {
            Url = "https://www.rpachallenge.com"
        };
        string botaoBaixarExcel = "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a";
        bool fileExists = File.Exists(filePath);
        if (!fileExists)
        {
            driver.FindElement(By.XPath(botaoBaixarExcel)).Click();
            Thread.Sleep(1000);
        }
        var listaPessoas = ExcelUtils.LerExcel(filePath);
        string botaoStart = "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button";
        driver.FindElement(By.XPath(botaoStart)).Click();

        InputarDadosPessoa(driver, listaPessoas);
        SalvarResultado(driver);

        Thread.Sleep(1000);
        driver.Quit();
    }

    private static void InputarDadosPessoa(IWebDriver driver, List<Pessoa> listaPessoas)
    {
        for (int i = 0; i < 10; i++)
        {
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelFirstName']")).SendKeys(listaPessoas[i].FirstName);
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelLastName']")).SendKeys(listaPessoas[i].LastName);
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelCompanyName']")).SendKeys(listaPessoas[i].CompanyName);
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelRole']")).SendKeys(listaPessoas[i].Role);
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelAddress']")).SendKeys(listaPessoas[i].Address);
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelEmail']")).SendKeys(listaPessoas[i].Email);
            driver.FindElement(By.XPath("//*[@ng-reflect-name='labelPhone']")).SendKeys(listaPessoas[i].PhoneNumber);
            driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input")).Click();
        }
    }
    private static void SalvarResultado(IWebDriver driver)
    {
        string conteudo = driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]")).Text;
        System.IO.File.WriteAllText(downloadsFolderPath+"\\resultado.txt", conteudo);
    }
}