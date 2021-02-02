/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Principal;

import SmartSheet.OSEmAberto;
import TratarRelatorios.ConverterPCSV;
import com.smartsheet.api.SmartsheetException;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

/**
 *
 * @author vitor.heser
 */
public class ExtrairRelatorioEmAberto {
    public String enderecoSansys = "http://osonline.aegea.com.br:5555/sansysos/efetuarLoginUsuario.wf";
    public String usuario = "vitor.heser";
    public String senha = "155Heser";
    public void ExtrairSansys(String PastaDownload, ArrayList<String> dados, Integer MaxDeLinhas) throws InterruptedException, IOException, ParseException, SmartsheetException{
        System.out.println("=============================================================");
        System.out.println("==========================ATUALIZAR==========================");
        System.out.println("===========================OPE0002===========================");
        System.out.println("=============================================================");
        
        System.setProperty("webdriver.chrome.driver", "C:/Selenium/chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
//        options.addArguments("headless");
        WebDriver browser = new ChromeDriver(options);
        Date data = new Date();
        Date data2 = new Date();
        
        SimpleDateFormat dataNomerelatorio = new SimpleDateFormat("ddMMyyyy_");
        SimpleDateFormat datahoje = new SimpleDateFormat("ddMMyyyy");
        
        int x = -30;
        Calendar cal = GregorianCalendar.getInstance();
        cal.add(Calendar.DAY_OF_YEAR, x);
        Date datinGregorian = cal.getTime();
        
        String DataInicio = datahoje.format(datinGregorian);
        String DataHoje =  datahoje.format(data);
        
        System.out.println(DataInicio);
        System.out.println(DataHoje);
        
        try{
            VerificarDownload(1, PastaDownload,"OPE0002_").delete();
        }catch(Exception e){}
        Integer QtdItensPastaDownload = esperarDownload(PastaDownload);
        
        browser.get(enderecoSansys);
        
        WebDriverWait wait = (new WebDriverWait(browser, 60));

        WebElement CxUser = wait.until(ExpectedConditions.presenceOfElementLocated(By.name("nmLogin")));
        wait.until(ExpectedConditions.visibilityOf(CxUser));
        CxUser.sendKeys(usuario);

        WebElement CxSenha = browser.findElement(By.name("nmSenha"));
        wait.until(ExpectedConditions.visibilityOf(CxSenha));
        CxSenha.sendKeys(senha);
        
        WebElement Entrar = browser.findElement(By.name("_eventId_btLogar"));
        Entrar.click();

        WebElement tabOperacional = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-gen243")));
        tabOperacional.click();
        
        
        WebElement relatorio = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/div[1]/li[3]/div[1]/a[1]/span[1]")));
        relatorio.click();
        
        WebElement ListarOrdemDeServico = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/div[1]/li[3]/ul[1]/li[2]/div[1]/a[1]/span[1]")));
        ListarOrdemDeServico.click();
        
        WebElement Servicos = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1063"))); //By.id("ext-comp-1062")
        
        
        
        for(int i = 0; i<dados.size(); i++){
            //System.out.println("Chegou até aqui: " + dados.get(i) );
            Thread.sleep(800);
            Servicos.sendKeys(dados.get(i));
            Thread.sleep(1000);
            WebElement proximacombo = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1067"))); //
            proximacombo.click();
            Thread.sleep(1000);
        }
        
        WebElement execDe = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1051"))); //By.id("ext-comp-1050")
        for(int i=0; i<11;i++){ execDe.sendKeys(Keys.ARROW_LEFT); }
        execDe.sendKeys(DataInicio);
        
        WebElement execAte = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1052")));//By.id("ext-comp-1051")
        for(int i=0; i<11;i++){ execAte.sendKeys(Keys.ARROW_LEFT); }
        execAte.sendKeys(DataHoje);
        
        WebElement Pendente = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-gen634"))); //(By.id("ext-gen631")
        Pendente.click();
        
        WebElement Programado = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-gen649"))); // By.id("ext-gen646")
        Programado.click();
        
        WebElement excel2 = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-gen937"))); //By.id("ext-gen901")
        excel2.click();
        
        WebElement Gerar = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-gen271"))); 
        Gerar.click();
        
        String NomeRelatorio =  dataNomerelatorio.format(data2);
        
        while(QtdItensPastaDownload == esperarDownload(PastaDownload)){
            Thread.sleep(3000);
        }
        Thread.sleep(10000);
        browser.close();
        
        String CaminhoDestino = PastaDownload+"\\arquivo.csv";
        
        File arquivoOrigem = VerificarDownload(2, PastaDownload, "OPE0002_"+NomeRelatorio);

        ConverterPCSV CPCSV = new ConverterPCSV();
        CPCSV.TratarRelatorio(arquivoOrigem, CaminhoDestino);
        //377406209976196
        //7240420351600516
        OSEmAberto OSEA = new OSEmAberto(7240420351600516L);
        int i=0;
        while(i==0){
            try{
                OSEA.ApagarLinhas();
                i++;
            }catch(Exception e){
                i=0;
            }
        }
        OSEA.InserirLinhas(CaminhoDestino, MaxDeLinhas);
        
        arquivoOrigem.delete();
        
        
    }
    
    public File VerificarDownload(Integer idVerificador, String Caminho, String Contem){
        File arquivo=null;
        String NomeArq = null;
        File LocalDoArquivo = new File(Caminho);
        File[] listArquivos = LocalDoArquivo.listFiles();
        String PathArq = LocalDoArquivo.getPath();

        for (int i = 0; i < listArquivos.length; i++) {
            File file = listArquivos[i];
            String Arquivo = file.getName();

            if(Arquivo.contains(Contem) && Arquivo.contains(".csv") || Arquivo.contains(Contem) && Arquivo.contains(".xlsx")){
                NomeArq = PathArq+"\\"+Arquivo;
                arquivo = new File(NomeArq);

            }
        }
       
        return arquivo;
    }
    
    public Integer esperarDownload(String Caminho){
        File LocalDoArquivo = new File(Caminho);
        File[] listArquivos = LocalDoArquivo.listFiles();
        Integer valor = listArquivos.length;
        return valor;
    }
}
