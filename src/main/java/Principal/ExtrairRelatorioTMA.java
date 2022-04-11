
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Principal;


import SmartSheet.TMA;
import TratarRelatorios.TratarTMA;
import com.smartsheet.api.SmartsheetException;
import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
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
public class ExtrairRelatorioTMA {
    public String enderecoSansys = "http://osonline.aegea.com.br/sansysos/efetuarLoginUsuario.wf";
    public String usuario = "vitor.heser";
    public String senha = "155Heser";
    public void ExtrairSansys(String PastaDownload, ArrayList<String> dados, Integer MaxDeLinhas) throws InterruptedException, IOException, ParseException, SmartsheetException, ClassNotFoundException, SQLException{
        System.out.println("=============================================================");
        System.out.println("==========================ATUALIZAR==========================");
        System.out.println("===========================OPE0019===========================");
        System.out.println("=============================================================");
        System.setProperty("webdriver.chrome.driver", "C:/Selenium/chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        
        WebDriver browser = new ChromeDriver(options);
        Date data = new Date();
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
            File arqui = new File(PastaDownload+"\\arquivo.csv");
            arqui.delete();
            VerificarDownload(1, PastaDownload).delete();
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

        WebElement TMACompleto = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/div[1]/li[3]/ul[1]/li[17]/div[1]/a[1]/span[1]")));
        TMACompleto.click();

        WebElement Servicos = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1052")));


        for(int i = 0; i<dados.size(); i++){
            Servicos.sendKeys(dados.get(i));
            Thread.sleep(1000);
            WebElement proximacombo = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1057")));
            proximacombo.click();
            Thread.sleep(1000);
        }

        WebElement execDe = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1038")));
        for(int i=0; i<11;i++){ execDe.sendKeys(Keys.ARROW_LEFT); }
        execDe.sendKeys(DataInicio);

        WebElement execAte = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-comp-1039")));
        for(int i=0; i<11;i++){ execAte.sendKeys(Keys.ARROW_LEFT); }
        execAte.sendKeys(DataHoje);
        
        //System.exit(0);
        
        WebElement Gerar = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ext-gen271")));
        Gerar.click();

        //System.exit(0);  // Aqui tinha feito um die
        System.out.println(QtdItensPastaDownload+" == "+esperarDownload(PastaDownload));
        while(QtdItensPastaDownload == esperarDownload(PastaDownload)){
            Thread.sleep(3000);
        }

        Thread.sleep(10000);
        browser.close();

        String CaminhoDestino = PastaDownload+"\\arquivo.csv";
        File arquivoOrigem = VerificarDownload(2, PastaDownload);
        File arquivoDestino = new File(CaminhoDestino);
        arquivoOrigem.renameTo(arquivoDestino);


        TratarTMA TrataTMA = new TratarTMA();
        TrataTMA.InserirLinhas(CaminhoDestino);
        TrataTMA.PlotarCSV(CaminhoDestino);
        
        TMA SS = new TMA(377406209976196L);
        int i=0;
        while(i==0){
            try{
                SS.ApagarLinhas();
                i++;
            }catch(Exception e){
                i=0;
            }
        }
        SS.InserirLinhas(CaminhoDestino, MaxDeLinhas);  
        System.out.println("Atualizado o OPE019 Sim!!!!!");
    }
    
    public File VerificarDownload(Integer idVerificador, String Caminho){
        File arquivo=null;
        String NomeArq = null;
        File LocalDoArquivo = new File(Caminho);
        File[] listArquivos = LocalDoArquivo.listFiles();
        String PathArq = LocalDoArquivo.getPath();
        String Contem = "OPE0019";
        for (int i = 0; i < listArquivos.length; i++) {
            File file = listArquivos[i];
            String Arquivo = file.getName();

            if(Arquivo.contains(Contem) && Arquivo.contains(".csv") || idVerificador==1 && Arquivo.contains("crdownload") ){
                NomeArq = PathArq+"\\"+Arquivo;
                arquivo = new File(NomeArq);
            }
        }
       
        return arquivo;
    }
    
    public Integer esperarDownload(String Caminho){
        Integer menos = 0;
        File LocalDoArquivo = new File(Caminho);
        File[] listArquivos = LocalDoArquivo.listFiles();
        for(File Arquivo : listArquivos){
            if(Arquivo.getName().contains("crdownloads")){
                menos=menos+1;
            }
        }
        Integer valor = (listArquivos.length-menos);
        return valor;
    }
}
