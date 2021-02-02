/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package MontarRelatorio.Email;


import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import org.apache.commons.io.FileUtils;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.EmailException;
import org.apache.commons.mail.HtmlEmail;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
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
public class EnviarEmailHtml {
    public ArrayList<String> Emails = new ArrayList<>();
    public ArrayList<String> Cc = new ArrayList<>();
    public String Nome;
    String servidor = "correio.level3br.com";
    int porta = 587;
    String usuario = "sispc.prl@prolagos.com.br";
    String senha = "c7j3O07t@19";
    String remetente = "planejamento.prl@prolagos.com.br";
    String [] destinatario;
    String [] destinatarioCopia;
    
    Date date = new Date();
    SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
    String hoje  = formato.format(date);
   
    public void enviarEmail(
        String pastadestino,
        String pastadados
        ) throws EmailException, IOException, InterruptedException, SQLException, ClassNotFoundException, ParseException{
        
        //WEBDRIVER=====================================================================================
        System.setProperty("webdriver.chrome.driver", "C:/Selenium/chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("headless");
        options.addArguments("window-size=1000x2700");
        WebDriver browser = new ChromeDriver(options);
        browser.get("https://app.smartsheet.com/sheets/WFwRM5q9c57QxpcV3G53cXxWxmXrcj6X995xGqg1?view=grid");
        WebDriverWait wait = (new WebDriverWait(browser, 60));
        
        
        WebElement user = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("loginEmail")));
        wait.until(ExpectedConditions.visibilityOf(user));
        user.sendKeys("vitor.heser@prolagos.com.br");
        
        WebElement Salvar = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("formControl")));
        wait.until(ExpectedConditions.visibilityOf(Salvar));
        Salvar.click();
        
        WebElement pass = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("loginPassword")));
        wait.until(ExpectedConditions.visibilityOf(pass));
        pass.sendKeys("Prolago12");
        
        Salvar = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("formControl")));
        wait.until(ExpectedConditions.visibilityOf(Salvar));
        Salvar.click();
        Thread.sleep(10000);
        
        //FIMWEBDRIVER=====================================================================================
        
        
        ModelarBaseEmails mbe = new ModelarBaseEmails();
        mbe.BaseAreas(pastadados);
        Emails = mbe.Departamento;
        Cc = mbe.Cc;
        
        String[] convertido = (String[]) Emails.toArray(new String[Emails.size()]);
        destinatario = convertido;
        
        String[] convertido2 = (String[]) Cc.toArray(new String[Cc.size()]);
        destinatarioCopia = convertido2;
        
//        System.out.println(destinatario);
        
        HtmlEmail email = new HtmlEmail();
        
        String assunto = "Relatório Pipas e Reclamações";
        
        email.setDebug(false);
        email.setHostName(servidor);
        email.setSmtpPort(porta);
        email.setStartTLSEnabled(true);
        email.setAuthentication(usuario, senha);
        System.out.println("    Realizada Conexão com e-mail");
        email.setFrom(remetente, "Planejamento - Relatorio de RH");
        email.setSubject(assunto);
        email.addTo(destinatario);
//        email.addBcc(destinatarioCopia);
        System.out.println("    Criado o e-mail");
        
        
        
        
        
        
        
                
        //WEBDRIVER=====================================================================================
        Thread.sleep(6000);
        Salvar = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html[1]/body[1]/div[1]/div[1]/div[4]/div[3]/div[4]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/button[1]")));
        wait.until(ExpectedConditions.visibilityOf(Salvar));
        try{
            Salvar.click();
        }catch(Exception e){}
        
        Thread.sleep(9000);
        browser.get("https://app.smartsheet.com/sheets/67Cg9pgr58p2MjGPr2R4jQQvr3MxgMcVWj8fw9M1?view=grid");
        Thread.sleep(9000);
        Salvar = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html[1]/body[1]/div[1]/div[1]/div[4]/div[3]/div[4]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/button[1]")));
        wait.until(ExpectedConditions.visibilityOf(Salvar));
        try{
            Salvar.click();
        }catch(Exception e){}
        
        Thread.sleep(6000);
        browser.get("https://app.smartsheet.com/b/publish?EQBCT=48e7b145699841f093e464fd28ee22a8");
        Thread.sleep(7000);
        TakesScreenshot screenshot = (TakesScreenshot) browser;
        File sourceFile = screenshot.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(sourceFile, new File(pastadestino+"\\captura.png"));
        browser.close();
        //FIMWEBDRIVER=====================================================================================
        
        
        
        
        String cid1 = email.embed(new File("\\\\Fsprl01\\prolagos\\RPA\\UTIL\\Prolagos.png"));
        String cid2 = email.embed(new File(pastadestino+"\\captura.png"));
        String txtHtml = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n" +
"xmlns:w=\"urn:schemas-microsoft-com:office:word\"\n" +
"xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\"\n" +
"xmlns=\"http://www.w3.org/TR/REC-html40\">\n" +
"\n" +
"<head>\n" +
"<meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\">\n" +
"<meta name=ProgId content=Word.Document>\n" +
"<meta name=Generator content=\"Microsoft Word 15\">\n" +
"<meta name=Originator content=\"Microsoft Word 15\">\n" +
"<link rel=File-List href=\"Planejamento_arquivos/filelist.xml\">\n" +
"<!--[if gte mso 9]><xml>\n" +
" <o:OfficeDocumentSettings>\n" +
"  <o:AllowPNG/>\n" +
" </o:OfficeDocumentSettings>\n" +
"</xml><![endif]-->\n" +
"<link rel=themeData href=\"Planejamento_arquivos/themedata.thmx\">\n" +
"<link rel=colorSchemeMapping href=\"Planejamento_arquivos/colorschememapping.xml\">\n" +
"<!--[if gte mso 9]><xml>\n" +
" <w:WordDocument>\n" +
"  <w:View>Normal</w:View>\n" +
"  <w:Zoom>0</w:Zoom>\n" +
"  <w:TrackMoves/>\n" +
"  <w:TrackFormatting/>\n" +
"  <w:HyphenationZone>21</w:HyphenationZone>\n" +
"  <w:PunctuationKerning/>\n" +
"  <w:ValidateAgainstSchemas/>\n" +
"  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>\n" +
"  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>\n" +
"  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>\n" +
"  <w:DoNotPromoteQF/>\n" +
"  <w:LidThemeOther>PT-BR</w:LidThemeOther>\n" +
"  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>\n" +
"  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>\n" +
"  <w:DoNotShadeFormData/>\n" +
"  <w:Compatibility>\n" +
"   <w:BreakWrappedTables/>\n" +
"   <w:SnapToGridInCell/>\n" +
"   <w:WrapTextWithPunct/>\n" +
"   <w:UseAsianBreakRules/>\n" +
"   <w:DontGrowAutofit/>\n" +
"   <w:SplitPgBreakAndParaMark/>\n" +
"   <w:EnableOpenTypeKerning/>\n" +
"   <w:DontFlipMirrorIndents/>\n" +
"   <w:OverrideTableStyleHps/>\n" +
"   <w:UseFELayout/>\n" +
"  </w:Compatibility>\n" +
"  <m:mathPr>\n" +
"   <m:mathFont m:val=\"Cambria Math\"/>\n" +
"   <m:brkBin m:val=\"before\"/>\n" +
"   <m:brkBinSub m:val=\"&#45;-\"/>\n" +
"   <m:smallFrac m:val=\"off\"/>\n" +
"   <m:dispDef/>\n" +
"   <m:lMargin m:val=\"0\"/>\n" +
"   <m:rMargin m:val=\"0\"/>\n" +
"   <m:defJc m:val=\"centerGroup\"/>\n" +
"   <m:wrapIndent m:val=\"1440\"/>\n" +
"   <m:intLim m:val=\"subSup\"/>\n" +
"   <m:naryLim m:val=\"undOvr\"/>\n" +
"  </m:mathPr></w:WordDocument>\n" +
"</xml><![endif]--><!--[if gte mso 9]><xml>\n" +
" <w:LatentStyles DefLockedState=\"false\" DefUnhideWhenUsed=\"false\"\n" +
"  DefSemiHidden=\"false\" DefQFormat=\"false\" DefPriority=\"99\"\n" +
"  LatentStyleCount=\"375\">\n" +
"  <w:LsdException Locked=\"false\" Priority=\"0\" QFormat=\"true\" Name=\"Normal\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" QFormat=\"true\" Name=\"heading 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 7\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 8\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"9\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"heading 9\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 6\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 7\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 8\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index 9\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 7\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 8\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"toc 9\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Normal Indent\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"footnote text\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"annotation text\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"header\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"footer\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"index heading\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"35\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"caption\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"table of figures\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"envelope address\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"envelope return\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"footnote reference\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"annotation reference\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"line number\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"page number\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"endnote reference\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"endnote text\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"table of authorities\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"macro\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"toa heading\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Bullet\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Number\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Bullet 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Bullet 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Bullet 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Bullet 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Number 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Number 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Number 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Number 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"10\" QFormat=\"true\" Name=\"Title\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Closing\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Signature\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"1\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"Default Paragraph Font\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text Indent\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Continue\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Continue 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Continue 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Continue 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"List Continue 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Message Header\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"11\" QFormat=\"true\" Name=\"Subtitle\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Salutation\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Date\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text First Indent\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text First Indent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Note Heading\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text Indent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Body Text Indent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Block Text\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Hyperlink\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"FollowedHyperlink\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"22\" QFormat=\"true\" Name=\"Strong\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"20\" QFormat=\"true\" Name=\"Emphasis\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Document Map\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Plain Text\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"E-mail Signature\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Top of Form\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Bottom of Form\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Normal (Web)\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Acronym\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Address\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Cite\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Code\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Definition\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Keyboard\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Preformatted\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Sample\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Typewriter\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"HTML Variable\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Normal Table\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"annotation subject\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"No List\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Outline List 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Outline List 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Outline List 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Simple 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Simple 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Simple 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Classic 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Classic 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Classic 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Classic 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Colorful 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Colorful 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Colorful 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Columns 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Columns 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Columns 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Columns 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Columns 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 6\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 7\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Grid 8\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 4\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 5\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 6\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 7\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table List 8\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table 3D effects 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table 3D effects 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table 3D effects 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Contemporary\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Elegant\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Professional\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Subtle 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Subtle 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Web 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Web 2\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Web 3\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Balloon Text\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" Name=\"Table Grid\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Table Theme\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" Name=\"Placeholder Text\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"1\" QFormat=\"true\" Name=\"No Spacing\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" Name=\"Revision\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"34\" QFormat=\"true\"\n" +
"   Name=\"List Paragraph\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"29\" QFormat=\"true\" Name=\"Quote\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"30\" QFormat=\"true\"\n" +
"   Name=\"Intense Quote\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"60\" Name=\"Light Shading Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"61\" Name=\"Light List Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"62\" Name=\"Light Grid Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"63\" Name=\"Medium Shading 1 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"64\" Name=\"Medium Shading 2 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"65\" Name=\"Medium List 1 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"66\" Name=\"Medium List 2 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"67\" Name=\"Medium Grid 1 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"68\" Name=\"Medium Grid 2 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"69\" Name=\"Medium Grid 3 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"70\" Name=\"Dark List Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"71\" Name=\"Colorful Shading Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"72\" Name=\"Colorful List Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"73\" Name=\"Colorful Grid Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"19\" QFormat=\"true\"\n" +
"   Name=\"Subtle Emphasis\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"21\" QFormat=\"true\"\n" +
"   Name=\"Intense Emphasis\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"31\" QFormat=\"true\"\n" +
"   Name=\"Subtle Reference\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"32\" QFormat=\"true\"\n" +
"   Name=\"Intense Reference\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"33\" QFormat=\"true\" Name=\"Book Title\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"37\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" Name=\"Bibliography\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"39\" SemiHidden=\"true\"\n" +
"   UnhideWhenUsed=\"true\" QFormat=\"true\" Name=\"TOC Heading\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"41\" Name=\"Plain Table 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"42\" Name=\"Plain Table 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"43\" Name=\"Plain Table 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"44\" Name=\"Plain Table 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"45\" Name=\"Plain Table 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"40\" Name=\"Grid Table Light\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\" Name=\"Grid Table 1 Light\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\" Name=\"Grid Table 6 Colorful\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\" Name=\"Grid Table 7 Colorful\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"Grid Table 1 Light Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"Grid Table 6 Colorful Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"Grid Table 7 Colorful Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"Grid Table 1 Light Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"Grid Table 6 Colorful Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"Grid Table 7 Colorful Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"Grid Table 1 Light Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"Grid Table 6 Colorful Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"Grid Table 7 Colorful Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"Grid Table 1 Light Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"Grid Table 6 Colorful Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"Grid Table 7 Colorful Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"Grid Table 1 Light Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"Grid Table 6 Colorful Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"Grid Table 7 Colorful Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"Grid Table 1 Light Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"Grid Table 2 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"Grid Table 3 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"Grid Table 4 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"Grid Table 5 Dark Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"Grid Table 6 Colorful Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"Grid Table 7 Colorful Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\" Name=\"List Table 1 Light\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\" Name=\"List Table 6 Colorful\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\" Name=\"List Table 7 Colorful\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"List Table 1 Light Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4 Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"List Table 6 Colorful Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"List Table 7 Colorful Accent 1\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"List Table 1 Light Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4 Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"List Table 6 Colorful Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"List Table 7 Colorful Accent 2\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"List Table 1 Light Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4 Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"List Table 6 Colorful Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"List Table 7 Colorful Accent 3\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"List Table 1 Light Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4 Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"List Table 6 Colorful Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"List Table 7 Colorful Accent 4\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"List Table 1 Light Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4 Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"List Table 6 Colorful Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"List Table 7 Colorful Accent 5\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"46\"\n" +
"   Name=\"List Table 1 Light Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"47\" Name=\"List Table 2 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"48\" Name=\"List Table 3 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"49\" Name=\"List Table 4 Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"50\" Name=\"List Table 5 Dark Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"51\"\n" +
"   Name=\"List Table 6 Colorful Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" Priority=\"52\"\n" +
"   Name=\"List Table 7 Colorful Accent 6\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Mention\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Smart Hyperlink\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Hashtag\"/>\n" +
"  <w:LsdException Locked=\"false\" SemiHidden=\"true\" UnhideWhenUsed=\"true\"\n" +
"   Name=\"Unresolved Mention\"/>\n" +
" </w:LatentStyles>\n" +
"</xml><![endif]-->\n" +
"<style>\n" +
"<!--\n" +
" /* Font Definitions */\n" +
" @font-face\n" +
"	{font-family:\"Cambria Math\";\n" +
"	panose-1:2 4 5 3 5 4 6 3 2 4;\n" +
"	mso-font-charset:0;\n" +
"	mso-generic-font-family:roman;\n" +
"	mso-font-pitch:variable;\n" +
"	mso-font-signature:3 0 0 0 1 0;}\n" +
"@font-face\n" +
"	{font-family:Calibri;\n" +
"	panose-1:2 15 5 2 2 2 4 3 2 4;\n" +
"	mso-font-charset:0;\n" +
"	mso-generic-font-family:swiss;\n" +
"	mso-font-pitch:variable;\n" +
"	mso-font-signature:-536859905 -1073732485 9 0 511 0;}\n" +
"@font-face\n" +
"	{font-family:\"Trebuchet MS\";\n" +
"	panose-1:2 11 6 3 2 2 2 2 2 4;\n" +
"	mso-font-charset:0;\n" +
"	mso-generic-font-family:swiss;\n" +
"	mso-font-pitch:variable;\n" +
"	mso-font-signature:1671 0 0 0 159 0;}\n" +
" /* Style Definitions */\n" +
" p.MsoNormal, li.MsoNormal, div.MsoNormal\n" +
"	{mso-style-unhide:no;\n" +
"	mso-style-qformat:yes;\n" +
"	mso-style-parent:\"\";\n" +
"	margin:0cm;\n" +
"	margin-bottom:.0001pt;\n" +
"	mso-pagination:widow-orphan;\n" +
"	font-size:11.0pt;\n" +
"	font-family:\"Calibri\",sans-serif;\n" +
"	mso-ascii-font-family:Calibri;\n" +
"	mso-ascii-theme-font:minor-latin;\n" +
"	mso-fareast-font-family:\"Times New Roman\";\n" +
"	mso-fareast-theme-font:minor-fareast;\n" +
"	mso-hansi-font-family:Calibri;\n" +
"	mso-hansi-theme-font:minor-latin;\n" +
"	mso-bidi-font-family:\"Times New Roman\";\n" +
"	mso-bidi-theme-font:minor-bidi;}\n" +
"a:link, span.MsoHyperlink\n" +
"	{mso-style-noshow:yes;\n" +
"	mso-style-priority:99;\n" +
"	mso-style-parent:\"\";\n" +
"	color:#0563C1;\n" +
"	text-decoration:underline;\n" +
"	text-underline:single;}\n" +
"a:visited, span.MsoHyperlinkFollowed\n" +
"	{mso-style-noshow:yes;\n" +
"	mso-style-priority:99;\n" +
"	color:#954F72;\n" +
"	mso-themecolor:followedhyperlink;\n" +
"	text-decoration:underline;\n" +
"	text-underline:single;}\n" +
".MsoChpDefault\n" +
"	{mso-style-type:export-only;\n" +
"	mso-default-props:yes;\n" +
"	font-size:11.0pt;\n" +
"	mso-ansi-font-size:11.0pt;\n" +
"	mso-bidi-font-size:11.0pt;\n" +
"	mso-ascii-font-family:Calibri;\n" +
"	mso-ascii-theme-font:minor-latin;\n" +
"	mso-fareast-font-family:\"Times New Roman\";\n" +
"	mso-fareast-theme-font:minor-fareast;\n" +
"	mso-hansi-font-family:Calibri;\n" +
"	mso-hansi-theme-font:minor-latin;\n" +
"	mso-bidi-font-family:\"Times New Roman\";\n" +
"	mso-bidi-theme-font:minor-bidi;}\n" +
"@page WordSection1\n" +
"	{size:612.0pt 792.0pt;\n" +
"	margin:70.85pt 3.0cm 70.85pt 3.0cm;\n" +
"	mso-header-margin:36.0pt;\n" +
"	mso-footer-margin:36.0pt;\n" +
"	mso-paper-source:0;}\n" +
"div.WordSection1\n" +
"	{page:WordSection1;}\n" +
"-->\n" +
"</style>\n" +
"<!--[if gte mso 10]>\n" +
"<style>\n" +
" /* Style Definitions */\n" +
" table.MsoNormalTable\n" +
"	{mso-style-name:\"Tabela normal\";\n" +
"	mso-tstyle-rowband-size:0;\n" +
"	mso-tstyle-colband-size:0;\n" +
"	mso-style-noshow:yes;\n" +
"	mso-style-priority:99;\n" +
"	mso-style-parent:\"\";\n" +
"	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;\n" +
"	mso-para-margin:0cm;\n" +
"	mso-para-margin-bottom:.0001pt;\n" +
"	mso-pagination:widow-orphan;\n" +
"	font-size:11.0pt;\n" +
"	font-family:\"Calibri\",sans-serif;\n" +
"	mso-ascii-font-family:Calibri;\n" +
"	mso-ascii-theme-font:minor-latin;\n" +
"	mso-hansi-font-family:Calibri;\n" +
"	mso-hansi-theme-font:minor-latin;\n" +
"	mso-bidi-font-family:\"Times New Roman\";\n" +
"	mso-bidi-theme-font:minor-bidi;}\n" +
"</style>\n" +
"<![endif]-->\n" +
"</head>\n" +
"\n" +
"<body lang=PT-BR link=\"#0563C1\" vlink=\"#954F72\" style='tab-interval:35.4pt'>\n" +
"\n" 
                
                +"<p style=\"line-height: 1.27; font-family: &quot;YACgEf1HUE0 0&quot;, _fb_, auto; font-size: 16px; text-transform: none; color: rgb(115, 115, 115);\" class=\"direction-ltr align-center para-style-body\"><strong style=\"rgb(115, 115, 115);\">"
                + "<em>"+BomDiaTardeNoite()+"</em></strong></p>"
                +"<p style=\"line-height: 1.27; font-family: &quot;YACgEf1HUE0 0&quot;, _fb_, auto; font-size: 14px; text-transform: none; color: rgb(115, 115, 115);\" class=\"direction-ltr align-center para-style-body\"><strong style=\"rgb(115, 115, 115);\">"
                + "Segue Abaixo o relatório de Pipas e Reclamações."
                +"<p style=\"line-height: 1.27; font-family: &quot;YACgEf1HUE0 0&quot;, _fb_, auto; font-size: 14px; text-transform: none; color: rgb(115, 115, 115);\" class=\"direction-ltr align-center para-style-body\"><strong style=\"rgb(115, 115, 115);\">"
                + "<img src=\"cid:" + cid2 + "\"  "
                + "width=\"1500\""
                + ">" 
                
                +"<p style=\"line-height: 1.27; font-family: &quot;YACgEf1HUE0 0&quot;, _fb_, auto; font-size: 14px; text-transform: none; color: rgb(115, 115, 115);\" class=\"direction-ltr align-center para-style-body\"><strong style=\"rgb(115, 115, 115);\">"
                + "<em>Atenciosamente,</em></strong></p>"
                +"<br>"
                +"<img src=\"cid:" + cid1 + "\"  width=\"127\" height=\"90\">"+
"<div class=WordSection1>\n" +
"\n" +
"<p class=MsoNormal><b><span style='font-size:10.0pt;font-family:\"Trebuchet MS\",sans-serif;\n" +
"color:#253F93'>Setor Planejamento<o:p></o:p></span></b></p>\n" +
"\n" +
"<p class=MsoNormal><span class=MsoHyperlink><span style='font-size:8.0pt;\n" +
"font-family:\"Trebuchet MS\",sans-serif'><a href=\"http://www.prolagos.com.br/\">http://www.prolagos.com.br</a></span></span><span\n" +
"class=MsoHyperlink><span style='mso-bidi-font-family:\"Times New Roman\"'><o:p></o:p></span></span></p>\n" +
"\n" +
"<p class=MsoNormal><o:p>&nbsp;</o:p></p>\n" +
"\n" +
"</div>\n" +
"\n" +
"</body>\n" +
"\n" +
"</html>" ;
 


        try{
            email.setHtmlMsg(txtHtml);
//            EmailAttachment anexo = new EmailAttachment();
//            anexo.setPath(Planilha);
//            anexo.setDisposition(EmailAttachment.ATTACHMENT);
//            anexo.setName("Relatorio de Veículos.xslx");
//            email.attach(anexo);
//
//             //envia o email
            email.send();
            System.out.println("    Email enviado!");
        }catch(Exception e){}
    
        

    }  
    
    public String BomDiaTardeNoite() throws ParseException{
        String DiaTardeNoite = "";
        Calendar dataatual = new GregorianCalendar();
        SimpleDateFormat formdata = new SimpleDateFormat("HH:mm:SS");
        String dataatual2 = formdata.format(dataatual.getTime());
        
        String t0 = "05:00:00";
        String t1 = "12:00:00";
        String t2 = "18:00:00";
        
        Date data0;
        Date data1;
        Date data2;
        Date Agora;
        
        data0 = (formdata.parse(t0));
        Calendar temp0 = new GregorianCalendar();
        temp0.setTime(data0);
        
        data1 = (formdata.parse(t1));
        Calendar temp1 = new GregorianCalendar();
        temp1.setTime(data1);
        
        data2 = (formdata.parse(t2));
        Calendar temp2 = new GregorianCalendar();
        temp2.setTime(data2);
        
        

        Agora = (formdata.parse(dataatual2));
        Calendar hoje = new GregorianCalendar();
        hoje.setTime(Agora);
        
        
        if(Agora.after(data0) && Agora.before(data1)){
            DiaTardeNoite = "Bom Dia";
        }else if(Agora.after(data1) && Agora.before(data2)){
            DiaTardeNoite = "Boa Tarde";
        }else if(Agora.before(data0) && Agora.after(data2)){
            DiaTardeNoite = "Boa Noite";
        }else{
            DiaTardeNoite = "Bom d";
        }
        
        
        
        return DiaTardeNoite;
    }
}
