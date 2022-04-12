/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Principal;

import MontarRelatorio.Email.EnviarEmailHtml;
import SmartSheet.Atualiz;
import SmartSheet.OSEmAberto;
import com.smartsheet.api.SmartsheetException;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.ArrayList;
import javax.swing.JFileChooser;
import org.apache.commons.mail.EmailException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *Coluna2
 * @author vitor.heser
 */
public class Home {
    public static void main(String args[]) throws InterruptedException, IOException, ParseException, FileNotFoundException, InvalidFormatException, EmailException, SQLException, ClassNotFoundException, SmartsheetException{
        
        
        
        EnviarEmailHtml EEH = new EnviarEmailHtml();
        SispcStatusBot SSB = new SispcStatusBot(6);
        SSB.Status("Iniciando");
        
        
        
        String nome1 = "sispc.prl";
        String PastaDownload = corrigirdiretorio(nome1);
        String relatoriofinal = "\\\\fsprl01\\prolagos\\Publico\\temp\\";
        System.out.println(EEH.BomDiaTardeNoite());
        System.out.println("=============================================================");
        System.out.println("==========================INICIANDO==========================");
        System.out.println("==========================="+EEH.BomDiaTardeNoite()+"===========================");
        System.out.println("=============================================================");
        SSB.Status("Trabalhando");
        Atualiz TesteAtualizando = new Atualiz();

        Runtime.getRuntime().exec("taskkill /f /im chromedriver.exe");  
        Runtime.getRuntime().exec("taskkill /f /im excel.exe");  
        ExtrairRelatorioTMA ERS = new ExtrairRelatorioTMA();
        ExtrairRelatorioEmAberto ERAB = new ExtrairRelatorioEmAberto();

        
        ArrayList<String> dados = new ArrayList<>();
        
        /*dados.add("16");
        dados.add("75");
        dados.add("80");
        dados.add("124");
        dados.add("220");
        dados.add("299");
        dados.add("371");
        dados.add("420");
        dados.add("754");
        dados.add("755");
        dados.add("832");*/
        dados.add("125004");
        dados.add("146003");
        dados.add("125001");
        dados.add("125003");
        dados.add("125005");
        dados.add("125007");
        dados.add("125006");
        dados.add("125002");

        Integer MaxDeLinhas = 300;
        String status ="";
        
//        //INICIO DO TESTE
        String CaminhoDestino = PastaDownload+"\\arquivo.csv";
       // OSEmAberto OSEA = new OSEmAberto(7240420351600516L);
       /* int i=0;
        while(i==0){
            try{
                OSEA.ApagarLinhas();
                i++;
            }catch(Exception e){
                i=0;
            }
        }
        OSEA.InserirLinhas(CaminhoDestino, MaxDeLinhas);
//        //FIM DO TESTE*/
        
        try{
            TesteAtualizando.Atualizando("Atualizando OPE0019");
            ERS.ExtrairSansys(PastaDownload, dados, MaxDeLinhas);
        }catch(Exception e){
            System.out.println("Falhou Ao Atualizar OPE0019 de novo!!!!!!!!!!!!!!!!");
            System.out.println(e);
            status = status.length()>0 ? status + " e OPE0019 ": "Erro ao atualizar o OPE0019";  //Erro ao atualizar o OPE0019
            SSB.Status("Falha","Falha ao atualizado o Relatorio OPE0019");
            TesteAtualizando.Atualizando(status);
        }
        /*try{
            TesteAtualizando.Atualizando("Atualizando OPE0002");
            ERAB.ExtrairSansys(PastaDownload, dados, MaxDeLinhas);
        }catch(Exception e){
            status = status.length()>0 ? status + " e OPE0002 ": "Erro ao atualizar o OPE0002";
            SSB.Status("Falha","Falha ao atualizado o Relatorio OPE0002");
            TesteAtualizando.Atualizando(status);
        }*/

        status = status.length()>0 ? status : "Atualizado";
        TesteAtualizando.Atualizando(status);
        
//        EEH.enviarEmail(PastaDownload,relatoriofinal);
        SSB.Status("Dormindo");
        
        
        
    }
    
    
    
    public static String corrigirdiretorio(String nome1){
        String PastaDownload = "C:\\Users\\"+nome1+"\\Downloads";
        
        File diretorio = new File(PastaDownload);
        if (!diretorio.exists()) { 
            PastaDownload = "D:\\Users\\"+nome1+"\\Downloads";
            System.out.println ("Definindo pasta de Download: " + PastaDownload);
        } else { 
            PastaDownload = "C:\\Users\\"+nome1+"\\Downloads";
            System.out.println ("Definindo pasta de Download: " + PastaDownload);
        }
        
        if (!diretorio.exists()) { 
            PastaDownload = "E:\\Users\\"+nome1+"\\Downloads";
            System.out.println ("Definindo pasta de Download: " + PastaDownload);
        }
        
        return PastaDownload;
    }
    
    public static File escolherArquivos(){ 
        File arquivos = null; 
        JFileChooser fc = new JFileChooser(); 
        fc.setDialogTitle("Escolha o(s) arquivo(s)..."); 
        fc.setDialogType(JFileChooser.OPEN_DIALOG); 
        fc.setApproveButtonText("OK"); 
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY); 
        fc.setMultiSelectionEnabled(false); 
        int resultado = fc.showOpenDialog(fc); 
        if (resultado == JFileChooser.CANCEL_OPTION){ 
            System.exit(1); 
        } 
        arquivos = fc.getSelectedFile(); 
        return arquivos; 
    }
    
}
