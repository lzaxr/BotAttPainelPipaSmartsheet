 /*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package SmartSheet;
import TratarRelatorios.TextFile;
import com.smartsheet.api.*;
import com.smartsheet.api.models.*;
import java.io.IOException;
import java.util.*;
/**
 *
 * @author vitor.heser
 */
public class OSEmAberto {
    public String accessToken = "fg18ja86io8n4brrkz54tt9lt6";
    public long idPlanilha;
    
    public OSEmAberto(long id){
        this.idPlanilha = id;
    }
    
    public void ApagarLinhas() throws SmartsheetException{
        //INICIANDO CLIENT
        System.out.println("ENTRAMOS");
        Smartsheet smartsheet = SmartsheetFactory.createDefaultClient(accessToken);

        //CARREGANDO PLANILHA
        Sheet sheet = smartsheet.sheetResources().getSheet(this.idPlanilha,null,null,null,null,null,null,null);
        
        //Contando Linhas e Colunas
        List<Column> column = sheet.getColumns();
        System.out.println("Contando colunas Resultado = "+column.size());
        List<Row> row = sheet.getRows();
        System.out.println("Contando Linhas Resultado = "+row.size());
        
        //Para MONTAR ARRAY DE ID'S DE LINHAS PARA EXCLUIR
        List<Long> rowIds=new ArrayList<Long>();
        System.out.println("A Planilha tem " + sheet.getTotalRowCount()+ " Linhas");
        
        //MONTANDO ARRAY DE ID'S DE LINHAS PARA EXCLUIR
        /*VARIAVEL PARA EXCLUIR LINHAS DE */int qtdLinhaExcluir = 300;
        int t = 0;
        for(int i =0;i<row.size();i++){
            Long dado = row.get(i).getId().longValue();
            rowIds.add(new Long(dado));
            if(t==qtdLinhaExcluir){
                System.out.println("Apagando 100 Linhas");
                //EXCLUINDO LINHAS DE 100 EM 100
                smartsheet.sheetResources().rowResources().deleteRows(this.idPlanilha, new HashSet(rowIds), true);
                t=0;
            }
            t++;
        }
        //EXCLUINDO O RESTANTE DAS LINHAS
        if(t>0){
            System.out.println("Apagando o restante das linhas");
            smartsheet.sheetResources().rowResources().deleteRows(this.idPlanilha, new HashSet(rowIds), true);
        }
        //RECALCULO DE PLANILHA
    }
    
    public void InserirLinhas(String Caminho, Integer MaxDeLinhas) throws SmartsheetException, IOException{
//        System.out.println(Caminho);
        Smartsheet smartsheet = SmartsheetFactory.createDefaultClient(accessToken);
        Sheet sheet = smartsheet.sheetResources().getSheet(this.idPlanilha,null,null,null,null,null,null,null);
        List<Column> column = sheet.getColumns();
//        System.out.println(sheet.getName());
        
        List<Row> Linhas2 = new ArrayList<>();
        List<String> Colunas = new ArrayList<>();
        
        TextFile arquivo = new TextFile(Caminho);
        arquivo.openTextFile();
        //SUBSTITUIR PELO LEITOR DE CSV
        int prim = 0;
        int ind = 0;
        while (arquivo.next()){
            String linhasqq = "";
            String Linha = arquivo.readLine();
            System.out.println(Linha);
            
            String[] vDados = Linha.split(";");
            for(int i =0;i<16;i++){
                linhasqq = linhasqq + "";
                try{
                    String dado = vDados[i].toString().replaceAll("=", "").replaceAll("\"", "");
                    Colunas.add(dado);

                }catch(Exception e){
                    Colunas.add("");
                }
            }
            //COLUNAS = VDADOS
            //SUBSTITUIR PELO LEITOR DE CSV
            Row rowA = new Row();
            rowA.setCells(preencheCells(column, Colunas)).setToBottom(true);
            
            Linhas2.add(rowA);
            
            if(ind==MaxDeLinhas){
                List<Row> newRows = smartsheet.sheetResources().rowResources().addRows(this.idPlanilha,Linhas2);
                Linhas2.clear();
                ind=0;
            }
            ind++;
            Colunas.clear();
        }
        
        
        if(Linhas2.size()>0){
            List<Row> newRows = smartsheet.sheetResources().rowResources().addRows(this.idPlanilha,Linhas2);
            Linhas2.clear();
            ind=0;
        }
        System.out.println("old" + sheet.getTotalRowCount() + " rows from sheet: " + sheet.getName());
        

        
    }
    
    
    
    public List<Cell> preencheCells(List<Column> column, List<String> Colunas){
//        System.out.println(Colunas.get(0) + " - "+ Colunas.get(1) + " - "+ Colunas.get(2) + " - "+ Colunas.get(3) + " - "+ Colunas.get(4) + " - "+ Colunas.get(5) + " - "+ Colunas.get(6) + " - "+ Colunas.get(7) + " - "+ Colunas.get(8) + " - "+ Colunas.get(9) + " - "+ Colunas.get(10) + " - "+ Colunas.get(11) + " - "+ Colunas.get(12) + " - "+ Colunas.get(13) + " - "+ Colunas.get(14) + " - "+ Colunas.get(15));
        List<Cell> rowACells;
//        if(Colunas.get(11).length()>2){
//            System.out.println("Tipo1");
        String execate = Colunas.get(14);
        execate = execate.length()>5 ? execate : "";
//         System.out.println(
//                8+ " = "+ConvertData(Colunas.get(8))+" ||| "+//DATA
//                10 + " = "+ConvertData(Colunas.get(10))+" ||| "+//DATA
//                12 + " = "+ConvertData(execate)+" ||| "+
//                11 + " = "+ConvertData(Colunas.get(11))+" ||| "+
//                0 + " = "+Colunas.get(0)+" ||| "+
//                1 + " = "+Colunas.get(1)+" ||| "+
//                2 + " = "+Colunas.get(2)+" ||| "+
//                3 + " = "+Colunas.get(3)+" ||| "+
//                4 + " = "+Colunas.get(4)+" ||| "+
//                5 + " = "+Colunas.get(5)+" ||| "+
//                6 + " = "+Colunas.get(6)+" ||| "+
//                7 + " = "+Colunas.get(7)+" ||| "+
//                9 + " = "+Colunas.get(9)+" ||| "+
//                13 + " = "+Colunas.get(13)+" ||| "+
//                14 + " = "+Colunas.get(12)+" ||| "+
//                15 + " = "+Colunas.get(15));
//            );
         
            
            rowACells = Arrays.asList(
                new Cell(column.get(0).getId()).setValue(Colunas.get(0)),
                new Cell(column.get(1).getId()).setValue(Colunas.get(1)),
                new Cell(column.get(2).getId()).setValue(Colunas.get(2)),
                new Cell(column.get(3).getId()).setValue(Colunas.get(3)),
                new Cell(column.get(4).getId()).setValue(Colunas.get(4)),
                new Cell(column.get(5).getId()).setValue(Colunas.get(5)),
                new Cell(column.get(6).getId()).setValue(Colunas.get(6)),
                new Cell(column.get(7).getId()).setValue(Colunas.get(7)),
                new Cell(column.get(8).getId()).setValue(Colunas.get(8)),//DATA
                new Cell(column.get(9).getId()).setValue(Colunas.get(9)),
                new Cell(column.get(10).getId()).setValue(Colunas.get(10)),//DATA
                new Cell(column.get(11).getId()).setValue(Colunas.get(11)),
                new Cell(column.get(12).getId()).setValue(execate),//DATA
                new Cell(column.get(13).getId()).setValue(Colunas.get(13)),
                new Cell(column.get(14).getId()).setValue(Colunas.get(12)),
                new Cell(column.get(15).getId()).setValue(Colunas.get(15))
            );
//        }else{
           
//        }
        return rowACells;
    }
    public String ConvertData(String date){
//        System.out.println(date);
        date = date
        .replaceAll("-jan-", "-01-")
        .replaceAll("-fev-", "-02-")
        .replaceAll("-mar-", "-03-")
        .replaceAll("-abr-", "-04-")
        .replaceAll("-mai-", "-05-")
        .replaceAll("-jun-", "-06-")
        .replaceAll("-jul-", "-07-")
        .replaceAll("-ago-", "-08-")
        .replaceAll("-set-", "-09-")
        .replaceAll("-out-", "-10-")
        .replaceAll("-nov-", "-11-")
        .replaceAll("-dez-", "-12-");
        
//        01/01/2020
        Integer i = date.indexOf("-");
        try{
            if(i<=2){
                date = date.substring(6,10)+"-"+date.substring(3,5)+"-"+date.substring(0,2);
            }
        }catch(Exception e){
            date = "";
        }
//        System.out.println(date);
//        System.out.println("===================\n");
        return date;
    }
}
