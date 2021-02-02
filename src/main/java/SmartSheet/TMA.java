/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package SmartSheet;
import TratarRelatorios.TextFile;
import com.smartsheet.api.*;
import com.smartsheet.api.models.*;
import com.smartsheet.api.oauth.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
/**
 *
 * @author vitor.heser
 */
public class TMA {
    public String accessToken = "fg18ja86io8n4brrkz54tt9lt6";
    public long idPlanilha;
    
    public TMA(long id){
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
        /*VARIAVEL PARA EXCLUIR LINHAS DE */int qtdLinhaExcluir = 350;
        int t = 0;
        for(int i =0;i<row.size();i++){
            Long dado = row.get(i).getId().longValue();
            rowIds.add(new Long(dado));
            if(t==qtdLinhaExcluir){
                System.out.println("Apagando "+qtdLinhaExcluir+" Linhas");
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
        System.out.println(Caminho);
        Smartsheet smartsheet = SmartsheetFactory.createDefaultClient(accessToken);
        Sheet sheet = smartsheet.sheetResources().getSheet(this.idPlanilha,null,null,null,null,null,null,null);
        List<Column> column = sheet.getColumns();
        
        List<Row> Linhas2 = new ArrayList<>();
        List<Object> Colunas = new ArrayList<>();
        
        TextFile arquivo = new TextFile(Caminho);
        arquivo.openTextFile();
        //SUBSTITUIR PELO LEITOR DE CSV
        int prim = 0;
        int ind = 0;
        while (arquivo.next()){
            String Linha = arquivo.readLine();
            String[] vDados = Linha.split(";");
//            System.out.println(Linha);
            String ReclamaPipa = "";
            for(int i =0;i<=62;i++){
                try{
                    String dado = vDados[i].toString().replaceAll("=", "").replaceAll("\"", "").replaceAll("null", "");
                    switch(i){
                        case 0:
                            Integer dad = Integer.valueOf(dado);
                            Colunas.add(dad);
                            break;
                        case 7:
                            if(dado.contains("RECLAMA")){
                                ReclamaPipa = "RECLAMAÇÃO";
                            }else{
                                ReclamaPipa = "PIPA";
                            }
                        case 1:
                        case 2:
                        case 3:
                        case 4:
                        case 5:
                        case 6:
                        case 8:
                        case 13:
                        case 14:
                        case 15:
                        case 17:
                        case 9:
                        case 10:
                        case 11:
                        case 12:
                        case 53:
                        case 54:
                        case 55:
                        case 18:
                        case 19:
                        case 20:
                        case 21:
                        case 22:
                        case 23:
                        case 24:
                        case 25:
                        case 26:
                        case 27:
                        case 28:
                        case 29:
                        case 30:
                        case 31:
                            //Colunas.add(Integer.valueOf(dado));
                            //break;
                        case 32:
                        case 33:
                        case 34:
                        case 35:
                        case 36:
                        case 37:
                        case 38:
                        case 39:
                        case 40:
                        case 41:
                        case 42:
                        case 43:
                        case 44:
                        case 45:
                        case 46:
                        case 47:
                        case 48:
                        case 49:
                        case 50:
                        case 51:
                        case 52:
                        case 56:
                        case 57:
                        case 58:
                        case 59:
                        case 60:
                        case 61:
                            Colunas.add(dado);
                            break;
                        case 16:
                            Colunas.add(ReclamaPipa);
                            break;
                    }
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
            
            prim++;
        }
        
        if(Linhas2.size()>0){
            List<Row> newRows = smartsheet.sheetResources().rowResources().addRows(this.idPlanilha,Linhas2);
            Linhas2.clear();
            ind=0;
        }
        System.out.println("old" + sheet.getTotalRowCount() + " rows from sheet: " + sheet.getName());
    }
    
    
    
    public List<Cell> preencheCells(List<Column> column, List<Object> Colunas){
        List<Cell> rowACells;
        String linha = "";
//        System.out.println(Colunas.size());
//        for(int i = 0; i<Colunas.size();i++){
//            linha = linha +" | "+Colunas.get(i);
//        }
//        System.out.println(linha);
        rowACells = Arrays.asList(
                new Cell(column.get(0).getId()).setValue(Colunas.get(0)),
                new Cell(column.get(1).getId()).setValue(Colunas.get(1)),
                new Cell(column.get(2).getId()).setValue(Colunas.get(2)),
                new Cell(column.get(3).getId()).setValue(Colunas.get(3)),
                new Cell(column.get(4).getId()).setValue(Colunas.get(4)),
                new Cell(column.get(5).getId()).setValue(Colunas.get(5)),
                new Cell(column.get(6).getId()).setValue(Colunas.get(6)),
                new Cell(column.get(7).getId()).setValue(Colunas.get(7)),
                new Cell(column.get(8).getId()).setValue(Colunas.get(8)),
                new Cell(column.get(9).getId()).setValue(Colunas.get(9)),
                new Cell(column.get(10).getId()).setValue(Colunas.get(10)),
                new Cell(column.get(11).getId()).setValue(Colunas.get(11)),
                new Cell(column.get(12).getId()).setValue(Colunas.get(12)),
                new Cell(column.get(13).getId()).setValue(Colunas.get(13)),
                new Cell(column.get(14).getId()).setValue(Colunas.get(14)),
                new Cell(column.get(15).getId()).setValue(Colunas.get(15)),
                new Cell(column.get(16).getId()).setValue(Colunas.get(16)),
                new Cell(column.get(17).getId()).setValue(Colunas.get(17)),
                new Cell(column.get(18).getId()).setValue(Colunas.get(18)),
                new Cell(column.get(19).getId()).setValue(Colunas.get(19)),
                new Cell(column.get(20).getId()).setValue(Colunas.get(20)),
                new Cell(column.get(21).getId()).setValue(Colunas.get(21)),
                new Cell(column.get(22).getId()).setValue(Colunas.get(22)),
                new Cell(column.get(23).getId()).setValue(Colunas.get(23)),
                new Cell(column.get(24).getId()).setValue(Colunas.get(24)),
                new Cell(column.get(25).getId()).setValue(Colunas.get(25)),
                new Cell(column.get(26).getId()).setValue(Colunas.get(26)),
                new Cell(column.get(27).getId()).setValue(Colunas.get(27)),
                new Cell(column.get(28).getId()).setValue(Colunas.get(28)),
                new Cell(column.get(29).getId()).setValue(Colunas.get(29)),
                new Cell(column.get(30).getId()).setValue(Colunas.get(30)),
                new Cell(column.get(31).getId()).setValue(Colunas.get(31)),
                new Cell(column.get(32).getId()).setValue(Colunas.get(32)),
                new Cell(column.get(33).getId()).setValue(Colunas.get(33)),
                new Cell(column.get(34).getId()).setValue(Colunas.get(34)),
                new Cell(column.get(35).getId()).setValue(Colunas.get(35)),
                new Cell(column.get(36).getId()).setValue(Colunas.get(36)),
                new Cell(column.get(37).getId()).setValue(Colunas.get(37)),
                new Cell(column.get(38).getId()).setValue(Colunas.get(38)),
                new Cell(column.get(39).getId()).setValue(Colunas.get(39)),
                new Cell(column.get(40).getId()).setValue(Colunas.get(40)),
                new Cell(column.get(41).getId()).setValue(Colunas.get(41)),
                new Cell(column.get(42).getId()).setValue(Colunas.get(42)),
                new Cell(column.get(43).getId()).setValue(Colunas.get(43)),
                new Cell(column.get(44).getId()).setValue(Colunas.get(44)),
                new Cell(column.get(45).getId()).setValue(Colunas.get(45)),
                new Cell(column.get(46).getId()).setValue(Colunas.get(46)),
                new Cell(column.get(47).getId()).setValue(Colunas.get(47)),
                new Cell(column.get(48).getId()).setValue(Colunas.get(48)),
                new Cell(column.get(49).getId()).setValue(Colunas.get(49)),
                new Cell(column.get(50).getId()).setValue(Colunas.get(50)),
                new Cell(column.get(51).getId()).setValue(Colunas.get(51)),
                new Cell(column.get(52).getId()).setValue(Colunas.get(52)),
                new Cell(column.get(53).getId()).setValue(Colunas.get(53)),
                new Cell(column.get(54).getId()).setValue(Colunas.get(54)),
                new Cell(column.get(55).getId()).setValue(Colunas.get(55)),
                new Cell(column.get(56).getId()).setValue(Colunas.get(56)),
                new Cell(column.get(57).getId()).setValue(Colunas.get(57)),
                new Cell(column.get(58).getId()).setValue(Colunas.get(58)),
                new Cell(column.get(59).getId()).setValue(Colunas.get(59)),
                new Cell(column.get(60).getId()).setValue(Colunas.get(60)),
                new Cell(column.get(61).getId()).setValue(Colunas.get(61))
        );
        return rowACells;
    }
}
