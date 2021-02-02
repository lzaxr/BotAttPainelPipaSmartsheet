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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
/**
 *
 * @author vitor.heser
 */
public class Atualiz {
    public String accessToken = "fg18ja86io8n4brrkz54tt9lt6";
    public long idPlanilha= 5642387574810500L;
    
    public void ApagarLinhas() throws SmartsheetException{
        //INICIANDO CLIENT
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
   
    public void Atualizando(String Status) throws SmartsheetException, IOException{
        ApagarLinhas();
        Smartsheet smartsheet = SmartsheetFactory.createDefaultClient(accessToken);
        Sheet sheet = smartsheet.sheetResources().getSheet(this.idPlanilha,null,null,null,null,null,null,null);
        System.out.println(sheet.getName());
        
        List<Column> colunas = sheet.getColumns();
        List<Row> linhas = sheet.getRows();
        
        List<String> Colunas = new ArrayList<>();
        Colunas.clear();
        Colunas.add(Status);
        Colunas.add(DataHoraAtual());
        Colunas.add(ProximaAtualizacao());
        
        
        Row rowA = new Row();
        rowA.setCells(preencheCells(colunas, Colunas, linhas.size())).setToBottom(true);
        List<Row> Linhas2 = new ArrayList<>();
        Linhas2.add(rowA);
        List<Row> newRows = smartsheet.sheetResources().rowResources().addRows(this.idPlanilha,Linhas2);
        
    }
   
    public List<Cell> preencheCells(List<Column> column, List<String> Colunas, Integer ID){
        List<Cell> rowACells;
        rowACells = Arrays.asList(
                new Cell(column.get(0).getId()).setValue(ID),
                new Cell(column.get(1).getId()).setValue(Colunas.get(0)),
                new Cell(column.get(2).getId()).setValue(Colunas.get(1)),
                new Cell(column.get(3).getId()).setValue(Colunas.get(2))
        );
        return rowACells;
    }
    
    

    public String DataHoraAtual(){
        
        // data/hora atual
        LocalDateTime agora = LocalDateTime.now();

        // formatar a data
        DateTimeFormatter formatterData = DateTimeFormatter.ofPattern("dd/MM/uuuu");
        String dataFormatada = formatterData.format(agora);

        // formatar a hora
        DateTimeFormatter formatterHora = DateTimeFormatter.ofPattern("HH:mm:ss");
        String horaFormatada = formatterHora.format(agora);
        
        String datahoraatual = dataFormatada+ " "+ horaFormatada;
        
        return datahoraatual;
    }
    public String ProximaAtualizacao(){
        
        // data/hora atual
        LocalDateTime agora = LocalDateTime.now().plusMinutes(60);

        // formatar a data
        DateTimeFormatter formatterData = DateTimeFormatter.ofPattern("dd/MM/uuuu");
        String dataFormatada = formatterData.format(agora);

        // formatar a hora
        DateTimeFormatter formatterHora = DateTimeFormatter.ofPattern("HH:mm:ss");
        String horaFormatada = formatterHora.format(agora);
        
        String datahoraatual = dataFormatada+ " "+ horaFormatada;
        
        return datahoraatual;
    }
}
