/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package TratarRelatorios;
import com.smartsheet.api.*;
import com.smartsheet.api.models.*;
import com.smartsheet.api.oauth.*;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.util.*;
/**
 *
 * @author vitor.heser
 */
public class TratarTMA {

        public String consulta="SELECT \n"
            + "	count(*) as anoos,\n"
            + "	idosp,\n"
            + "	Anomes,\n"
            + "	matricula,\n"
            + "	ENDerecoosp,\n"
            + "	Bairro,\n"
            + "	servico,\n"
            + "	descricao,\n"
            + "\n"
            //+ " cdCognos\n"
            + "	Eqp,\n"
            + "	DataEmissao,\n"
            + "	DataProgramacao,\n"
            + "	DATAINICIO,\n"
            + "	DataConclusao,\n"
            + "	status,\n"
            + "	ATRASO,\n"
            + "	nomecor,\n"
            + "	codexecocor,\n"
            + "	motivoexecocor,\n"
            + "	idanomalia,\n"
            + "	descranomalia,\n"
            + "	tempopadrao,\n"
            + "	tempodesloc,\n"
            + "	tempodeslocsegundos,\n"
            + "	tempopreparo,\n"
            + "	tempopreparosegundos,\n"
            + "	tempoexecucao,\n"
            + "	tempoexecucaosegundos,\n"
            + "	tempopreenchimento,\n"
            + "	tempopreenchimentosegundos,\n"
            + "	tempoparalisacao,\n"
            + "	tempoparalisacaosegundos,\n"
            + "	tempoatendimento,\n"
            + "	DescrSetor,\n"
            + "	latitude,\n"
            + "	longitude,\n"
            + "	FotosMin,\n"
            + "	QtdeFotos,\n"
            + "	DAT_AGENDAMENTO,\n"
            + "	SF,\n"
            + "	latitudeExecucao,\n"
            + "	longitudeExecucao,\n"
            + "	ID_CONTRATO,\n"
            + "	CD_SUB_REGIAO,\n"
            + "	CD_REGIAO,\n"
            + "	idservico,\n"
            + "	idExecucaoServico,\n"
            + "	ServiçoPrincipal,\n"
            + "	idExecucaoServicoPrincipal,\n"
            + "	unidade,\n"
            + "	CD_SERVICO_SOLICITADO,\n"
            + "	DS_SERVICO_SOLICITADO,\n"
            + "	Executado,\n"
            + "	NM_TIPO_EXECUCAO,\n"
            + "	dtLimiteExecucao,\n"
            + "	dtServico,\n"
            + "	dtFechamento,\n"
            + "	funcionarios,\n"
            + "	DescrSetorSolicitante,\n"
            + "	TerceiroNomeEmpresa,\n"
            + "	TerceiroNomeFantasia,\n"
            + "	TerceiroCNPJ,\n"
            + "	MatrizdeServiços\n"
            + "\n"
            + "\n"
            + "FROM base_gsson_line_servicostmaemergencial\n"
            + "group by \n"
            + "	descricao,\n"
            + "	dataemissao,\n"
            + "	DataConclusao,\n"
            + "	status,\n"
            + "	eqp,\n"
            + "	atraso,\n"
            + "	codexecocor,\n"
            + "	sf,\n"
            + "	dtLimiteExecucao";
    
  /*  public Connection getConnection() throws ClassNotFoundException, SQLException{
        Connection connection=null;
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            connection = DriverManager.getConnection("jdbc:mysql://dskprl013862:3306/sispc?useTimezone=true&serverTimezone=UTC","sispcAPI", "Prol@go55i5PC@pi");
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
        return connection;
    }
    */
        public Connection getConnection() throws ClassNotFoundException, SQLException{
            Connection connection=null;
            try {
                Class.forName("com.mysql.cj.jdbc.Driver");
                connection = DriverManager.getConnection("jdbc:mysql://sispcprl01:3306/sispc?useTimezone=true&serverTimezone=UTC","sispcAPI", "Prol@go55i5PC@pi");
                System.out.println("CONECTADO");
            } catch (ClassNotFoundException | SQLException e) {
                e.printStackTrace();
                System.out.println("NÃO CONECTOU");
            }
            return connection;
        }
    public void InserirLinhas(String Caminho) throws SmartsheetException, IOException, ClassNotFoundException, SQLException{
        System.out.println("Entrou aqui, Juliano");
        List<Row> Linhas2 = new ArrayList<>();
        List<String> Colunas = new ArrayList<>();
        Connection connection = getConnection();
        
        String querry= "TRUNCATE base_gsson_line_servicostmaemergencial;";
        connection.prepareStatement(querry).executeUpdate();
        querry="";
                
        TextFile arquivo = new TextFile(Caminho);
        arquivo.openTextFile();
        //SUBSTITUIR PELO LEITOR DE CSV
        int prim = 0;
        int ident =0;
        String querryInicio="INSERT INTO base_gsson_line_servicostmaemergencial VALUES ";
        while (arquivo.next()){
            String Linha = arquivo.readLine().replaceAll("=", "");
            System.out.println(Linha);
            if(prim > 6){
                String[] vDados = Linha.split("\";\"");
                String ReclamaPipa = "";
                for(int i =0;i<62;i++){
                   try{
                        String dado = vDados[i].toString().replaceAll(";", ",").replaceAll("'", "");
                        switch(i){
                            case 10:
                            case 11:
                            case 12:
                            case 13:
                            
                            case 54:
                            case 55:
                            case 56:
                            case 38:
                               //System.out.println("=============================================");
                               //System.out.println(dado);
                                try{
                                    String Dia = dado.substring(0,2);
                                    String Mes = dado.substring(3,5);
                                    String Ano = dado.substring(6,10);
                                    //System.out.println("'"+Ano+"-"+Mes+"-"+Dia+"'");
                                    Colunas.add("'"+Ano+"-"+Mes+"-"+Dia+"'");
                                }catch(Exception e){
                                    Colunas.add("null");
                                }
                                break;
                            
                            case 7:
                                if(dado.contains("RECLAMA")){
                                    ReclamaPipa = "RECLAMAÇÃO";
                                }else{
                                    ReclamaPipa = "PIPA";
                                }
                            case 0:
                            case 1:
                            case 2:
                            case 3:
                            case 4:
                            case 5:
                            case 6:
                            case 8:
                            case 9:
                            case 14:
                            case 15:
                            case 16:
                            case 17:
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
                            case 32:
                            case 33:
                            case 34:
                            case 35:
                            case 36:
                            case 37:
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
                            case 53:
                            case 57:
                            case 58:
                            case 59:
                            case 60:
                            case 61:
                                Colunas.add(dado);
                                break;
                            case 18:
                                Colunas.add(ReclamaPipa);
                                break;
                        }
                    }catch(Exception e){
                        Colunas.add("NULL");
                    }
                }
                querry=  "('"+Colunas.get(0)+"','"+Colunas.get(1)+"','"+Colunas.get(2)+"','"+Colunas.get(3)+"','"+Colunas.get(4)
                        +"','"+Colunas.get(5)+"','"+Colunas.get(6)+"','"+Colunas.get(7)+"','"+Colunas.get(9)+"',"+Colunas.get(10)+","+Colunas.get(11)
                        +","+Colunas.get(12)+","+Colunas.get(13)+",'"+Colunas.get(14)+"','"+Colunas.get(15)+"','"+Colunas.get(16)+"','"+Colunas.get(17)+"','"+Colunas.get(18)
                        +"','"+Colunas.get(19)+"','"+Colunas.get(20)+"','"+Colunas.get(21)+"','"+Colunas.get(22)+"','"+Colunas.get(23)+"','"+Colunas.get(24)+"','"+Colunas.get(25)
                        +"','"+Colunas.get(26)+"','"+Colunas.get(27)+"','"+Colunas.get(28)+"','"+Colunas.get(29)+"','"+Colunas.get(30)+"','"+Colunas.get(31)+"','"+Colunas.get(32)
                        +"','"+Colunas.get(33)+"','"+Colunas.get(34)+"','"+Colunas.get(35)+"','"+Colunas.get(36)+"','"+Colunas.get(37)+"',"+Colunas.get(38)+",'"+Colunas.get(39)
                        +"','"+Colunas.get(40)+"','"+Colunas.get(41)+"','"+Colunas.get(42)+"','"+Colunas.get(43)+"','"+Colunas.get(44)+"','"+Colunas.get(45)+"','"+Colunas.get(46)
                        +"','"+Colunas.get(47)+"','"+Colunas.get(48)+"','"+Colunas.get(49)+"','"+Colunas.get(50)+"','"+Colunas.get(51)+"','"+Colunas.get(52)+"','"+Colunas.get(53)
                        +"',"+Colunas.get(54)+","+Colunas.get(55)+","+Colunas.get(56)+",'"+Colunas.get(57)+"','"+Colunas.get(58)+"','"+Colunas.get(59)+"','"+Colunas.get(60)+"','"+Colunas.get(61)+"',null)";
                System.out.println(Colunas.get(14));
                Colunas.clear();
                
                if(ident>0 ){
                    querry = querryInicio + querry+";";  
                    System.out.println(querry);
                    //System.out.println(querry.length());
                    
                    connection.prepareStatement(querry).executeUpdate();
                    
                    querry="";
                    ident=0;
                }
                
                ident++;
            }
            prim++;
        }
        /*if(ident>0){
            querry = querryInicio + querry+";";
            System.out.println(querry);
            connection.prepareStatement(querry).executeUpdate();
            //System.out.println(querry);
            querry="";
            ident=0;
        }*/
        
    }
    
    public void PlotarCSV(String Caminho) throws ClassNotFoundException, SQLException, IOException{
        
        Connection connection= getConnection();
        PreparedStatement ps=connection.prepareStatement(consulta);       
        ResultSet rs=ps.executeQuery();
        
        FileWriter relatorio = new FileWriter(new File(Caminho));
        BufferedWriter writer = new BufferedWriter(relatorio);
        
        while(rs.next()){
//            System.out.println("asd");
            String Linha =
            rs.getString("anoos")+";"+
            rs.getString("idosp")+";"+
            rs.getString("Anomes")+";"+
            rs.getString("matricula")+";"+
            rs.getString("ENDerecoosp")+";"+
            rs.getString("Bairro")+";"+
            rs.getString("servico")+";"+
            rs.getString("descricao")+";"+
            rs.getString("Eqp")+";"+
            rs.getString("DataEmissao")+";"+
            rs.getString("DataProgramacao")+";"+
            rs.getString("DATAINICIO")+";"+
            rs.getString("DataConclusao")+";"+
            rs.getString("status")+";"+
            rs.getString("ATRASO")+";"+
            rs.getString("nomecor")+";"+
            rs.getString("codexecocor")+";"+
            rs.getString("motivoexecocor")+";"+
            rs.getString("idanomalia")+";"+
            rs.getString("descranomalia")+";"+
            rs.getString("tempopadrao")+";"+
            rs.getString("tempodesloc")+";"+
            rs.getString("tempodeslocsegundos")+";"+
            rs.getString("tempopreparo")+";"+
            rs.getString("tempopreparosegundos")+";"+
            rs.getString("tempoexecucao")+";"+
            rs.getString("tempoexecucaosegundos")+";"+
            rs.getString("tempopreenchimento")+";"+
            rs.getString("tempopreenchimentosegundos")+";"+
            rs.getString("tempoparalisacao")+";"+
            rs.getString("tempoparalisacaosegundos")+";"+
            rs.getString("tempoatendimento")+";"+
            rs.getString("DescrSetor")+";"+
            rs.getString("latitude")+";"+
            rs.getString("longitude")+";"+
            rs.getString("FotosMin")+";"+
            rs.getString("QtdeFotos")+";"+
            rs.getString("DAT_AGENDAMENTO")+";"+
            rs.getString("SF")+";"+
            rs.getString("latitudeExecucao")+";"+
            rs.getString("longitudeExecucao")+";"+
            rs.getString("ID_CONTRATO")+";"+
            rs.getString("CD_SUB_REGIAO")+";"+
            rs.getString("CD_REGIAO")+";"+
            rs.getString("idservico")+";"+
            rs.getString("idExecucaoServico")+";"+
            rs.getString("ServiçoPrincipal")+";"+
            rs.getString("idExecucaoServicoPrincipal")+";"+
            rs.getString("unidade")+";"+
            rs.getString("CD_SERVICO_SOLICITADO")+";"+
            rs.getString("DS_SERVICO_SOLICITADO")+";"+
            rs.getString("Executado")+";"+
            rs.getString("NM_TIPO_EXECUCAO")+";"+
            rs.getString("dtLimiteExecucao")+";"+
            rs.getString("dtServico")+";"+
            rs.getString("dtFechamento")+";"+
            rs.getString("funcionarios")+";"+
            rs.getString("DescrSetorSolicitante")+";"+
            rs.getString("TerceiroNomeEmpresa")+";"+
            rs.getString("TerceiroNomeFantasia")+";"+
            rs.getString("TerceiroCNPJ")+";"+
            rs.getString("MatrizdeServiços");
            
            writer.write(Linha);
            writer.newLine();
            writer.flush();
        }
        relatorio.close();
    }
}
