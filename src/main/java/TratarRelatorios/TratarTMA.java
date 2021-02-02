 /*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package TratarRelatorios;

import java.sql.ResultSet;
import java.sql.PreparedStatement;
import java.io.Writer;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.File;
import java.io.IOException;
import com.smartsheet.api.SmartsheetException;
import java.util.List;
import com.smartsheet.api.models.Row;
import java.util.ArrayList;
import java.sql.SQLException;
import java.sql.DriverManager;
import java.sql.Connection;

public class TratarTMA
{
    public String consulta;
    
    public TratarTMA() {
        this.consulta = "SELECT \n\tcount(*) as anoos,\n\tidosp,\n\tAnomes,\n\tmatricula,\n\tENDerecoosp,\n\tBairro,\n\tservico,\n\tdescricao,\n\tEqp,\n\tDataEmissao,\n\tDataProgramacao,\n\tDATAINICIO,\n\tDataConclusao,\n\tstatus,\n\tATRASO,\n\tnomecor,\n\tcodexecocor,\n\tmotivoexecocor,\n\tidanomalia,\n\tdescranomalia,\n\ttempopadrao,\n\ttempodesloc,\n\ttempodeslocsegundos,\n\ttempopreparo,\n\ttempopreparosegundos,\n\ttempoexecucao,\n\ttempoexecucaosegundos,\n\ttempopreenchimento,\n\ttempopreenchimentosegundos,\n\ttempoparalisacao,\n\ttempoparalisacaosegundos,\n\ttempoatendimento,\n\tDescrSetor,\n\tlatitude,\n\tlongitude,\n\tFotosMin,\n\tQtdeFotos,\n\tDAT_AGENDAMENTO,\n\tSF,\n\tlatitudeExecucao,\n\tlongitudeExecucao,\n\tID_CONTRATO,\n\tCD_SUB_REGIAO,\n\tCD_REGIAO,\n\tidservico,\n\tidExecucaoServico,\n\tServi\u00e7oPrincipal,\n\tidExecucaoServicoPrincipal,\n\tunidade,\n\tCD_SERVICO_SOLICITADO,\n\tDS_SERVICO_SOLICITADO,\n\tExecutado,\n\tNM_TIPO_EXECUCAO,\n\tdtLimiteExecucao,\n\tdtServico,\n\tdtFechamento,\n\tfuncionarios,\n\tDescrSetorSolicitante,\n\tTerceiroNomeEmpresa,\n\tTerceiroNomeFantasia,\n\tTerceiroCNPJ,\n\tMatrizdeServi\u00e7os\n\n\nFROM sispc.tmaemergencial\ngroup by \n\tdescricao,\n\tdataemissao,\n\tDataConclusao,\n\tstatus,\n\teqp,\n\tatraso,\n\tcodexecocor,\n\tsf,\n\tdtLimiteExecucao";
    }
    
    public Connection getConnection() throws ClassNotFoundException, SQLException {
        Connection connection = null;
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            connection = DriverManager.getConnection("jdbc:mysql://dskprl013862:3306/sispc?useTimezone=true&serverTimezone=UTC", "sispcAPI", "Prol@go55i5PC@pi");
        }
        catch (ClassNotFoundException | SQLException ex) {
            //final Exception ex;
            final Exception e = ex;
            ex.printStackTrace();
        }
        return connection;
    }
    
    public void InserirLinhas(final String Caminho) throws SmartsheetException, IOException, ClassNotFoundException, SQLException {
        System.out.println("Entrou aqui, Juliano");
        final List<Row> Linhas2 = new ArrayList<Row>();
        final List<String> Colunas = new ArrayList<String>();
        final Connection connection = this.getConnection();
        String querry = "TRUNCATE sispc.tmaEmergencial;";
        connection.prepareStatement(querry).executeUpdate();
        querry = "";
        final TextFile arquivo = new TextFile(Caminho);
        arquivo.openTextFile();
        int prim = 0;
        int ident = 0;
        final String querryInicio = "INSERT INTO sispc.tmaEmergencial VALUES ";
        while (arquivo.next()) {
            final String Linha = arquivo.readLine().replaceAll("=", "");
            System.out.println(Linha);
            if (prim > 6) {
                final String[] vDados = Linha.split("\";\"");
                String ReclamaPipa = "";
                for (int i = 0; i < Linha.length(); ++i) {
                    try {
                        final String dado = vDados[i].toString().replaceAll(";", ",").replaceAll("'", "");
                        Label_0536: {
                            switch (i) {
                                case 10:
                                case 11:
                                case 12:
                                case 13:
                                case 54:
                                case 55:
                                case 56: {
                                    try {
                                        final String Dia = dado.substring(0, 2);
                                        final String Mes = dado.substring(3, 5);
                                        final String Ano = dado.substring(6, 10);
                                        Colunas.add("'" + Ano + "-" + Mes + "-" + Dia + "'");
                                    }
                                    catch (Exception e) {
                                        Colunas.add("null");
                                    }
                                    break;
                                }
                                case 7: {
                                    if (dado.contains("RECLAMA")) {
                                        ReclamaPipa = "RECLAMA\u00c7\u00c3O";
                                        break Label_0536;
                                    }
                                    ReclamaPipa = "PIPA";
                                    break Label_0536;
                                }
                                case 0:
                                case 1:
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                case 6:
                                case 9:
                                case 14:
                                case 15:
                                case 16:
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
                                case 53:
                                case 57:
                                case 58:
                                case 59:
                                case 60:
                                case 61:
                                case 62: {
                                    Colunas.add(dado);
                                    break;
                                }
                                case 17: {
                                    Colunas.add(ReclamaPipa);
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception e2) {
                        Colunas.add("NULL");
                    }
                }
                querry = String.valueOf(querry) + "('" + Colunas.get(0) + "','" + Colunas.get(1) + "','" + Colunas.get(2) + "','" + Colunas.get(3) + "','" + Colunas.get(4) + "','" + Colunas.get(5) + "','" + Colunas.get(6) + "','" + Colunas.get(7) + "'," + Colunas.get(8) + "," + Colunas.get(9) + "," + Colunas.get(10) + "," + Colunas.get(10) + "," + Colunas.get(11) + ",'" + Colunas.get(12) + "','" + Colunas.get(13) + "','" + Colunas.get(14) + "','" + Colunas.get(15) + "','" + Colunas.get(16) + "','" + Colunas.get(17) + "','" + Colunas.get(18) + "','" + Colunas.get(19) + "','" + Colunas.get(20) + "','" + Colunas.get(21) + "','" + Colunas.get(22) + "','" + Colunas.get(23) + "','" + Colunas.get(24) + "','" + Colunas.get(25) + "','" + Colunas.get(26) + "','" + Colunas.get(27) + "','" + Colunas.get(28) + "','" + Colunas.get(29) + "','" + Colunas.get(30) + "','" + Colunas.get(31) + "','" + Colunas.get(32) + "','" + Colunas.get(33) + "','" + Colunas.get(34) + "','" + Colunas.get(35) + "','" + Colunas.get(36) + "','" + Colunas.get(37) + "','" + Colunas.get(38) + "','" + Colunas.get(39) + "','" + Colunas.get(40) + "','" + Colunas.get(41) + "','" + Colunas.get(42) + "','" + Colunas.get(43) + "','" + Colunas.get(44) + "','" + Colunas.get(45) + "','" + Colunas.get(46) + "','" + Colunas.get(47) + "','" + Colunas.get(48) + "','" + Colunas.get(49) + "','" + Colunas.get(50) + "','" + Colunas.get(51) + "'," + Colunas.get(52) + "," + Colunas.get(53) + "," + Colunas.get(54) + ",'" + Colunas.get(55) + "','" + Colunas.get(56) + "','" + Colunas.get(57) + "','" + Colunas.get(58) + "','" + Colunas.get(59) + "','" + Colunas.get(60) + "'),";
                
                Colunas.clear();
                if (ident == 70) {
                    querry = String.valueOf(querryInicio) + querry.substring(0, querry.length() - 1) + ";";
                    System.out.println(querry);
                    connection.prepareStatement(querry).executeUpdate();
                    querry = "";
                    ident = 0;
                }
                ++ident;
            }
            ++prim;
        }
        if (ident > 0) {
            querry = String.valueOf(querryInicio) + querry.substring(0, querry.length() - 1) + ";";
            System.out.println(querry);
            connection.prepareStatement(querry).executeUpdate();
            querry = "";
            ident = 0;
        }
    }
    
    public void PlotarCSV(final String Caminho) throws ClassNotFoundException, SQLException, IOException {
        final Connection connection = this.getConnection();
        final PreparedStatement ps = connection.prepareStatement(this.consulta);
        final ResultSet rs = ps.executeQuery();
        final FileWriter relatorio = new FileWriter(new File(Caminho));
        final BufferedWriter writer = new BufferedWriter(relatorio);
        while (rs.next()) {
            final String Linha = String.valueOf(rs.getString("anoos")) + ";" + rs.getString("idosp") + ";" + rs.getString("Anomes") + ";" + rs.getString("matricula") + ";" + rs.getString("ENDerecoosp") + ";" + rs.getString("Bairro") + ";" + rs.getString("servico") + ";" + rs.getString("descricao") + ";" + rs.getString("Eqp") + ";" + rs.getString("DataEmissao") + ";" + rs.getString("DataProgramacao") + ";" + rs.getString("DATAINICIO") + ";" + rs.getString("DataConclusao") + ";" + rs.getString("status") + ";" + rs.getString("ATRASO") + ";" + rs.getString("nomecor") + ";" + rs.getString("codexecocor") + ";" + rs.getString("motivoexecocor") + ";" + rs.getString("idanomalia") + ";" + rs.getString("descranomalia") + ";" + rs.getString("tempopadrao") + ";" + rs.getString("tempodesloc") + ";" + rs.getString("tempodeslocsegundos") + ";" + rs.getString("tempopreparo") + ";" + rs.getString("tempopreparosegundos") + ";" + rs.getString("tempoexecucao") + ";" + rs.getString("tempoexecucaosegundos") + ";" + rs.getString("tempopreenchimento") + ";" + rs.getString("tempopreenchimentosegundos") + ";" + rs.getString("tempoparalisacao") + ";" + rs.getString("tempoparalisacaosegundos") + ";" + rs.getString("tempoatendimento") + ";" + rs.getString("DescrSetor") + ";" + rs.getString("latitude") + ";" + rs.getString("longitude") + ";" + rs.getString("FotosMin") + ";" + rs.getString("QtdeFotos") + ";" + rs.getString("DAT_AGENDAMENTO") + ";" + rs.getString("SF") + ";" + rs.getString("latitudeExecucao") + ";" + rs.getString("longitudeExecucao") + ";" + rs.getString("ID_CONTRATO") + ";" + rs.getString("CD_SUB_REGIAO") + ";" + rs.getString("CD_REGIAO") + ";" + rs.getString("idservico") + ";" + rs.getString("idExecucaoServico") + ";" + rs.getString("Servi\u00e7oPrincipal") + ";" + rs.getString("idExecucaoServicoPrincipal") + ";" + rs.getString("unidade") + ";" + rs.getString("CD_SERVICO_SOLICITADO") + ";" + rs.getString("DS_SERVICO_SOLICITADO") + ";" + rs.getString("Executado") + ";" + rs.getString("NM_TIPO_EXECUCAO") + ";" + rs.getString("dtLimiteExecucao") + ";" + rs.getString("dtServico") + ";" + rs.getString("dtFechamento") + ";" + rs.getString("funcionarios") + ";" + rs.getString("DescrSetorSolicitante") + ";" + rs.getString("TerceiroNomeEmpresa") + ";" + rs.getString("TerceiroNomeFantasia") + ";" + rs.getString("TerceiroCNPJ") + ";" + rs.getString("MatrizdeServi\u00e7os");
            writer.write(Linha);
            writer.newLine();
            writer.flush();
        }
        relatorio.close();
    }
}