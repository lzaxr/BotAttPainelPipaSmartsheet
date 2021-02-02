    /*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package MontarRelatorio.Email;

    


import TratarRelatorios.TextFile;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author vitor.heser
 */
public class ModelarBaseEmails {
    
    //Gerencias
    
    public ArrayList<String> Departamento = new ArrayList<>();
    public ArrayList<String> Cc = new ArrayList<>();
    
    private String email;
    private String departamento;

    
    
    
    public void BaseAreas(String pastadestino) throws IOException, SQLException, ClassNotFoundException{
        Departamento.clear();
        ResultSet rs=null;
        TextFile arquivo = new TextFile(pastadestino+"EmailsReceptores.csv");
        arquivo.openTextFile();        
        
        while (arquivo.next()){
            String Linha = arquivo.readLine();
            String[] vDados = Linha.split("[;]");
            Departamento.add(vDados[0]);
            System.out.println(vDados[0]);
        }
//        Departamento.clear();
//        Departamento.add("vitor.heser@prolagos.com.br");
    }
}
