/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package TratarRelatorios;

import com.smartsheet.api.SmartsheetException;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Date;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author vitor.heser
 */
public class ConverterPCSV {
    private String RelatorioFinal;
    private String RelatorioFinal2;
    private String Conteudo;
    

    
    public void TratarRelatorio(File Caminho, String CaminhoCSV) throws IOException, ParseException, SmartsheetException{
        File CaminhoPlanilha = Caminho;
        FileInputStream fis = new FileInputStream(CaminhoPlanilha);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row; 
        XSSFCell cell;

        
        FileWriter relatorio = new FileWriter(new File(CaminhoCSV));
        BufferedWriter writer = new BufferedWriter(relatorio);
        
        Integer j=0;
        for(int i = 7;i<10000;i++){
            row = sheet.getRow(i);
            try{
                if(row.getCell(0).toString().replaceAll(";","").equals("")){
                    break;
                }
            }catch(Exception p){
                break;
            }
            String Coluna1 ="";   try{ Coluna1 = row.getCell(0).toString().replaceAll(";","");    } catch(Exception e) { Coluna1="0";  }
            Double Coluna2 =0.0;   try{ Coluna2 = Double.valueOf(row.getCell(1).toString().replaceAll(";",""));    } catch(Exception e) { Coluna2=0.0;  }
            Double Coluna3 =0.0;   try{ Coluna3 = Double.valueOf(row.getCell(2).toString().replaceAll(";",""));    } catch(Exception e) { Coluna3=0.0;  }
            String Coluna4 ="";   try{ Coluna4 = row.getCell(3).toString().replaceAll(";","");    } catch(Exception e) { Coluna4="0";  }
            Double Coluna5 =0.0;   try{ Coluna5 = Double.valueOf(row.getCell(4).toString().replaceAll(";",""));    } catch(Exception e) { Coluna5=0.0;  }
            String Coluna6 ="";   try{ Coluna6 = row.getCell(5).toString().replaceAll(";","");    } catch(Exception e) { Coluna6="0";  }
            String Coluna7 ="";   try{ Coluna7 = row.getCell(6).toString().replaceAll(";","");    } catch(Exception e) { Coluna7="0";  }
            String Coluna8 ="";   try{ Coluna8 = row.getCell(7).toString().replaceAll(";","");    } catch(Exception e) { Coluna8="0";  }
            String Coluna9 = null;   
            try{ Coluna9 = conversordeData(row.getCell(8).toString().replaceAll(";",""));    } catch(Exception e) { Coluna9=null;  }
            String Coluna10 ="";  try{ Coluna10 = row.getCell(9).toString().replaceAll(";","");   } catch(Exception e) { Coluna10="0"; }
            String Coluna11 ="";  try{ Coluna11 = row.getCell(10).toString().replaceAll(";","");  } catch(Exception e) { Coluna11="0"; }
            String Coluna12 =null;   
            try{ Coluna12 = conversordeData(row.getCell(11).toString().replaceAll(";",""));    } catch(Exception e) { Coluna12=null;  }
            String Coluna13 ="";  try{ Coluna13 = row.getCell(12).toString().replaceAll(";","");  } catch(Exception e) { Coluna13="0"; }
            String Coluna14 ="";  try{ Coluna14 = row.getCell(13).toString().replaceAll(";","");  } catch(Exception e) { Coluna14="0"; }
            String Coluna15 ="";  try{ Coluna15 = row.getCell(14).toString().replaceAll(";","");  } catch(Exception e) { Coluna15="0"; }
            String Coluna16 = null;   
            try{ Coluna16 = conversordeData(row.getCell(15).toString().replaceAll(";",""));    } catch(Exception e) { Coluna16=null;  }
            Double Coluna17 =0.0;  try{ Coluna17 = Double.valueOf(row.getCell(16).toString().replaceAll(";",""));  } catch(Exception e) { Coluna17=0.0; }
            String Coluna18 ="";  try{ Coluna18 = row.getCell(17).toString().replaceAll(";","");  } catch(Exception e) { Coluna18="0"; }

            String Row = Coluna1 + ";"+Coluna2+ ";"+Coluna3+ ";"+Coluna4+ ";"+Coluna5+ ";"+Coluna6+ ";"+Coluna7+ ";"+Coluna8+ ";"+Coluna9+ ";"+Coluna10+
            Coluna11 + ";"+Coluna12+ ";"+Coluna13+ ";"+Coluna14+ ";"+Coluna15+ ";"+Coluna16+ ";"+Coluna17+ ";"+Coluna18;
            writer.write(Row);
            writer.newLine();
            writer.flush();
        }
        writer.close();
        
        
    }
    
    public String conversordeData(String Data) throws ParseException{
        String Dia = Data.substring(0,2);
        String Mes = Data.substring(3,6).toLowerCase();
        String Ano = Data.substring(7,11);
        switch (Mes){
            case "jan":
                Mes = "01";
                break;
            case "fev":
                Mes = "02";
                break;
            case "mar":
                Mes = "03";
                break;
            case "abr":
                Mes = "04";
                break;
            case "mai":
                Mes = "05";
                break;
            case "jun":
                Mes = "06";
                break;
            case "jul":
                Mes = "07";
                break;
            case "ago":
                Mes = "08";
                break;
            case "set":
                Mes = "09";
                break;
            case "out":
                Mes = "10";
                break;
            case "nov":
                Mes = "11";
                break;
            case "dez":
                Mes = "12";
                break;
        }
//        System.out.println(Dia+"/"+Mes+"/"+Ano);
        String dat = Ano+"-"+Mes+"-"+Dia;
        
        return dat;
    }
}
