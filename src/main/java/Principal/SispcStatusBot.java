/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Principal;

import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

/**
 *
 * @author vitor.heser
 */
public class SispcStatusBot {
    Integer bot;
    String Hostname;
    String Ip;
    String user;
    String hora;
        
    public SispcStatusBot(Integer bot) throws UnknownHostException{
        SimpleDateFormat datahoje = new SimpleDateFormat("yyy-MM-dd HH:mm:ss");
        int x = -30;
        Calendar cal = GregorianCalendar.getInstance();
        Date datinGregorian = cal.getTime();
        System.out.println(this.user);
        this.bot=bot;
        this.Hostname = InetAddress.getLocalHost().getHostName();
        this.Ip =InetAddress.getLocalHost().getHostAddress();
        this.user = System.getProperty("user.name");
        this.hora = datahoje.format(datinGregorian);
    }
    
    
    public Connection getnewConnection() throws ClassNotFoundException, SQLException{
        Connection connection=null;
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            connection = DriverManager.getConnection("jdbc:mysql://SISPCPRL01:3306/sispc?autoReconnect=true&useSSL=false","sispcRPA", "Prol@go55i5PCRp@s");

        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
    //            System.out.println(e);
        }
    //        System.out.println(connection.getWarnings());
        return connection;
    }
    public void Status(String Status) throws IOException, SQLException, ClassNotFoundException{
        
        String Insert = "INSERT INTO sispc.apprpa_rpa_statusbots VALUES (null,'"+hora+"','"+user+"','"+Hostname+"','"+Ip+"',"+bot+",'"+Status+"',null);";
        Connection connection=null;
        connection=getnewConnection();

        //===============================================================================
        PreparedStatement ps=connection.prepareStatement(Insert);
        //ps.executeUpdate();    //!!!!!!!!!!!!!!!!!!!!!!!!!REsolver comunicação
    }
    public void Status(String Status, String Exception) throws IOException, ClassNotFoundException, SQLException{
        String Insert = "INSERT INTO sispc.apprpa_rpa_statusbots VALUES (null,'"+hora+"','"+user+"','"+Hostname+"','"+Ip+"',"+bot+",'"+Status+"','"+Exception+"');";
        Connection connection=null;
        connection=getnewConnection();

        //===============================================================================
        PreparedStatement ps=connection.prepareStatement(Insert);
        //ps.executeUpdate();         //!!!!!!!!!!!!!!!!!!!!!!!!!REsolver comunicação
    }
  
}
