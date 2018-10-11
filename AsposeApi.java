

package asposeapplication;

import java.util.*;
import com.aspose.tasks.Asn;
import com.aspose.tasks.Calendar;
import com.aspose.tasks.CostAccrualType;
import com.aspose.tasks.DayType;
import com.aspose.tasks.ExtendedAttributeDefinition;
import com.aspose.tasks.NullableBool;
import com.aspose.tasks.Prj;
import com.aspose.tasks.Project;
import com.aspose.tasks.RateFormatType;
import com.aspose.tasks.Resource;
import com.aspose.tasks.ResourceAssignment;
import com.aspose.tasks.ResourceCollection;
import com.aspose.tasks.ResourceType;
import com.aspose.tasks.Rsc;
import com.aspose.tasks.SaveFileFormat;
import com.aspose.tasks.Task;
import com.aspose.tasks.TaskLink;
import com.aspose.tasks.TaskLinkCollection;
import com.aspose.tasks.TaskLinkType;
import com.aspose.tasks.TimeUnitType;
import com.aspose.tasks.Tsk;
import com.aspose.tasks.WeekDay;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
/**
*
* @author ishu
*/
public class AsposeApp {
    public static final boolean DEBUG = false;
    public static final long MILLISECS_PER_MINUTE = 60*1000;
    public static final long MILLISECS_PER_HOUR   = 60*MILLISECS_PER_MINUTE;
    public static final long MILLISECS_PER_DAY = 24*MILLISECS_PER_HOUR;
    public static AsposeApp asposeApp = new AsposeApp();
   
    public static void setAsposeLicense() throws Exception{
        com.aspose.tasks.License license = new com.aspose.tasks.License();
        license.setLicense("Aspose.Tasks.lic");
    }
   
    public static Connection getConnection()throws ClassNotFoundException, SQLException{
        Class.forName("oracle.jdbc.OracleDriver");
        Connection connection = DriverManager.getConnection("dbconn”); 
        return connection;
    }
   
    private static Date getDateTime(Timestamp timestamp) throws ParseException{   
        Date retDate = new Date(timestamp.getTime() + (timestamp.getNanos() / 1000000));       
        return retDate;
    }
   
    public String getQueryStatement(String pStartDate, String pEndDate, String pTablename){
        StringBuffer queryStmt = new StringBuffer();
        queryStmt.append("select * FROM  ");
        if (pTablename != null && pTablename.length()>0){
            if (!pTablename.equalsIgnoreCase("null"))
                queryStmt.append(pTablename).append(" ");
        }else{
            pTablename = "T1 ";
        }
       
        queryStmt.append(",xmltable(('\"'|| REPLACE(equipment, ',', '\",\"') || '\"'))");
        queryStmt.append("WHERE Cond");
        return queryStmt.toString();
    }
   
   
    public static long getActualDurationInDays(Date startDate, Date endDate){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("EEEE");
       java.util.Calendar start = java.util.Calendar.getInstance();
        start.setTime(startDate);
        java.util.Calendar end = java.util.Calendar.getInstance();
        end.setTime(endDate);
        int numberOfDays = 0;
        for (Date date = start.getTime();
                start.before(end);
                start.add(java.util.Calendar.DATE, 1), date = start.getTime()) {
            String dayOfWeek = simpleDateFormat.format(date);
            if(dayOfWeek.equals("Saturday") || dayOfWeek.equals("Sunday")){
                continue;
            }else{
                ++numberOfDays;
            }
        }
        return numberOfDays;
    }
   
    public static long getActualDurationInHours(Date startDate, Date endDate){
        return getActualDurationInDays(startDate, endDate) * 8;
    }
     public static long getDurationInDays(java.util.Date startDate, java.util.Date endDate)
    {
        java.util.Calendar start = new GregorianCalendar();
        start.setTimeInMillis(startDate.getTime());
        java.util.Calendar end = new GregorianCalendar();
        end.setTimeInMillis(endDate.getTime());
        
        long endL   =  end.getTimeInMillis() +  end.getTimeZone().getOffset(  end.getTimeInMillis() );
        long startL = start.getTimeInMillis() + start.getTimeZone().getOffset( start.getTimeInMillis() );
        return (endL - startL) / MILLISECS_PER_DAY;
    }
   
     public static Map<String, Integer> resourceMap = new HashMap<>();
     public static Resource ResourceExists(ResourceCollection resourceCollection, String equipment, Project project){
         if(resourceMap.containsKey(equipment)){
            int uid = resourceMap.get(equipment);
            Resource res = resourceCollection.getByUid(uid);
            return res;
         }else{
             Resource res = project.getResources().add(equipment);
             int uid = res.get(Rsc.UID);
             resourceMap.put(equipment, uid);
             return res;
         }
     }
    
    public static List adjustStartAndEndDate(Date sDate, Date endDate){
        List<Date> dateList = new ArrayList<>();
        java.util.Calendar cal = java.util.Calendar.getInstance();
        cal.setTime(sDate);
        cal.set(java.util.Calendar.HOUR_OF_DAY,8);
        cal.set(java.util.Calendar.MINUTE,0);
        cal.set(java.util.Calendar.SECOND,0);
        cal.set(java.util.Calendar.MILLISECOND,0);
        dateList.add(cal.getTime());
       
        cal.setTime(endDate);
        cal.set(java.util.Calendar.HOUR_OF_DAY,17);
        cal.set(java.util.Calendar.MINUTE,0);
        cal.set(java.util.Calendar.SECOND,0);
        cal.set(java.util.Calendar.MILLISECOND,0);
        dateList.add(cal.getTime());
        return dateList;
    }
    public void executeSQLQuery(String queryStmt, Project project, Task task, Calendar cal)throws ClassNotFoundException,
            SQLException, ParseException, IOException{
        Connection con = getConnection();
        Statement stmt = con.createStatement();
        ResultSet rset = stmt.executeQuery(queryStmt);
       
        while (rset.next()) {
            String appNum = rset.getString(1) + "  "+rset.getString(2);
            String equipment = rset.getString(2);
            Timestamp sDate = rset.getTimestamp(3);
            Timestamp endDate = rset.getTimestamp(4);
            Date projectStartDate = null;
                  
            if (projectStartDate == null)
                projectStartDate = getDateTime(sDate);
            else if (getDateTime(sDate).getTime() < projectStartDate.getTime())
                projectStartDate = getDateTime(sDate);
            
            Date _start = getDateTime(sDate);
            Date _end = getDateTime(endDate);
                List<Date> dateList = adjustStartAndEndDate(_start, _end);
                _start = dateList.get(0);
                _end = dateList.get(1);  
            long numOfDays = getActualDurationInDays(_start, _end);
            if(numOfDays == 0)
                continue;
           
            project.set(Prj.START_DATE, projectStartDate);
            Task subtask = task.getChildren().add(appNum);
            subtask.set(Tsk.IS_MARKED, true);
            subtask.set(Tsk.IGNORE_WARNINGS, true);
            subtask.set(Tsk.ACTUAL_START,_start);
            subtask.set(Tsk.ACTUAL_FINISH,_end);
            subtask.set(Tsk.ACTUAL_DURATION, project.getDuration(numOfDays,TimeUnitType.Day));
            ResourceCollection resourceCollection = project.getResources();
            Resource rsc = ResourceExists(resourceCollection, equipment, project);
            rsc.set(Rsc.TYPE, ResourceType.Work);
            rsc.set(Rsc.CALENDAR,cal);
            rsc.set(Rsc.INITIALS, "WR");
            rsc.set(Rsc.ACCRUE_AT, CostAccrualType.Prorated);
            rsc.set(Rsc.MAX_UNITS, 1d);
            rsc.set(Rsc.CODE, "Code 1");
            rsc.set(Rsc.GROUP, "Workers");
            rsc.set(Rsc.E_MAIL_ADDRESS, "1@gmail.com");
            rsc.set(Rsc.WINDOWS_USER_ACCOUNT, "user_acc1");
            rsc.set(Rsc.IS_GENERIC, new NullableBool(true));
            rsc.set(Rsc.ACCRUE_AT, CostAccrualType.End);
            rsc.set(Rsc.OVERTIME_RATE_FORMAT, RateFormatType.Hour);
            rsc.set(Rsc.START,_start);
            rsc.set(Rsc.FINISH,_end);
            rsc.set(Rsc.IS_TEAM_ASSIGNMENT_POOL, true);
            rsc.set(Rsc.COST_CENTER, "Cost Center 1");
            rsc.set(Rsc.GROUP, "Workgroup1");
                 
            ResourceAssignment assn = project.getResourceAssignments().add(subtask, rsc);
            
        }
        rset.close();
        rset = null;
        stmt.close();   // Close the connection object
        stmt = null;
        con.close();
        con = null;
           
        project.save("updated.xml", SaveFileFormat.XML);  
        project = new Project("file1.xml");
        Project mppFile = new Project("file1.mpp");
        project.copyTo(mppFile);
        mppFile.save("output.mpp", SaveFileFormat.MPP);
       
    }
    
    public static void main(String [] args)throws Exception{
    
        String pStartDate = "08/09/2018";
        String pEndDate   = "08/23/2018";
        String pTablename = "orig_table";
    
        // Set Aspose License
        setAsposeLicense();
        String queryStmt = asposeApp.getQueryStatement(pStartDate, pEndDate, pTablename);
       
        //Get DB connection object
        Project project = new Project();
       
        com.aspose.tasks.Calendar cal = project.getCalendars().add("My Cal");
        Calendar.makeStandardCalendar(cal);
       
        
        Task task = project.getRootTask().getChildren().add("Title of project");
        task.set(Tsk.CALENDAR, cal);
        asposeApp.executeSQLQuery(queryStmt, project, task, cal);
    }
   
}


