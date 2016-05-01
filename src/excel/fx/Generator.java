package excel.fx;


import java.io.*;
import java.util.*;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generator {    
    //attaches spreadsheets
    public static final String FileNameResponses = "xlsx/Grade 8 (Responses).xlsx";
    public static final String FileNameSchedule = "xlsx/Grade 8 (Schedule).xlsx";
    public static final String FileNameLabels = "xlsx/Grade 8 (Labels).xlsx";
    public static final String FileNameSubject = "xlsx/Grade 8 (Subject).xlsx";
    public static final String FileNameScheduleType = "xlsx/Grade 8 (ScheduleType).xlsx";
    
    //array lists for all objects 
    private ArrayList<Response> TemporaryResponseList = new ArrayList();
    private ArrayList<Response> ResponseList = new ArrayList();
    private ArrayList<ScheduleLabel> LabelList = new ArrayList();
    private ArrayList<Subject> SubjectList = new ArrayList();
    private ArrayList<Schedule> ScheduleList = new ArrayList();
    private ArrayList<ScheduleType> ScheduleTypeList = new ArrayList(); 
 /**
* This function reads the response and subject spreadsheets, run the algorithm 
* and generates schedule and label spreadsheets.  
     * @throws java.io.IOException
*/
//    public void run () throws IOException
//    { 
//        ExcelFX.lbl.setText("Reading Subjects...");
//        readSubjects();
//        ExcelFX.lbl.setText("Reading Responses...");
//        readResponses(); //converts spreadsheet to array list
//         ExcelFX.lbl.setText("Reading Schedule Type...");
//        readScheduleType();
//        
//        ExcelFX.lbl.setText("Generating Schedule...");
//        generateSchedule();
//         
//         ExcelFX.lbl.setText("Generating Labels...");
//        generateLabels(); //coverts schedule array list to label array list
//        
//         ExcelFX.lbl.setText("Writing Schedule...");
//        writeSchedule();//creates spreadsheet from schedule array list 
//        
//        ExcelFX.lbl.setText("Writing Label...");
//        writeLabel();//creates spreadsheet from label array list 
//        
//         ExcelFX.lbl.setText("Completed...");
//    }
    
    public void readResponses (Boolean bRandom) throws FileNotFoundException, IOException
    {
        File myFile = new File(FileNameResponses);
        FileInputStream fis = new FileInputStream(myFile);
        // Finds the workbook instance for XLSX file 
        XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
        // Traversing over each row of XLSX file
        int rowIndex = 0;
        while (rowIterator.hasNext()) { 
            Row row = rowIterator.next();
            // For each row, iterate through each columns
            if (rowIndex == 0 ) {
                rowIndex++;continue;
            }
            Iterator<Cell> cellIterator = row.cellIterator();
            //edit here
            int column = 1;
            Response response = new Response();
            response.ID = rowIndex;
            while (cellIterator.hasNext()) {
                
                Cell cell = cellIterator.next();
                switch (column)
                {
                    case 1: //Date
                        //response.Timestamp = cell.getDateCellValue();
                        break;
                    case 2: //First Name
                        response.FirstName = cell.getStringCellValue();
                        break;
                    case 3: //Last Name
                        response.LastName = cell.getStringCellValue();
                        break;
                    case 4: //Response
                        setSubject(response,findSubjectByID(1),cell.getNumericCellValue());
                        break;
                    case 5: //Response
                        setSubject(response,findSubjectByID(2),cell.getNumericCellValue());
                        break;
                    case 6: //Response
                        setSubject(response,findSubjectByID(3),cell.getNumericCellValue());
                        break;
                    case 7: //Response
                        setSubject(response,findSubjectByID(4),cell.getNumericCellValue());
                        break;
                    case 8: //Response
                        setSubject(response,findSubjectByID(5),cell.getNumericCellValue());
                        break;
                    case 9: //Response
                        setSubject(response,findSubjectByID(6),cell.getNumericCellValue());
                        break;
                    case 10: //Response
                        setSubject(response,findSubjectByID(7),cell.getNumericCellValue());
                        break;
                    case 11: //Response
                        response.School = cell.getStringCellValue();
                        break;
                        
                }
                column ++;
            }
            if (response.FirstName.length() > 0 ||  response.LastName.length() > 0 )
            {
                rowIndex++;
                TemporaryResponseList.add(response);
            }
        }
        
        if (bRandom)
        {
            System.out.println("ResponseList Size Is: " + TemporaryResponseList.size());
            Integer[] arr = new Integer[TemporaryResponseList.size()];
            for (int i = 0; i < arr.length; i++) {
                arr[i] = i+1;
            }
            Collections.shuffle(Arrays.asList(arr));
            System.out.println(Arrays.toString(arr));

            //Save the random order to the response list
            for (int i = 0; i < arr.length; i++) {
                int id = arr[i];
                Response response = findResponseById(id);
                if (response != null)
                {
                    System.out.println("Processing Id=" + response.ID );
                    ResponseList.add(response);
                }
            }
        }
        else
        {
            //Save the original order to the response list
            for(Response response : TemporaryResponseList)
                ResponseList.add(response);
        }
    }
    public void readSubjects () throws FileNotFoundException, IOException
    {   
        int rowIndex = 0;
        //object file
        File myFile = new File(FileNameSubject); 
        //let's it read information from the file
        FileInputStream fis = new FileInputStream(myFile); 
        // Finds the workbook instance for XLSX file 
        XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) { 
            Row row = rowIterator.next();
            // For each row, iterate through each columns
            if (rowIndex == 0) {
                rowIndex++;continue;
            }
            int column = 1;
            Subject subject = new Subject(); 
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                
                Cell cell = cellIterator.next();
                switch (column)
                {
                    case 1: 
                        subject.ID = rowIndex;
                        subject.Name = cell.getStringCellValue();
                        break;
                    case 2: //Name
                         subject.Limit = cell.getNumericCellValue();
                         break;       
                }
                column ++;
            }
            rowIndex++;
            SubjectList.add(subject);
        } 
    }
    public void readScheduleType () throws FileNotFoundException, IOException {
         int rowIndex = 1;
        //object file
        File myFile = new File(FileNameScheduleType); 
        //let's it read information from the file
        FileInputStream fis = new FileInputStream(myFile); 
        // Finds the workbook instance for XLSX file 
        XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) { 
            Row row = rowIterator.next();
            // For each row, iterate through each columns
            if (rowIndex == 1 ) {
                rowIndex++;continue;
            }
            int column = 1;
            ScheduleType scheduletype = new ScheduleType(); 
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                
                Cell cell = cellIterator.next();
                switch (column)
                {
                    case 1: 
                        scheduletype.ID = cell.getNumericCellValue();
                        break;
                    case 2: //Name
                         scheduletype.Name = cell.getStringCellValue(); 
                         break;       
                }
                column ++;
            }
            rowIndex++;
            ScheduleTypeList.add(scheduletype);
        } 
    
    }
    //******************************
            //Prints labels
    //******************************
    public void generateLabels () throws FileNotFoundException
    {
       //generate labels
        for(Response response: ResponseList)
        {
            ScheduleLabel label = new ScheduleLabel();
            label.Name = response.FirstName + ' ' + response.LastName;
            label.SchoolName = response.School;
            
            label.Subject1 = findScheduleByScheduleTypeAndResponse ( ScheduleTypeList.get(0), response ).Subject; // prints choice 1
            label.Subject2 = findScheduleByScheduleTypeAndResponse ( ScheduleTypeList.get(1), response ).Subject; // prints choice 2
            label.Subject3 = findScheduleByScheduleTypeAndResponse ( ScheduleTypeList.get(2), response ).Subject; // prints choice 3 
            LabelList.add(label);
        }
    }
    
    public Schedule findScheduleByScheduleTypeAndResponse ( ScheduleType st, Response r  )
    {
          for(Schedule s: ScheduleList)
              if (s.Type.ID == st.ID && s.Response.ID == r.ID )
                  return s;
          
           return null;
    }
    
     private Subject findSubjectByID (double id) 
    {     //by priority it finds the object that corresponds
          for(Subject s: SubjectList)
              if (s.ID == id)
                  return s; //returns object 
          
           return null;
    }
     
    private void setSubject (Response response, Subject subject, double priority) 
    {   //converts from rows to list
        if (priority == 1) 
              response.Priority1 = subject;
        else if (priority == 2)
            response.Priority2 = subject;
        else if (priority == 3)
            response.Priority3 = subject;
        else if (priority == 4)
            response.Priority4 = subject;
        else if (priority == 5)
            response.Priority5 = subject;
        else if (priority == 6)
            response.Priority6 = subject;
        else if (priority == 7)
            response.Priority7 = subject;
    }
     
    public void generateSchedule(){
     
        System.out.println("Priority 1: ###################################################" );
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority1, response);
        
        System.out.println("Priority 2: ###################################################" );
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority2, response);
        
        System.out.println("Priority 3: ###################################################" ); 
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority3, response);
        
        System.out.println("Priority 4: ###################################################" ); 
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority4, response);
        
        System.out.println("Priority 5: ###################################################" ); 
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority5, response);
        
        System.out.println("Priority 6: ###################################################" ); 
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority6, response);
     
        System.out.println("Priority 7: ###################################################" ); 
        for (ScheduleType st: ScheduleTypeList)
            for(Response response: ResponseList)
                addRespose( st, response.Priority7, response);
        
}
    
    private boolean addRespose ( ScheduleType scheduleType, Subject subject, Response response  )
    {
        try
        {
            Boolean findResponseSubject = findResponse(subject, response);
            Boolean findResponseType = findResponse(scheduleType, response);
            int count = currentStudent(scheduleType, subject);
            System.out.println("***************************************************" );
            System.out.println("Processing: Time=" + scheduleType.Name + "; Subject=" + subject.Name + "; Name=" + response.LastName + " " + response.FirstName );
            System.out.println("Is Subject Taken=" + findResponseSubject + "; Is Time Taken=" + findResponseType + "; Count=" + count  );
            if ( ! findResponseSubject && ! findResponseType && currentStudent(scheduleType, subject) < subject.Limit  )
            {
                Schedule schedule = new Schedule();    
                schedule.ID = ScheduleList.size() + 1;   
                schedule.Type = scheduleType;
                schedule.Subject = subject;
                schedule.Response = response;
                ScheduleList.add(schedule);

                System.out.println("Added: ID=" + schedule.ID + "; Time=" + schedule.Type.Name + "; Subject=" + schedule.Subject.Name + "; Name=" + schedule.Response.LastName + " " + schedule.Response.FirstName  );
                return true;
            }
        }
        catch ( Exception ex)
        {
            System.out.println( "ERROR:" + ex.toString()  );
        }
        
        return false;
    }
    
    private boolean findResponse ( ScheduleType scheduleType, Subject subject, Response response){
        for(Schedule s : ScheduleList){    
            if(s.Response == response && s.Subject == subject && s.Type == scheduleType )
                return true;   
        } 
        return false;
    }
     
     private boolean findResponse ( Subject subject, Response response){
        for(Schedule s : ScheduleList){    
            if(s.Response == response && s.Subject == subject )
                return true;   
        } 
        return false;
    }
     
     private boolean findResponse ( ScheduleType scheduleType, Response response){
        for(Schedule s : ScheduleList){    
            if(s.Response == response && s.Type == scheduleType )
                return true;   
        } 
    return false;
    }
    
    private int currentStudent (ScheduleType scheduletype,Subject subject){
        int count = 0;    
        for(Schedule s : ScheduleList){    
            if(s.Type == scheduletype && s.Subject == subject)
                count++;   
        }
        return count;
    }
    
    public void writeSchedule() throws FileNotFoundException, IOException{
       
       Boolean bShowHeaderScheduleType; 
       Boolean bShowHeaderSubject;  

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Schedule");
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        
        int rowCount = 0;
         XSSFRow row = sheet.createRow(rowCount);
         
         XSSFCell cell = row.createCell(0);
         cell.setCellValue(new XSSFRichTextString("Generated Schedule"));
         XSSFCellStyle cellStyle = workbook.createCellStyle();
         XSSFFont font = workbook.createFont();
         font.setFontHeightInPoints((short)30);
         font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
         font.setColor(HSSFColor.BLUE_GREY.index);
         cellStyle.setFont(font);
         cell.setCellStyle(cellStyle);
         rowCount++;
        
         for(ScheduleType st: ScheduleTypeList ) {   
             bShowHeaderScheduleType = true;
            for(Subject s: SubjectList ) {
                bShowHeaderSubject = true;
                for(Schedule sh: ScheduleList ) {
                    if ( sh.Type.ID == st.ID && sh.Subject.ID == s.ID ) {
                                    
                        row = sheet.createRow(rowCount);
                                    
                        //Schedule Type // style cell here
                        if (bShowHeaderScheduleType) {
                            cell = row.createCell(0);
                            cell.setCellValue(new XSSFRichTextString(st.Name));
                            cellStyle = workbook.createCellStyle();
                            font = workbook.createFont();
                            font.setFontHeightInPoints((short)12);
                            font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                            font.setColor(HSSFColor.BLUE_GREY.index);
                            cellStyle.setFont(font);
                            cell.setCellStyle(cellStyle);
                        }
                        //Subject
                        if ( bShowHeaderSubject )
                        {
                            cell = row.createCell(1);
                            cell.setCellValue(new XSSFRichTextString(s.Name));
                        }

                        cell = row.createCell(2);
                        cell.setCellValue(new XSSFRichTextString(sh.Response.FirstName + ' ' + sh.Response.LastName  ));
  
                        bShowHeaderScheduleType = false; 
                        bShowHeaderSubject = false;  
                        rowCount ++;
                      
                    }
                }                              
            }            
         }
         
        sheet.setColumnWidth(0,5000);
        sheet.setColumnWidth(1,5000);
         
         try (FileOutputStream outputStream = new FileOutputStream(FileNameSchedule)) {
            workbook.write(outputStream);
        }
 }
    
    public void writeLabel() throws IOException {
        
        ///////////////////////////////////////////////////////////////
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Labels");
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        int rowCount = 0;
        int rowCountTemp = 0;
        int column = 0;
        XSSFRow row = null ;
        XSSFRow rowTemp0 = null ;
        XSSFRow rowTemp1 = null ;
        XSSFRow rowTemp2 = null ;
        XSSFRow rowTemp3 = null ;
        XSSFRow rowTemp4 = null ;
        XSSFRow rowTemp5 = null ;
        
        for(ScheduleLabel l: LabelList )
        {
            if (column == 0)
            {
                row = sheet.createRow(rowCount);
                row.setHeight((short)1250);
                rowTemp0 = row;
            }
             
            // System.out.println("column=" + column );
            
            if (column == 0 )
            {
                rowCountTemp = rowCount;
                addLogo (workbook,sheet, rowCount, column);
            }
            
            if (column == 3 )
            {
                rowCount = rowCountTemp;
                row = rowTemp0;
                addLogo (workbook,sheet, rowCount, column);
            }
            
            //System.out.println("rowCount=" + rowCount );
            column++;
            
            //name 
             if (column == 1 || column == 4 )
             {
                XSSFCell cell = row.createCell(column);
                cell.setCellValue(new XSSFRichTextString(l.Name));
                XSSFCellStyle cellStyle = workbook.createCellStyle();
                XSSFFont font = workbook.createFont();
                font.setFontHeightInPoints((short)18);
                font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                font.setColor(HSSFColor.BLUE_GREY.index);
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
                rowCount++;
           
            
                if (column == 1)
                {
                   row = sheet.createRow(rowCount);
                   rowTemp1 = row;
                }
              
                if (column == 4)
                {
                   row = rowTemp1;
                }
              
                //SchoolName 
                cell = row.createCell(column);
                cell.setCellValue(new XSSFRichTextString( l.SchoolName));
                cellStyle = workbook.createCellStyle();
                font = workbook.createFont();
                font.setFontHeightInPoints((short) 14);
                font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                font.setColor(HSSFColor.BLUE_GREY.index);
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
                 rowCount++;
             
                //Class1
                if (column == 1)
                {
                   row = sheet.createRow(rowCount);
                   rowTemp2 = row;
                }
              
                if (column == 4)
                {
                   row = rowTemp2;
                }
                cell = row.createCell(column);
                cell.setCellValue(new XSSFRichTextString( "#1:" + l.Subject1.Name));
                rowCount++;
             
                //Class2
                if (column == 1)
                {
                   row = sheet.createRow(rowCount);
                   rowTemp3 = row;
                }

                if (column == 4)
                {
                   row = rowTemp3;
                }
                cell = row.createCell(column);
                cell.setCellValue(new XSSFRichTextString( "#2:" + l.Subject2.Name));
                rowCount++;
             
                //Class3 
                if (column == 1)
                {
                   row = sheet.createRow(rowCount);
                   rowTemp4 = row;
                }

                if (column == 4)
                {
                   row = rowTemp4;
                }
                cell = row.createCell(column);
                cell.setCellValue(new XSSFRichTextString( "#3:" + l.Subject3.Name));
                 rowCount++;
              
                //empty space 
                 if (column == 1)
                 {
                    row = sheet.createRow(rowCount);
                    rowTemp5 = row;
                 }

                 if (column == 4)
                    row = rowTemp5;
                 
                 cell = row.createCell(column);
                 cell.setCellValue((String) "");  
                 rowCount++;
                    
                column++;
                if(column==2) //extra space between 0,1 and 3,4
                    column++;
          }
          
          if(column > 4) column = 0;  
        }
         
        sheet.setColumnWidth(0,3500); //logo
        sheet.setColumnWidth(1,7000); //text
        sheet.setColumnWidth(2,1400); //space
        sheet.setColumnWidth(3,3500); // logo2
        sheet.setColumnWidth(4,7000); //text 2
            
        try (FileOutputStream outputStream = new FileOutputStream(FileNameLabels)) {
            workbook.write(outputStream);
        }
    }
    
  //ayj logo
    private void addLogo (XSSFWorkbook wb,XSSFSheet sheet,int row, int column) throws FileNotFoundException, IOException
    {
        int pictureIdx;
        //Get the contents of an InputStream as a byte[].
        try (InputStream inputStream = new FileInputStream("xlsx/ayj.jpg")) {
            //Get the contents of an InputStream as a byte[].
            byte[] bytes = IOUtils.toByteArray(inputStream);
            //Adds a picture to the workbook
            pictureIdx = wb.addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
            //close the input stream
        }
           //Returns an object that handles instantiating concrete classes
           CreationHelper helper = wb.getCreationHelper();
           //Creates the top-level drawing patriarch.
           Drawing drawing = sheet.createDrawingPatriarch();

           //Create an anchor that is attached to the worksheet
           ClientAnchor anchor = helper.createClientAnchor();

           //create an anchor with upper left cell _and_ bottom right cell
           anchor.setCol1(column); 
           anchor.setRow1(row+1); 
           anchor.setCol2(column+1); 
           anchor.setRow2(row+5); 

           //Creates a picture
           Picture pict = drawing.createPicture(anchor, pictureIdx);
    }

    private Response findResponseById(int id)
    {
        for(Response response : TemporaryResponseList){    
           if(response.ID == id )
               return response;   
        } 
        return null;
    }
    
}
