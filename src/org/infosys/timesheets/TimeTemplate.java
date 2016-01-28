/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.infosys.timesheets;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.TimeZone;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 *
 * @author 58128
 */
public class TimeTemplate {
    
    private static final String template = "timesheet_template.xls";
    private static final String outFilePath = "Jaroslaw Smorczewski_#MONTH#_Timesheet.xls";
    
    private static final Map<String, String> publicHolidays2015 = new HashMap();
    static {
        publicHolidays2015.put("25-Dec", "Christmas Day");
        publicHolidays2015.put("26-Dec", "St Stephens Day");
    } 
    
    private static final Map<String, String> publicHolidays2016 = new HashMap();
    static {
        publicHolidays2016.put("01-Jan", "New Year Eve");
        publicHolidays2016.put("17-Mar", "St Patrics Day");
        publicHolidays2016.put("25-Mar", "Good Friday");
        publicHolidays2016.put("28-Mar", "Easter Monday");
        publicHolidays2016.put("02-May", "May Bank Holiday");
        publicHolidays2016.put("06-Jun", "June Bank Holiday");
        publicHolidays2016.put("01-Aug", "August Bank Holiday");
        publicHolidays2016.put("31-Oct", "October Bank Holiday");
        publicHolidays2016.put("26-Dec", "St Stephens Day");
        publicHolidays2016.put("27-Dec", "In lieu of Christmas Day");
    }
    
    private String monthName;
    
    public void prepareTimesheet(String[] args) {
        try {
            String[] ym = args[0].split("/");
            int month = Integer.parseInt(ym[0]);
            int year = Integer.parseInt(ym[1]);
            
            Calendar cal = Calendar.getInstance(TimeZone.getDefault());
            cal.set(Calendar.YEAR, year);
            cal.set(Calendar.MONTH, month-1);
            int days = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
            monthName = cal.getDisplayName(Calendar.MONTH, Calendar.SHORT_FORMAT, Locale.ENGLISH);
            String periodName = monthName+"-"+year;
            
            System.out.println("Month: "+periodName);
            System.out.println("Days in month: "+days);
            //System.out.println(" ");
            
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(template));
            CellStyle cs = wb.createCellStyle();
            cs.setFillPattern(CellStyle.ALIGN_FILL);
            cs.setFillBackgroundColor(HSSFColor.DARK_GREEN.index);
            HSSFSheet sheet = wb.getSheetAt(0);
            
            Map<String, String> bankHolidays = year==2016 ? publicHolidays2016 : publicHolidays2015;
            Map<String, String> holidays = this.extractHolidays(args);
            
            HSSFRow labelRow = sheet.getRow(13);
            HSSFRow valueRow = sheet.getRow(14);
            HSSFRow leaveRow = sheet.getRow(15);
            
            int i = 1;
            while (i<=days) {                
                cal.set(Calendar.DAY_OF_MONTH, i);
                int weekDay = cal.get(Calendar.DAY_OF_WEEK);
                String dayOfMonth =  StringUtils.leftPad(String.valueOf(i), 2, '0').concat("-").concat(monthName);
                
                HSSFCell cell = labelRow.getCell(i+1);
                cell.setCellValue(dayOfMonth);
                
                if (weekDay==1 || weekDay==7 || bankHolidays.get(dayOfMonth)!=null) {
                    valueRow.getCell(i+1).setCellStyle(cs);
                    leaveRow.getCell(i+1).setCellStyle(cs);
                }
                else {
                    if (holidays.get(dayOfMonth)!=null)
                        leaveRow.getCell(i+1).setCellValue(8);
                    else
                        valueRow.getCell(i+1).setCellValue(8);
                }
                i++;
            }
            // refresh summaries
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(valueRow.getCell(33));
            evaluator.evaluateFormulaCell(leaveRow.getCell(33));                        
            evaluator.evaluateFormulaCell(sheet.getRow(18).getCell(33));
            
            wb.write(new FileOutputStream(outFilePath.replace("#MONTH#", periodName)));
            
        } catch (IOException ex) {
            Logger.getLogger(Timesheets.class.getName()).log(Level.SEVERE, null, ex);
        }
        System.out.println("Timesheet created.");
    }
    
    private Map extractHolidays(String[] args) {
        Map<String, String> holidays = new HashMap();
        int j=1;
        while (j<args.length) {
            String[] period = args[j].split("-");
            int start = Integer.parseInt(period[0]);
            int end = period.length>1 ? Integer.parseInt(period[1]) : 0;
            if (end>0) {
                for (int i=start; i<=end; i++) {
                    String dayOfMonth =  StringUtils.leftPad(String.valueOf(i), 2, '0').concat("-").concat(monthName);
                    holidays.put(dayOfMonth, "Leave");
                }
            } else {
                String dayOfMonth =  StringUtils.leftPad(String.valueOf(start), 2, '0').concat("-").concat(monthName);
                holidays.put(dayOfMonth, "Leave");
            }
            j++;
        }
        return holidays;
    }
}
