package com.mycompany.gayamtakehomeexam;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
 
/**
 *
 * @author S525719
 */
public class WritetoExcel {
     
/*
    specifying the input file path
*/
    private static final String FILE_PATH = "Gayam_Output.xlsx";
   
 
 
/*
    We are making use of a single instance to prevent multiple write access to same file.
    */
    private static final WritetoExcel INSTANCE = new WritetoExcel();
 
    public static WritetoExcel getInstance() {
        return INSTANCE;
    }
 
 /**
 * No ArgsContructor with no body defined
 */
   
    public WritetoExcel() {
    }
 
 
   
/*
    create a method to write data to excel file
    */
    public  void writeSongsListToExcel(List<Song> songList){
 
       
        /*
        Use XSSF for xlsx format and for xls use HSSF
        */
        Workbook workbook = new XSSFWorkbook();
       
       
        /*
        create new sheet
        */
        Sheet songsSheet = workbook.createSheet("Albums");
       
        XSSFCellStyle my_style = (XSSFCellStyle) workbook.createCellStyle();    
        /* Create XSSFFont object from the workbook */
        XSSFFont my_font=(XSSFFont) workbook.createFont();
       
       
        /*
        setting cell color
        */
        CellStyle style = workbook.createCellStyle();
    style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
           
    CellStyle style1 = workbook.createCellStyle();//Create style
    Font font = workbook.createFont();//Create font
    font.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
    style1.setFont(font);//set it to bold
    style1.setAlignment(CellStyle.ALIGN_CENTER);
 
    CellStyle style13 = workbook.createCellStyle();//Create style
    Font font13 = workbook.createFont();//Create font
    style13.setAlignment(CellStyle.ALIGN_LEFT);
    
    CellStyle style14 = workbook.createCellStyle();//Create style
    Font font14 = workbook.createFont();//Create font
    style14.setAlignment(CellStyle.ALIGN_RIGHT);
        /*
         setting Header color
        */
        CellStyle style2 = workbook.createCellStyle();
    style2.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        style2.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
    style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
           
       
        Row rowName = songsSheet.createRow(1);
       
        /*
        Merging the cells
        */
        songsSheet.addMergedRegion(new CellRangeAddress(1, 1, 2,3 ));
     
       
        /*
        Applying style to attribute name
        */
        int nameCellIndex = 1;
        Cell namecell =  rowName.createCell(nameCellIndex++);
        namecell.setCellValue("Name");
        namecell.setCellStyle(style);
       
       
       Cell cel = rowName.createCell(nameCellIndex++);
       
       /*
       Applying underline to Name
       */
         Font underlineFont = workbook.createFont();
          underlineFont.setUnderline(HSSFFont.U_SINGLE);
         
        /* Attaching the style to the cell */
        CellStyle combined = workbook.createCellStyle();
        combined.setFont(underlineFont);
        cel.setCellStyle(combined);
        cel.setCellValue("Gayam, Prathibha");
       
        /*
        Applying  colors to header
        */
        Row rowMain = songsSheet.createRow(3);
        SheetConditionalFormatting sheetCF = songsSheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("3");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
       
        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A4:G4")
        };
       
 
        sheetCF.addConditionalFormatting(regions, rule1);
       
       
        /*
        setting new rule to apply alternate colors to cells having same Genre
        */
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule("4");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.LEMON_CHIFFON.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
 
         CellRangeAddress[] regionsAction = {
            CellRangeAddress.valueOf("A5:G5"),
            CellRangeAddress.valueOf("A6:G6"),
            CellRangeAddress.valueOf("A7:G7"),
            CellRangeAddress.valueOf("A8:G8"),
            CellRangeAddress.valueOf("A13:G13"),
            CellRangeAddress.valueOf("A14:G14"),
            CellRangeAddress.valueOf("A15:G15"),
            CellRangeAddress.valueOf("A16:G16"),
            CellRangeAddress.valueOf("A23:G23"),
            CellRangeAddress.valueOf("A24:G24"),
            CellRangeAddress.valueOf("A25:G25"),
            CellRangeAddress.valueOf("A26:G26")
               
        };
         
         
        /*        
        setting new rule to apply alternate colors to cells having same Genre
         */
        ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule("4");
        PatternFormatting fill3 = rule3.createPatternFormatting();
        fill3.setFillBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.index);
        fill3.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
 
         CellRangeAddress[] regionsAdv = {
             CellRangeAddress.valueOf("A9:G9"),
             CellRangeAddress.valueOf("A10:G10"),
             CellRangeAddress.valueOf("A11:G11"),
             CellRangeAddress.valueOf("A12:G12"),
             CellRangeAddress.valueOf("A17:G17"),
             CellRangeAddress.valueOf("A18:G18"),
             CellRangeAddress.valueOf("A19:G19"),
             CellRangeAddress.valueOf("A20:G20"),
             CellRangeAddress.valueOf("A21:G21"),
             CellRangeAddress.valueOf("A22:G22"),          
             CellRangeAddress.valueOf("A27:G27"),
             CellRangeAddress.valueOf("A28:G28"),
             CellRangeAddress.valueOf("A29:G29")        
        };
       
 
     
        /*
        Applying above created rule formatting to cells
        */
        sheetCF.addConditionalFormatting(regionsAction, rule2);
        sheetCF.addConditionalFormatting(regionsAdv, rule3);
       
       
     
        /*
         Setting coloumn header values
        */
        int mainCellIndex = 0;
     
        Cell cell0 = rowMain.createCell(mainCellIndex++);
        cell0.setCellValue("SNO");
        cell0.setCellStyle(style1);
        Cell cell1 = rowMain.createCell(mainCellIndex++);
        cell1.setCellValue("Genre");
        cell1.setCellStyle(style1);
        Cell cell2 = rowMain.createCell(mainCellIndex++);
        cell2.setCellValue("Credit Score");
        cell2.setCellStyle(style1);
        Cell cell3 = rowMain.createCell(mainCellIndex++);
        cell3.setCellValue("Album Name");
        cell3.setCellStyle(style1);
        Cell cell4 = rowMain.createCell(mainCellIndex++);
        cell4.setCellValue("Artist");
        cell4.setCellStyle(style1);
        Cell cell5 = rowMain.createCell(mainCellIndex++);
        cell5.setCellValue("Release Date");
        cell5.setCellStyle(style1);
       
       
       
       
       
        /*
        populating cell values
        */
        int rowIndex = 4;
        int sno = 1;
        for(Song song : songList){
            if(song.getSno() != 0){
 
            Row row = songsSheet.createRow(rowIndex++);
            int cellIndex = 0;
           
           
            /*
            first place in row is Sno
            */
            
            Cell cell20 = row.createCell(cellIndex++);
            cell20.setCellValue(sno++);
            cell20.setCellStyle(style14);
 
           
            /*
            second place in row is  Genre
            */
            Cell cell21 = row.createCell(cellIndex++);
            cell21.setCellValue(song.getGenre());
            cell21.setCellStyle(style13);
                   
 
           
            /*
            third place in row is Critic score
            */
            Cell cell22 = row.createCell(cellIndex++);
            cell22.setCellValue(song.getCriticscore());
            cell22.setCellStyle(style14);
 
           
            /*
            fourth place in row is Album name
            */
            Cell cell23 = row.createCell(cellIndex++);
            cell23.setCellValue(song.getAlbumname());
            cell23.setCellStyle(style13);
           
           
            /*
            fifth place in row is Artist
            */
            Cell cell24 = row.createCell(cellIndex++);
            cell24.setCellValue(song.getArtist());
            cell24.setCellStyle(style13);
           
           
            /*
            sixth place in row is marks in date
            */
            if (song.getReleasedate() != null){
               
                Cell date = row.createCell(cellIndex++);
               
                DataFormat format = workbook.createDataFormat();
                CellStyle dateStyle = workbook.createCellStyle();
                dateStyle.setDataFormat(format.getFormat("dd-MMM-yyyy"));
                date.setCellStyle(dateStyle);  
 
           
            date.setCellValue(song.getReleasedate());
           
           
            /*
            auto-resizing columns
            */
            songsSheet.autoSizeColumn(6);
            songsSheet.autoSizeColumn(5);
            songsSheet.autoSizeColumn(4);
            songsSheet.autoSizeColumn(3);
            songsSheet.autoSizeColumn(2);
            }
           
 
        }
    }
       
        /*
        writing this workbook to excel file.
        */
        try {
            FileOutputStream fos = new FileOutputStream(FILE_PATH);
            workbook.write(fos);
            fos.close();
 
            System.out.println(FILE_PATH + " is successfully written");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
 
       
    }
}