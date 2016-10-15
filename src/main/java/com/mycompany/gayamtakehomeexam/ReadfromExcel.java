/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.gayamtakehomeexam;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author S525719
 */
public class ReadfromExcel {
    /*
 * specifying the input file path  
 */
    private static final String FILE_PATH = "C:\\Users\\S525719\\Documents\\NetBeansProjects\\GayamTakeHomeExam\\Gayam_Input.xlsx";

   
/*
    * method to get list of movies from the input excel file
*/
    public  List getSongsListFromExcel() {
        List songList = new ArrayList();
        FileInputStream fis = null;
         
        try {
            fis = new FileInputStream(FILE_PATH);

            
            /*
              Use XSSF for xlsx format, for xls use HSSF
            */
            Workbook workbook = new XSSFWorkbook(fis);

            int numberOfSheets = workbook.getNumberOfSheets();

          
            
            /*
            looping over each workbook sheet
            */
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                Iterator rowIterator = sheet.iterator();

                
            /*
                iterating over each row
                */
                while (rowIterator.hasNext()) {

                    Song song = new Song();
                    Row row = (Row) rowIterator.next();
                    
                    
                  
                         Iterator cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                       
                        Cell cell = (Cell) cellIterator.next();
                     
                   
                        /*
                        checking if the cell is having a String value .
                        */
                        if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                           
                            
                            /*
                            Cell with index 1 contains Album name 
                            */
                            if (cell.getColumnIndex() == 1) {
                                 song.setAlbumname(cell.getStringCellValue());
                            }
                            
                            
                            /*
                            Cell with index 2 contains Genre
                            */
                            if (cell.getColumnIndex() == 2) {
                                 song.setGenre(cell.getStringCellValue());
                            }
                            
                            
                            /*
                            Cell with index 3 contains Artist name
                            */
                            if (cell.getColumnIndex() == 3) {
                                 song.setArtist(cell.getStringCellValue());
                            }
                     
                        }
                        
                        /*
                         checking if the cell is having a numeric value
                        */
                        else if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {

                            
                            /*
                            Cell with index 0 contains Sno
                            */
                            if (cell.getColumnIndex() == 0) {
                                song.setSno((int) cell.getNumericCellValue());
                            }
                            
                            /*
                            Cell with index 5 contains Critic score.
                            */
                            else if (cell.getColumnIndex() == 5) {
                                song.setCriticscore((int)cell.getNumericCellValue());
                            }
                            
                            /*
                            Cell with index 4 contains Release date
                            */
                           else if (cell.getColumnIndex() == 4) {
                               Date dateValue = null;
                           
                                 if (DateUtil.isCellDateFormatted(cell)) {
                                 dateValue = cell.getDateCellValue();
                                 }
                                 song.setReleasedate(dateValue);
                            }
               

                        }
                    
                    }
                    
                    /*
                    end iterating a row, add all the elements of a row in list
                    */
                    songList.add(song);
                }
            }

            fis.close();

        } catch (FileNotFoundException e) {
        } catch (IOException e) {
        }
        return songList;
    }
}
