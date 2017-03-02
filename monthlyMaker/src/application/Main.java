package application;
	
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.BorderPane;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;


public class Main extends Application {
	static File selectedFile;
	WeekRange[] wR = new WeekRange[53];
	@Override
	public void start(Stage primaryStage) {
		try {
			Button file = new Button("Choose File");
			file.setOnAction(new EventHandler<ActionEvent>(){

				@SuppressWarnings("deprecation")
				@Override
				public void handle(ActionEvent arg0) {
					
					FileChooser fileChooser = new FileChooser();
					fileChooser.setTitle("Open Resource File");
					fileChooser.getExtensionFilters().addAll(
			         new ExtensionFilter("Excell Files", "*.xlsx"));
					 selectedFile = fileChooser.showOpenDialog(primaryStage);
				        if(selectedFile == null){
				        	
				        }
				        else{
					 FileInputStream inputStream;
					 FileOutputStream outputStream;

						try {
								inputStream = new FileInputStream(selectedFile);
								File outputFile = new File(selectedFile.getPath().substring(0, selectedFile.getPath().lastIndexOf("/")) + "/Final_Books" + LocalDate.now() + ".xlsx");
								outputStream = new FileOutputStream(outputFile);
								
				         
				        Workbook wb;
				        Workbook wbo;

							wb = new XSSFWorkbook(inputStream);
							wbo = new XSSFWorkbook();
							CreationHelper createHelper = wbo.getCreationHelper();
							Sheet sheetOut = wbo.createSheet(selectedFile.getName());
							Row rowOut = sheetOut.createRow(0);
							Cell cellOut;
							cellOut = rowOut.createCell(0);
							cellOut.setCellValue(createHelper.createRichTextString("Week"));
							cellOut = rowOut.createCell(1);
							cellOut.setCellValue(createHelper.createRichTextString("Behavior/Decel"));
							cellOut = rowOut.createCell(2);
							cellOut.setCellValue(createHelper.createRichTextString("Data Input Total"));
							cellOut = rowOut.createCell(3);
							cellOut.setCellValue(createHelper.createRichTextString("Measurment Type"));
							cellOut = rowOut.createCell(4);
							cellOut.setCellValue(createHelper.createRichTextString("Measurment Unit"));
							int rowCount = 1;
				        for (int k = 0; k < wb.getNumberOfSheets(); k++) {
							Sheet sheet = wb.getSheetAt(k);
							int weekCount = 0;
							double weekTotal = 0;
							boolean replace = false;
							String decelName = null;
							String replaceName = null;
							String measType = null;
							String measUnit = null;
							String replacedName = null;
							WeekRange wr = null; 
							int weekNum = -1; 
							int lastWeek = -1;
							int year = 0;
							
							int rows = sheet.getPhysicalNumberOfRows();
							//System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows
									//+ " row(s).");
							for (int r = 0; r < rows; r++) {
								boolean newWeek = false;
								Row row = sheet.getRow(r);
								if (row == null) {
									continue;
								}

								int cells = row.getPhysicalNumberOfCells();
								//System.out.println("\nROW " + row.getRowNum() + " has " + cells
										//+ " cell(s).");
								for (int c = 0; c < cells; c++) {
									Cell cell = row.getCell(c);
									if(cell == null)
									{
										//c++;
										cells++;
									}
									else{
									 switch (cell.getCellTypeEnum()) {
						                case STRING:
						                	if(r == 0 && cell.getRichStringCellValue().getString().contains("Replacement"))
						                		replace = true;
						                	if(r == 0 && replace == true && c == 1)
						                		replaceName = cell.getRichStringCellValue().getString();
						                	else if(r == 0 && replace == false && c == 1)
						                		decelName = cell.getRichStringCellValue().getString();
						                	else if(r == 0 && replace == true && c == 2)
						                		replacedName = cell.getRichStringCellValue().getString();
						                	if(r == 1 && c == 0)
						                		measType = cell.getRichStringCellValue().getString();
						                	if(r == 1 && c == 1)
						                		measUnit = cell.getRichStringCellValue().getString();
						                    //System.out.println(cell.getRichStringCellValue().getString());
						                    break;
						                case NUMERIC:
						                    if (DateUtil.isCellDateFormatted(cell)) {
						                    	if(year == 0){
						                    		year = cell.getDateCellValue().getYear();
						                    		findWeeks(year);
						                    	}
						                    	else if(year != cell.getDateCellValue().getYear()){
						                    		year = cell.getDateCellValue().getYear();
						                    		findWeeks(year);
						                    	}
						                    	lastWeek = weekNum;
						                    	for(int i = 0; i < 51; i++){
						                    		if(cell.getDateCellValue().after(wR[i].startWeek) && cell.getDateCellValue().before(wR[i].endWeek) || cell.getDateCellValue().equals(wR[i].startWeek) || cell.getDateCellValue().equals(wR[i].endWeek)){
						                    			weekNum = i;
						                    		}
						                    	}
						                    	
						                    		
						                        //System.out.println(cell.getDateCellValue());
						                     } else {
						                        //System.out.println(cell.getNumericCellValue());
						                    	 if(lastWeek != weekNum && lastWeek != -1)
						                    		 r--;
						                    	 else
						                    		 weekTotal += cell.getNumericCellValue();
						                    }
						                    break;
						                case BOOLEAN:
						                    //System.out.println(cell.getBooleanCellValue());
						                    break;
						                case FORMULA:
						                    //System.out.println(cell.getCellFormula());
						                    break;
						                case BLANK:
						                    //System.out.println();
						                    break;
						                default:
						                    //System.out.println();
						            }
									
									//System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE="
											//+ value);
									}
								}//System.out.println(lastWeek + " " + weekNum + " " + weekTotal);
								String[] months = {"january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"};
								if(lastWeek != weekNum && lastWeek != -1)
									{

										//weekCount ++;
										rowOut	= sheetOut.createRow(rowCount);
										Cell cell = rowOut.createCell(0);
										cell.setCellValue(createHelper.createRichTextString(months[wR[weekNum].month-1] + " week " + wR[weekNum].monthWeek + ", " + (year+1900)));
										cell = rowOut.createCell(1);
										if(replace == true)
											cell.setCellValue(createHelper.createRichTextString(replaceName));
										else
											cell.setCellValue(createHelper.createRichTextString(decelName));
										cell = rowOut.createCell(2);
										cell.setCellValue(weekTotal);
										cell = rowOut.createCell(3);
										cell.setCellValue(createHelper.createRichTextString(measType));
										cell = rowOut.createCell(4);
										cell.setCellValue(createHelper.createRichTextString(measUnit));
										System.out.println(months[wR[weekNum].month-1] + " week " + (wR[weekNum].monthWeek) + ", " + (year+1900) + ": " + weekTotal + " Measurment Type: " + measType + " Measument Unit: " + measUnit);
										weekTotal = 0;
										lastWeek = weekNum;
										//newWeek = false;
										rowCount++;
									}
								
							}
							if(replace == true)
									System.out.println(replaceName + " for " + replacedName);
								else
									System.out.println(decelName);
						}
				        wbo.write(outputStream);
				        wbo.close();
				        wb.close();
				        outputStream.close();
				        inputStream.close();
							} catch (FileNotFoundException e) {
								e.printStackTrace();
							} catch (IOException e) {
								e.printStackTrace();
							}
				}
				}
			});
			
			 
			BorderPane root = new BorderPane();
			root.setCenter(file);
			Scene scene = new Scene(root,400,400);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.show();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) throws IOException {
		launch(args);
		
    }
	
	@SuppressWarnings("deprecation")
	public void findWeeks(int year){
		int dow = 0; 
		int w = 0;
		boolean newMonth = true;
		int mw = 1;
		Date currentDate;
		for(int i = 1; i < 13; i++){
			for(int d = 1; d < LocalDate.of(year, i, 1).lengthOfMonth(); d++){
				currentDate = new Date(year, i-1, d);
				dow = currentDate.getDay();
				
				if(d == 1){
					newMonth = true;
				}
				if(dow == 0 || d == 1){
					wR[w] = new WeekRange();
					wR[w].startWeek = currentDate;
				}
				if(dow == 6 || d == LocalDate.of(year, i, 1).lengthOfMonth()){
					wR[w].endWeek = currentDate;
					if(newMonth == true){
						newMonth = false;
						mw = 1;
					}
					wR[w].monthWeek = mw;
					wR[w].month = i;
					mw++;
					w++;
				}
				dow++; 
			}
		}
		/*for(int j = 0; j < 53; j++){
			if(wR[j] != null)
				System.out.println(wR[j].startWeek + " - " + wR[j].endWeek);
		}*/
	}
	
}
