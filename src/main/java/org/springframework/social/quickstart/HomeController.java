/*
 * Copyright 2011 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.springframework.social.quickstart;

import static org.springframework.util.StringUtils.hasText;
import static org.springframework.web.bind.annotation.RequestMethod.GET;
import static org.springframework.web.bind.annotation.RequestMethod.POST;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletResponse;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;
import org.codehaus.jackson.node.ObjectNode;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.social.ExpiredAuthorizationException;
import org.springframework.social.google.api.Google;
import org.springframework.social.google.api.drive.DriveFile;
import org.springframework.social.google.api.drive.DriveFileQueryBuilder;
import org.springframework.social.google.api.drive.DriveFilesPage;
import org.springframework.social.google.api.drive.DriveOperations;
import org.springframework.social.google.api.drive.FileProperty;
import org.springframework.social.google.api.drive.PropertyVisibility;
import org.springframework.social.quickstart.drive.DateOperators;
import org.springframework.social.quickstart.drive.DriveSearchForm;
import org.springframework.social.quickstart.drive.OptionalBoolean;
import org.springframework.social.quickstart.drive.WorksheetForm;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class HomeController {

	private final Google google;

	private static final int FORMULA_REMOVAL_START_COL = 11;
	private static final int FORMULA_REMOVAL_END_COL = 15;

	private static final double HRM_VERSION = 11.0;
	private static final String HRM_TYPE_HASLER = "HRM";
	private static final String HRM_TYPE_NATIONALS = "NRM";
	private static final String HRM_TYPE_ASSESSMENT = "ARM";

	private static final String PROP_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

	private static final String SHEET_FINISHES = "Finishes";
	private static final String SHEET_STARTS = "Starts";
	private static final String SHEET_CLUBS = "Clubs";
	private static final String SHEET_RESULTS = "Results";
	private static final String SHEET_SUMMARY = "Summary";

	private static Pattern numberPattern = Pattern.compile("\\d+");
	private static Pattern timePattern = Pattern.compile("(\\d{1,2}):(\\d{2}):(\\d{2})");

	@Autowired
	public HomeController(Google google) {
		this.google = google;
	}
	
	@ExceptionHandler(ExpiredAuthorizationException.class)
	public String handleExpiredToken() {
		return "redirect:/signout";
	}
	
	@ExceptionHandler(Exception.class)
	public void handleException(Exception e) {
		e.printStackTrace();
	}

	@RequestMapping(value="/", method=GET)
	public String home(DriveSearchForm command) {
		return "redirect:/hrm";
	}

	@RequestMapping(value="/hrm", method=GET)
	public ModelAndView getHRMFiles(DriveSearchForm command) {
		return getDriveFiles(HRM_TYPE_HASLER, command);
	}
	
	@RequestMapping(value="/arm", method=GET)
	public ModelAndView getARMFiles(DriveSearchForm command) {
		return getDriveFiles(HRM_TYPE_ASSESSMENT, command);
	}
	
	@RequestMapping(value="/nrm", method=GET)
	public ModelAndView getNRMFiles(DriveSearchForm command) {
		return getDriveFiles(HRM_TYPE_NATIONALS, command);
	}

	private ModelAndView getDriveFiles(String hrmType, DriveSearchForm command) {

		DriveFileQueryBuilder queryBuilder = google.driveOperations().driveFileQuery()
				.fromPage(command.getPageToken());

		queryBuilder.maxResultsNumber(10);
		queryBuilder.trashed(false);
		queryBuilder.mimeTypeIs("application/vnd.google-apps.spreadsheet");
		//queryBuilder.parentIs("appdata");
		
		if(hasText(command.getTitleContains())) {
			queryBuilder.titleContains(command.getTitleContains());
		}
		
		queryBuilder.propertiesHas("hrmType", hrmType, PropertyVisibility.PUBLIC);

		DriveFilesPage files = queryBuilder.getPage();

		Map<DateOperators, String> dateOperators = new LinkedHashMap<DateOperators, String>();
		for(DateOperators operator : DateOperators.values()) {
			dateOperators.put(operator, operator.toString());
		}
		
		Map<OptionalBoolean, String> booleanOperators = new LinkedHashMap<OptionalBoolean, String>();
		for(OptionalBoolean operator : OptionalBoolean.values()) {
			booleanOperators.put(operator, operator.toString());
		}
		
		return new ModelAndView("drivefiles")
			.addObject("dateOperators", dateOperators)
			.addObject("booleanOperators", booleanOperators)
			.addObject("command", command)
			.addObject("files", files)
			.addObject("selected", hrmType.toLowerCase()) // Ensures nav context is shown correctly
			.addObject("hrmType", hrmType.toUpperCase()); // For labels
	}
	
	@RequestMapping(value="{hrmType}/new", method=GET)
	public ModelAndView createWorksheet(@PathVariable String hrmType) {
		
		WorksheetForm form = new WorksheetForm();
		form.setType(hrmType);
		return new ModelAndView("worksheet", "command", form)
			.addObject("selected", hrmType.toLowerCase()) // Ensures nav context is shown correctly
			.addObject("hrmType", hrmType.toUpperCase()); // For labels
	}
	
	@RequestMapping(value="workbook", method=GET, params="id")
	public ModelAndView task(String id) {
		
		DriveFile file = google.driveOperations().getFile(id);
		WorksheetForm command = new WorksheetForm(file.getId(), file.getTitle());
		return new ModelAndView("task", "command", command);
	}
	
	@RequestMapping(value="{hrmType}/new", method=POST)
	public ModelAndView saveWorksheet(WorksheetForm command, BindingResult result) {
		
		if(result.hasErrors()) {
			return new ModelAndView("worksheet", "command", command);
		}

		String[] parents = { "root" };
		String srcId = null;
		String raceType = command.getType().toUpperCase();
		if (raceType.equals(HRM_TYPE_HASLER)) {
			srcId = "1V7FTXOszEFtA23BAmYBYekbRHAA3gaz3sZPLaJfacnw";
		} else if (raceType.equals(HRM_TYPE_ASSESSMENT)) {
			srcId = "10JNPb7LA0QIERO93JPgxsSn129dR-AddsAnHAx2iBXM";
		} else if (raceType.equals(HRM_TYPE_NATIONALS)) {
			srcId = "1A9USkMFEZtJL2KmsdFljk4eHag07yPvmZq3oInCYPo0";
		}
		DriveOperations driveOperations = google.driveOperations();
		DriveFile file = driveOperations.copy(srcId, parents, command.getTitle());
		// Record HRM type in a custom property
		driveOperations.getProperties(file.getId());
		FileProperty fp = new FileProperty("hrmType", raceType, PropertyVisibility.PUBLIC);
		driveOperations.addProperty(file.getId(), fp);
		return new ModelAndView("redirect:/", "list", command.getList());
	}
	
	/*
	@RequestMapping(value="workbook", method=POST, params="parent")
	public ModelAndView createTask(String parent, String previous, WorksheetForm command, BindingResult result) {
		
		if(result.hasErrors()) {
			return new ModelAndView("worksheet", "command", command);
		}

		Task task = new Task(command.getId(), command.getTitle(), command.getNotes(), command.getDue(), command.getCompleted());
		google.taskOperations().createTaskAt(command.getList(), parent, previous, task);
		
		return new ModelAndView("redirect:/tasks", "list", command.getList());
	}
	*/
	
	@RequestMapping(value="starfile", method=POST)
	@ResponseBody
	public void starFile(String fileId, boolean star) {
		DriveOperations drive = google.driveOperations();
		if(star) {
			drive.star(fileId);
		} else {
			drive.unstar(fileId);
		}
	}
	
	@RequestMapping(value="trashfile", method=POST)
	@ResponseBody
	public void trashFile(String fileId, boolean trash) {
		DriveOperations drive = google.driveOperations();
		if(trash) {
			drive.trash(fileId);
		} else {
			drive.untrash(fileId);
		}
	}
	
	@RequestMapping(value="deletefile", method=POST)
	@ResponseBody
	public void deleteFile(String fileId) {
		google.driveOperations().delete(fileId);
	}
	
	@RequestMapping(value="copyfile", method=POST, produces="application/json")
	public void copyFile(String fileId, String parentId, String newName, HttpServletResponse response) throws JsonGenerationException, JsonMappingException, IOException {
		Writer writer = response.getWriter();
		DriveOperations driveOperations = google.driveOperations();
		DriveFile file = driveOperations.copy(fileId, new String[]{parentId}, newName);

		// Copy HRM type if defined in custom properties
		List<FileProperty> fps = driveOperations.getProperties(fileId);
		for (FileProperty fp : fps) {
			if (fp.getKey().equals("hrmType")) {
				driveOperations.addProperty(file.getId(), fp);
			}
		}

		// Return JSON response
		ObjectMapper mapper = new ObjectMapper();
		JsonNode rootNode = mapper.createObjectNode(); // will be of type ObjectNode
		((ObjectNode) rootNode).put("id", file.getId());
		((ObjectNode) rootNode).put("title", file.getTitle());
		response.setContentType("application/json");
		mapper.writeValue(writer, rootNode);
	}
	
	@RequestMapping(value="downloadfile/*", method=GET, params="fileId")
	public void downloadFile(String fileId, HttpServletResponse response) throws Exception {
		DriveFile file = google.driveOperations().getFile(fileId);
		if (file == null) {
			throw new Exception("File not found");
		}
		String exportUri = file.getExportLinks().get("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		if (exportUri != null) {

			// Download XLSX file from Google

			HttpClient client = new DefaultHttpClient();
			HttpGet httpget = new HttpGet(exportUri);
			httpget.setHeader("Authorization", "Bearer " + google.getAccessToken());
			HttpResponse resp = client.execute(httpget);
			HttpEntity entity = resp.getEntity();
			if (entity != null) {
				InputStream instream = entity.getContent();
				final File temp = File.createTempFile("hrmdownload-", ".xlsx");
				OutputStream tos = new FileOutputStream(temp);
				try {
					byte[] buffer = new byte[1024]; // Adjust if you want
					int bytesRead;
					while ((bytesRead = instream.read(buffer)) != -1)
					{
						tos.write(buffer, 0, bytesRead);
					}
				} finally {
					instream.close();
				}
				tos.close();

				String fileType = null;
				double hrmVersion = HRM_VERSION;

				// Modify the .xlsx file using POI
				// TODO Refactor this all into a helper class

				OPCPackage pkg = OPCPackage.open(temp);
				XSSFWorkbook wb = new XSSFWorkbook(pkg);
				CreationHelper createHelper = wb.getCreationHelper();

				int numSheets = wb.getNumberOfSheets();

				// Auto-detect the sheet type based on the first sheet name
				if (numSheets > 0) {
					String firstSheetName = wb.getSheetAt(0).getSheetName();
					if (firstSheetName.equals("Div1")) {
						fileType = HRM_TYPE_HASLER;
					} else if (firstSheetName.equals("SMK1")) {
						fileType = HRM_TYPE_ASSESSMENT;
					} else if (firstSheetName.equals("Div7") || firstSheetName.equals("U12 M")) {
						fileType = HRM_TYPE_NATIONALS;
					}
				}

				// Write properties
				POIXMLProperties props = wb.getProperties();
				POIXMLProperties.CustomProperties cust =  props.getCustomProperties();

				if (fileType != null) {
					CTProperty hrmProperty = cust.getUnderlyingProperties().addNewProperty();
					hrmProperty.setBool(true);
					hrmProperty.setName(fileType);
					hrmProperty.setFmtid(PROP_FMTID);
					hrmProperty.setPid(2);
				}

				CTProperty versionProperty = cust.getUnderlyingProperties().addNewProperty();
				versionProperty.setName("Version");
				versionProperty.setR8(hrmVersion);
				versionProperty.setFmtid(PROP_FMTID);
				versionProperty.setPid(3);

				String sheetPassword = fileType;
				XSSFSheet sheet;
				String sheetName;

				Font font = wb.createFont();
				font.setFontHeightInPoints((short)10);
				font.setFontName("Courier New");
				Font boldFont = wb.createFont();
				boldFont.setFontHeightInPoints((short)10);
				boldFont.setFontName("Courier New");
				boldFont.setBold(true);

				XSSFCellStyle bodyStyle = wb.createCellStyle();
				bodyStyle.setFont(font);

				XSSFCellStyle firstColumnStyle = wb.createCellStyle();
				firstColumnStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 255, 153)));
				firstColumnStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				firstColumnStyle.setBorderRight(CellStyle.BORDER_THIN);
				firstColumnStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
				firstColumnStyle.setFont(boldFont);

				CellStyle dateStyle = (CellStyle) bodyStyle.clone();
				short dateFormat = createHelper.createDataFormat().getFormat("dd/mm/yy");
				dateStyle.setDataFormat(dateFormat);

				CellStyle timeStyle = (CellStyle) bodyStyle.clone();
				short timeFormat = createHelper.createDataFormat().getFormat("H:MM:SS");
				timeStyle.setDataFormat(timeFormat);

				List<String> raceSheetNames = new ArrayList<String>();
				List<String> columnsToRemove = new ArrayList<String>();
				columnsToRemove.add("Due");
				// Go through sheets
				boolean isRaceSheet = true;
				for (int i = 0; i < numSheets; i++) {
					sheet = wb.getSheetAt(i);
					sheetName = sheet.getSheetName();
					isRaceSheet = isRaceSheet && !sheetName.equals(SHEET_FINISHES);
					if (isRaceSheet) {
						raceSheetNames.add(sheetName);
					}
					// Set sheet protection
					if (sheetName.equals(SHEET_FINISHES) || sheetName.equals(SHEET_CLUBS)) {
						sheet.protectSheet("");
					} else if (sheetName.indexOf(SHEET_RESULTS) > -1) {
						// No protection, at least for HRM (ARM seems to still apply empty password, Nationals applies none for Divisional / Singles but empty password for Doubles!)
						if (sheet.getProtect()) {
							sheet.protectSheet(null); // Throws IndexOutOfBoundsException if locking not enabled
						}
					} else {
						if (sheetPassword != null) {
							sheet.protectSheet(sheetPassword);
						}
					}
					// Remove formulas from the race sheets
					Row headerRow = sheet.getRow(0);
					if (headerRow != null) {
						int rowStart = 1;
						int rowEnd = sheet.getLastRowNum() + 1;
						for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
							Row r = sheet.getRow(rowNum);
							short colStart = r.getFirstCellNum();
							short colEnd = r.getLastCellNum();
							for (short cn = colStart; cn < colEnd; cn++) {
								Cell c = r.getCell(cn, Row.RETURN_NULL_AND_BLANK);
								if (c != null) {
									Cell headerCell = headerRow.getCell(cn, Row.RETURN_BLANK_AS_NULL);
									String colName = "";
									if (headerCell != null) {
										try {
											colName = headerCell.getStringCellValue();
										} catch(IllegalStateException e) {
											// We encountered a non-string value, move on
										}
										if ("Paid".equals(colName)) {
											c.setCellType(Cell.CELL_TYPE_BLANK);
										}
										if ("Notes".equals(colName) && !c.getStringCellValue().equals("ill")) {
											c.setCellType(Cell.CELL_TYPE_BLANK);
										}
									}
									if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
										c.setCellFormula(null);
										// Numeric values set as strings without this
										setNumericValue(c);
									}
									c.removeCellComment();
									if (isRaceSheet) {
										if (cn > 0) {
											if ("Expiry".equals(colName) && c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
												c.setCellStyle(dateStyle);
											} else if ((("Time+/-".equals(colName) || "Start".equals(colName) || "Finish".equals(colName) || "Elapsed".equals(colName))) && c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
												c.setCellStyle(timeStyle);
											} else {
												c.setCellStyle(bodyStyle);
											}
										} else {
											c.setCellStyle(firstColumnStyle);
										}
									}
								}
							}
						}
						// Remove unwanted columns, now that we have finished iterating
						ArrayList<Cell> headerCellsToRemove = new ArrayList<Cell>();
						for (short cn = 0; cn < headerRow.getLastCellNum(); cn++) {
							Cell cell = headerRow.getCell(cn);
							if (cell != null) {
								try {
									String colName = cell.getStringCellValue();
									if (columnsToRemove.indexOf(colName) >= 0) {
										headerCellsToRemove.add(cell);
									}
								} catch(IllegalStateException e) {
									// We encountered a non-string value, move on
								}
							}
						}
						for (Cell cell : headerCellsToRemove) {
							removeSheetColumn(sheet, cell.getColumnIndex());
						}
					}
					// Remove data validation from the sheet
					// First we have to set up a new validation, or sheet.getDataValidations() returns an empty list
					XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
					XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});
					CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
					XSSFDataValidation validation = (XSSFDataValidation)dvHelper.createValidation(dvConstraint, addressList);
					validation.setSuppressDropDownArrow(false);
					validation.setShowErrorBox(true);
					sheet.addValidationData(validation);
					if (sheet.getDataValidations().size() > 0) {
						sheet.getCTWorksheet().unsetDataValidations();
						// This sets the xsi:isNull property on the <dataValidations> element which causes the file to be unreadable by Excel
						//sheet.getCTWorksheet().setDataValidations(null);
					}
					// Remove freeze pane
					sheet.createFreezePane(0,0);
				}

				Map<String, Double> finishTimes = getStartTimes(wb);
				if (finishTimes != null && finishTimes.size() > 0 && raceSheetNames.size() > 0) {
					setHRMFinishTimes(wb, raceSheetNames, finishTimes, timeStyle);
				}

				// Remove any sheets which should not be present for the official HRM
				String[] disallowedSheets = { SHEET_STARTS, "Sheet1" };
				for (int i = 0; i < disallowedSheets.length; i++) {
					int sheetIndex = wb.getSheetIndex(disallowedSheets[i]);
					if (sheetIndex > -1) {
						wb.removeSheetAt(sheetIndex);
					}
				}

				// Finally, protect the workbook
				wb.setWorkbookPassword(sheetPassword, null);
				wb.lockStructure();

				final File exportFile = File.createTempFile("hrmexport-", ".xlsx");
				OutputStream eos = new FileOutputStream(exportFile);
				wb.write(eos);
				eos.close();

				// Serve up the file

				InputStream eis = new FileInputStream(exportFile);
				// Use setHeader as setContentType adds on a text encoding, which confuses Chome
				response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
				//response.setCharacterEncoding("binary");
				response.setHeader("Content-Disposition","attachment;filename=" + file.getTitle() + ".xlsx");
				//response.setContentLength((int) entity.getContentLength());
				response.setHeader("Content-Length", "" + exportFile.length());
				OutputStream outputStream = response.getOutputStream();
				try {
					byte[] buffer = new byte[1024]; // Adjust if you want
					int bytesRead;
					while ((bytesRead = eis.read(buffer)) != -1)
					{
						outputStream.write(buffer, 0, bytesRead);
					}
				} finally {
					eis.close();
				}
			} else {
				throw new Exception("Response entity is null!");
			}
		} else {
			throw new Exception("No Excel export found!");
		}
	}

	private static void setNumericValue(Cell c) {
		if (c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			return;
		}
		if (numberPattern.matcher(c.getStringCellValue()).matches()) {
			c.setCellValue(Double.parseDouble(c.getStringCellValue()));
			c.setCellType(Cell.CELL_TYPE_NUMERIC);
		} else {
			Matcher timeMatcher = timePattern.matcher(c.getStringCellValue());
			if (timeMatcher.matches()) {
				Calendar calendar = Calendar.getInstance();
				calendar.set(Calendar.HOUR_OF_DAY, Integer.parseInt(timeMatcher.group(1)));
				calendar.set(Calendar.MINUTE, Integer.parseInt(timeMatcher.group(2)));
				calendar.set(Calendar.SECOND, Integer.parseInt(timeMatcher.group(3)));
				calendar.set(Calendar.MILLISECOND, 0);
				c.setCellValue(calendar);
				c.setCellType(Cell.CELL_TYPE_NUMERIC);
			}
		}
	}

	private static Map<String, Double> getStartTimes(XSSFWorkbook wb) {
		Map<String, Double> times = new HashMap<String, Double>();
		XSSFSheet startsSheet = wb.getSheet(SHEET_STARTS);
		if (startsSheet == null) {
			return null;
		}
		for (Row row : startsSheet) {
			Cell raceNameCell = row.getCell(0);
			Cell startTimeCell = row.getCell(1);
			setNumericValue(startTimeCell);
			if (raceNameCell != null && startTimeCell != null &&
					raceNameCell.getCellType() == Cell.CELL_TYPE_STRING &&
					startTimeCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				times.put(raceNameCell.getStringCellValue(), startTimeCell.getNumericCellValue());
			}
		}
		return times;
	}

	private static void setHRMFinishTimes(XSSFWorkbook wb, List<String> sheetNames, Map<String, Double> times, CellStyle cellStyle) {
		XSSFSheet summarySheet = wb.getSheet(SHEET_SUMMARY);
		if (summarySheet != null) {
			int i = 1;
			for (String sheetName : sheetNames) {
				Row row = summarySheet.getRow(i);
				if (row == null) {
					row = summarySheet.createRow(i);
				}
				// Add times to column AA from row 2 down, in the order that the sheets appear
				Cell c = row.createCell(26, Cell.CELL_TYPE_BLANK);
				if (times.get(sheetName) != null) {
					c.setCellValue(times.get(sheetName));
				}
				if (cellStyle != null) {
					c.setCellStyle(cellStyle);
				}
				i ++;
			}
		}
	}

	private static void removeSheetColumn(XSSFSheet sheet, int columnToDelete) {
		int maxColumn = 0;
		for ( int r=0; r < sheet.getLastRowNum()+1; r++ ){
			Row row = sheet.getRow( r );

			// if no row exists here; then nothing to do; next!
			if ( row == null )
				continue;

			// if the row doesn't have this many columns then we are good; next!
			int lastColumn = row.getLastCellNum();
			if ( lastColumn > maxColumn )
				maxColumn = lastColumn;

			if ( lastColumn < columnToDelete )
				continue;

			for ( int x=columnToDelete+1; x < lastColumn + 1; x++ ){
				Cell oldCell	= row.getCell(x-1);
				if ( oldCell != null )
					row.removeCell( oldCell );

				Cell nextCell   = row.getCell( x );
				if ( nextCell != null ){
					Cell newCell	= row.createCell( x-1, nextCell.getCellType() );
					cloneCell(newCell, nextCell);
				}
			}
		}


		// Adjust the column widths
		for ( int c=0; c < maxColumn; c++ ){
			sheet.setColumnWidth( c, sheet.getColumnWidth(c+1) );
		}
	}

	/*
	 * Takes an existing Cell and merges all the styles and forumla
	 * into the new one
	 */
	private static void cloneCell( Cell cNew, Cell cOld ){

		switch ( cOld.getCellType() ){
			case Cell.CELL_TYPE_BOOLEAN:{
				cNew.setCellValue( cOld.getBooleanCellValue() );
				break;
			}
			case Cell.CELL_TYPE_NUMERIC:{
				cNew.setCellValue( cOld.getNumericCellValue() );
				break;
			}
			case Cell.CELL_TYPE_STRING:{
				cNew.setCellValue( cOld.getStringCellValue() );
				break;
			}
			case Cell.CELL_TYPE_ERROR:{
				cNew.setCellValue( cOld.getErrorCellValue() );
				break;
			}
			case Cell.CELL_TYPE_FORMULA:{
				cNew.setCellFormula( cOld.getCellFormula() );
				break;
			}
		}

		cNew.setCellStyle( cOld.getCellStyle() );

	}
}
