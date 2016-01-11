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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddressList;
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

				// Go through sheets
				for (int i = 0; i < numSheets; i++) {
					sheet = wb.getSheetAt(i);
					sheetName = sheet.getSheetName();
					// Set sheet protection
					if (sheetName.equals("Finishes") || sheetName.equals("Clubs")) {
						sheet.protectSheet("");
					} else if (sheetName.indexOf("Results") > -1) {
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
									}
									if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
										c.setCellFormula(null);
									}
									c.removeCellComment();
								}
							}
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

				// Remove any sheets which should not be present for the official HRM
				String[] disallowedSheets = { "Starts", "Sheet1" };
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
}
