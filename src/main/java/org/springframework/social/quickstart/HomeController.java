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
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
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
import org.springframework.social.google.api.userinfo.GoogleUserProfile;
import org.springframework.social.quickstart.drive.DateOperators;
import org.springframework.social.quickstart.drive.DriveSearchForm;
import org.springframework.social.quickstart.drive.OptionalBoolean;
import org.springframework.social.quickstart.drive.WorksheetForm;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class HomeController {

	private final Google google;
	
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
	public ModelAndView getDriveFiles(DriveSearchForm command) {

        GoogleUserProfile profile = google.userOperations().getUserProfile();

		DriveFileQueryBuilder queryBuilder = google.driveOperations().driveFileQuery()
				.fromPage(command.getPageToken());
		
		queryBuilder.trashed(false);
		queryBuilder.mimeTypeIs("application/vnd.google-apps.spreadsheet");
		//queryBuilder.parentIs("appdata");
		
		if(hasText(command.getTitleContains())) {
			queryBuilder.titleContains(command.getTitleContains());
		}
		
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
            .addObject("profile", profile);
	}
	
	@RequestMapping(value="workbook", method=GET)
	public ModelAndView task() {
		
		return new ModelAndView("worksheet", "command", new WorksheetForm());
	}
	
	@RequestMapping(value="workbook", method=GET, params="id")
	public ModelAndView task(String id) {
		
		DriveFile file = google.driveOperations().getFile(id);
		WorksheetForm command = new WorksheetForm(file.getId(), file.getTitle());
		return new ModelAndView("task", "command", command);
	}
	
	@RequestMapping(value="workbook", method=POST)
	public ModelAndView saveWorksheet(WorksheetForm command, BindingResult result) {
		
		if(result.hasErrors()) {
			return new ModelAndView("worksheet", "command", command);
		}

		String[] parents = { "root" };
		String srcId = null;
		if (command.getType().equals("hrm")) {
			srcId = "0AsA0SXNo_BkZdGxiTlhsQ08zcUxpTW5VaUt2N0F0MlE";
		} else if (command.getType().equals("arm")) {
			srcId = "0AsA0SXNo_BkZdC1hdUhfeDdBVWppdVBESHZyaThiMkE";
		} else if (command.getType().equals("nrm")) {
			srcId = "0AsA0SXNo_BkZdEd3N1ZsdzlkbnFyRFkzNm45YTJkbFE";
		}
		google.driveOperations().copy(srcId, parents, command.getTitle());
		// TODO Record HRM version in a custom property
		
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
		DriveFile file = google.driveOperations().copy(fileId, new String[]{parentId}, newName);
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

				// Write properties using POI

				OPCPackage pkg = OPCPackage.open(temp);
				XSSFWorkbook wb = new XSSFWorkbook(pkg);
				POIXMLProperties props = wb.getProperties();
				POIXMLProperties.CustomProperties cust =  props.getCustomProperties();

				CTProperty hrmProperty = cust.getUnderlyingProperties().addNewProperty();
				hrmProperty.setBool(true);
				hrmProperty.setName("HRM");
				hrmProperty.setFmtid("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
				hrmProperty.setPid(2);
				
				CTProperty versionProperty = cust.getUnderlyingProperties().addNewProperty();
				versionProperty.setName("Version");
				versionProperty.setR8(9.0);
				versionProperty.setFmtid("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
				versionProperty.setPid(3);

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