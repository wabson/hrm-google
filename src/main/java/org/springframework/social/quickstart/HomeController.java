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

import java.util.LinkedHashMap;
import java.util.Map;

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
		google.driveOperations().copy("0AsA0SXNo_BkZdGxiTlhsQ08zcUxpTW5VaUt2N0F0MlE", parents, command.getTitle());
		
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
}