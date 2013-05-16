package org.springframework.social.quickstart.drive;

import static org.springframework.util.StringUtils.hasText;
import static org.springframework.format.annotation.DateTimeFormat.ISO.DATE;

import java.util.Date;

import org.hibernate.validator.constraints.NotBlank;
import org.springframework.format.annotation.DateTimeFormat;

public class WorksheetForm {

	private String list;
	private String id;
	
	@NotBlank(message="Task list name may not be empty")
	private String title;
	
	private String notes;
	
	private String type;
	
	@DateTimeFormat(iso=DATE)
	private Date due;
	
	@DateTimeFormat(iso=DATE)
	private Date completed;
	
	public WorksheetForm() {
		
	}
	
	public WorksheetForm(String id, String title) {
		this.id = id;
		this.title = title;
	}
	
	public String getList() {
		return hasText(list) ? list : "@default";
	}

	public void setList(String list) {
		this.list = list;
	}

	public String getId() {
		return hasText(id) ? id : null;
	}
	
	public void setId(String id) {
		this.id = id;
	}
	
	public String getTitle() {
		return title;
	}
	
	public void setTitle(String title) {
		this.title = title;
	}
	
	public String getNotes() {
		return hasText(notes) ? notes : null;
	}

	public void setNotes(String notes) {
		this.notes = notes;
	}

	public Date getDue() {
		return due;
	}

	public void setDue(Date due) {
		this.due = due;
	}

	public Date getCompleted() {
		return completed;
	}

	public void setCompleted(Date completed) {
		this.completed = completed;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}
	
}
