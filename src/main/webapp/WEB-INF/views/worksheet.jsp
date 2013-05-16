<?xml version="1.0" encoding="UTF-8" ?>
<jsp:root xmlns:jsp="http://java.sun.com/JSP/Page" version="2.0"
	xmlns:c="http://java.sun.com/jsp/jstl/core"
	xmlns:form="http://www.springframework.org/tags/form"
	xmlns:spring="http://www.springframework.org/tags">
	<jsp:directive.page language="java"
		contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" />
	<jsp:text>
		<![CDATA[ <!DOCTYPE html> ]]>
	</jsp:text>
	<html xmlns="http://www.w3.org/1999/xhtml" lang="en">
<head>
<title>HRM Google - New Workbook</title>
<jsp:directive.include file="header.jspf" />
</head>

<body>
	<c:set var="selected" value="tasks" />
	<c:set var="subselected" value="tasklists" />
	<jsp:directive.include file="bar.jspf" />

	<div class="container">
	
		<h1>Workbook Details</h1>

		<div class="row">

			<div class="span10 columns">

				<form:form>
					<div class="clearfix">
						<label for="title">Name</label>
						<div class="input">
							<form:input path="title" cssClass="xlarge" />
						</div>
						<label for="title">Type</label>
						<div class="input">
							<form:select path="type" cssClass="xlarge">
								<form:option value="hrm">Hasler Race</form:option>
								<form:option value="arm">Assessment Race</form:option>
								<form:option value="nrm">Nationals Race</form:option>
							</form:select>
						</div>
					</div>
					<div class="actions">
						<input type="submit" class="btn btn-primary" value="Save" />
						<a href="./" class="btn leftMargin">Cancel</a>
						<c:if test="${param.id != null}">
							<input name="delete" type="submit" class="btn btn-danger leftMargin" value="Delete" 
								onclick="return confirm('Are you sure you want to delete this task list?')" />
						</c:if>
					</div>
					<spring:hasBindErrors name="taskListForm">
						<div class="error">
							<c:forEach items="${errors.allErrors}" var="error">
								<div><span class="help-inline"><spring:message message="${error}" /></span></div>
							</c:forEach>
						</div>
					</spring:hasBindErrors>
				</form:form>
			</div>
		</div>
	</div>
</body>
	</html>
</jsp:root>