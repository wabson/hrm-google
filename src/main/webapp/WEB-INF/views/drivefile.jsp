<?xml version="1.0" encoding="UTF-8" ?>
<jsp:root xmlns:jsp="http://java.sun.com/JSP/Page" version="2.0"
	xmlns:c="http://java.sun.com/jsp/jstl/core"
	xmlns:form="http://www.springframework.org/tags/form">
	<jsp:directive.page language="java"
		contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" />
	<jsp:text>
		<![CDATA[ <!DOCTYPE html> ]]>
	</jsp:text>
	<html xmlns="http://www.w3.org/1999/xhtml" lang="en">
<head>
<title>HRM Google - ${title}</title>
<jsp:directive.include file="header.jspf" />
</head>

<body>
	<jsp:directive.include file="bar.jspf" />

	<div class="container">
		
		<h1>${title}</h1>
		
		<c:if test="${not empty file}">
			<div class="span12 columns">
				<form>
					<div class="clearfix">
						<label><strong>Created:</strong></label>
						<label class="text">${file.createdDate}</label>
					</div>
					<div class="clearfix">
						<label><strong>Last modified:</strong></label>
						<label class="text">${file.modifiedDate} by ${file.lastModifyingUserName}</label>
					</div>
				</form>
			</div>
			<div class="span4 columns">
				<a href="https://docs.google.com/spreadsheet/ccc?key=${file.id}" target="_blank">Edit Spreadsheet</a>
			</div>
			<div class="span4 columns">
				<a href="${pageContext.request.contextPath}/downloadfile/${file.title}.xlsx?fileId=${file.id}" target="_blank" class="button" title="Download HRM">Download HRM</a>
			</div>
			<!--
			<div class="span4 columns">
				<img src="${file.thumbnailLink}" />
			</div>
			-->
			<div class="clear"></div>
		</c:if>
		<c:if test="${empty file and not empty param.fileId}">
			<div>The specified file could not be found</div>
		</c:if>
	</div>
</body>
</html>
</jsp:root>