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
<title>HRM Google - My Files</title>
<jsp:directive.include file="header.jspf" />
<link href="resources/css/silk-sprite.css" rel="stylesheet"/>
<script src="resources/js/bootbox.min.js"><![CDATA[<!-- -->]]></script>
<script src="resources/js/drive.js"><![CDATA[<!-- -->]]></script>
</head>

<body>
	<c:if test="${empty selected}">
		<c:set var="selected" value="none" />
	</c:if>
	<jsp:directive.include file="bar.jspf" />

	<div class="container">
	
		<div class="content">

			<h1>My Files</h1>
			
			<c:if test="${not empty selected}">
			<p class="pull-right">
				<a href="${selected}/new" class="btn success leftMargin">New Race File</a>
			</p>
			</c:if>
			
			<form:form method="get" cssClass="form-horizontal">
				<form:hidden path="parentId" />
				<form:input path="titleContains" cssClass="input-large"/><![CDATA[&nbsp;]]>
				<input type="submit" class="btn" value="Search"/>
			</form:form>
			
			<c:if test="${not empty param.parentId and param.parentId ne 'root'}">
				<a href="?parentId=root"><![CDATA[&larr; Root Folder]]></a>
			</c:if>
			
			<c:if test="${not empty files.items}">
				<table class="table table-hover">
					<thead>
						<th></th>
						<th colspan="2">File Name</th>
						<th></th>
						<th></th>
						<th></th>
						<th></th>
					</thead>
					<tbody>
						<c:forEach items="${files.items}" var="file">
							<tr file-id="${file.id}" file-name="${file.title}">
								<td width="16">
									<a href="javascript:void(0)" class="star ui-silk ${file.starred ? '' : 'gray'} ui-silk-star" title="${file.starred ? 'Unstar' : 'Star'}"><!--  --></a>
								</td>
								<td width="16">
									<script>
										var icon = getIcon('${file.mimeType}');
										document.write('<span class="ui-silk ui-silk-' + icon + '"><!--  --></span> ');
									</script>
								</td>
								<td class="name-cell ${file.viewed ? '' : 'unviewed'} ${file.trashed ? 'trashed' : ''}">
									<c:if test="${file.folder}">
										<a href="?parentId=${file.id}">${file.title}</a>
									</c:if>
									<c:if test="${not file.folder}">
										<a href="https://docs.google.com/spreadsheet/ccc?key=${file.id}" target="_blank">${file.title}</a>
									</c:if>
								</td>
								<td><a href="downloadfile/${file.title}.xlsx?fileId=${file.id}" target="_blank" class="export ui-silk ui-silk-page-white-put" title="Download HRM"><![CDATA[<!-- -->]]></a></td>
								<td><a href="javascript:void(0)" class="copy ui-silk ui-silk-page-white-copy" title="Copy"><![CDATA[<!-- -->]]></a></td>
								<td><a href="javascript:void(0)" class="trash ui-silk ui-silk-delete" title="Trash"><![CDATA[<!-- -->]]></a></td>
							</tr>
						</c:forEach>
					</tbody>
				</table>
				<c:if test="${not empty files.nextPageToken}">
					<p class="pull-right"><a href="?text=${param.text}&amp;pageToken=${files.nextPageToken}"><![CDATA[Next Page &rarr;]]></a></p>
				</c:if>
			</c:if>
			<c:if test="${empty files.items}">
				<div>No files were found</div>
			</c:if>
		</div>
	</div>
</body>
</html>
</jsp:root>