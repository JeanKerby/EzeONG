<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" session="false" %>
<%@ page import="java.io.*,java.nio.file.*,java.net.URLEncoder,java.nio.charset.StandardCharsets,java.util.Locale" %>
<%@ include file="/WEB-INF/jspf/common.jspf" %>
<%
String kind=request.getParameter("kind"), id=checkedId(request.getParameter("id"));
if(id==null||!("import".equals(kind)||"export".equals(kind))){response.sendError(404);return;}
Path file=findFile(storage(application).resolve("import".equals(kind)?"imports":"exports"),id);
if(file==null){response.sendError(404);return;}
String name=displayName(file), lower=name.toLowerCase(Locale.ROOT);
String type=lower.endsWith(".pdf")?"application/pdf":lower.endsWith(".xls")?"application/vnd.ms-excel":lower.endsWith(".xlsm")?"application/vnd.ms-excel.sheet.macroEnabled.12":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
response.reset(); response.setContentType(type); response.setHeader("X-Content-Type-Options","nosniff");
response.setHeader("Content-Disposition","attachment; filename=\"fichier\"; filename*=UTF-8''"+URLEncoder.encode(name,StandardCharsets.UTF_8).replace("+","%20"));
response.setContentLengthLong(Files.size(file));
try(InputStream in=Files.newInputStream(file);OutputStream outputStream=response.getOutputStream()){in.transferTo(outputStream);}
%>