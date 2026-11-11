<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="java.io.*,java.nio.file.*,java.net.URLEncoder,java.nio.charset.StandardCharsets,java.util.*,jakarta.servlet.http.Part,org.apache.poi.ss.usermodel.*" %>
<%@ include file="/WEB-INF/jspf/common.jspf" %>
<%
request.setCharacterEncoding("UTF-8"); String target;
try {
    Part part=request.getPart("excelFile");
    if(part==null||part.getSize()==0)throw new IllegalArgumentException("Choisissez un fichier non vide.");
    if(part.getSize()>52428800L)throw new IllegalArgumentException("Le fichier dépasse la limite de 50 Mo.");
    String name=cleanName(part.getSubmittedFileName()), lower=name.toLowerCase(Locale.ROOT);
    if(!(lower.endsWith(".xls")||lower.endsWith(".xlsx")||lower.endsWith(".xlsm")||lower.endsWith(".xltx")||lower.endsWith(".xltm")))throw new IllegalArgumentException("Formats acceptés : .xls, .xlsx, .xlsm, .xltx et .xltm.");
    try(Workbook checked=WorkbookFactory.create(part.getInputStream())){if(checked.getNumberOfSheets()==0)throw new IllegalArgumentException("Le classeur ne contient aucune feuille.");}
    String id=UUID.randomUUID().toString(); Path folder=storage(application).resolve("imports"), file=folder.resolve(id+"__"+name).normalize();
    if(!file.getParent().equals(folder.normalize()))throw new IllegalArgumentException("Nom de fichier invalide.");
    try(InputStream in=part.getInputStream()){Files.copy(in,file);}
    target="index.jsp?file="+id+"&message="+URLEncoder.encode("Fichier importé et archivé.",StandardCharsets.UTF_8);
}catch(IllegalArgumentException ex){target="index.jsp?error="+URLEncoder.encode(ex.getMessage(),StandardCharsets.UTF_8);}
catch(Exception ex){target="index.jsp?error="+URLEncoder.encode("Import impossible : vérifiez que le fichier est un classeur Excel valide.",StandardCharsets.UTF_8);}
response.sendRedirect(target);
%>