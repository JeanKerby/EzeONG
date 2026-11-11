<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="java.nio.file.*,java.util.*,java.text.*,org.apache.poi.ss.usermodel.*" %>
<%@ include file="/WEB-INF/jspf/common.jspf" %>
<%
request.setCharacterEncoding("UTF-8");
Path root = storage(application);
List<Path> imports = archive(root.resolve("imports"));
List<Path> exports = archive(root.resolve("exports"));
String selectedId = checkedId(request.getParameter("file"));
Path source = selectedId == null ? null : findFile(root.resolve("imports"), selectedId);
int sheetIndex = 0, start = 0;
try { sheetIndex = Math.max(0, Integer.parseInt(request.getParameter("sheet"))); } catch (Exception ignored) { }
try { start = Math.max(0, Integer.parseInt(request.getParameter("start"))); } catch (Exception ignored) { }
String error = request.getParameter("error"), message = request.getParameter("message");
Workbook book = null; Sheet sheet = null; int cols = 0, last = -1;
if (source != null) try {
    book = WorkbookFactory.create(source.toFile());
    if (sheetIndex >= book.getNumberOfSheets()) sheetIndex = 0;
    sheet = book.getSheetAt(sheetIndex); last = sheet.getLastRowNum();
    for (Row row : sheet) cols = Math.max(cols, Math.max(0, row.getLastCellNum()));
    if (start > last) start = Math.max(0, last - 99);
} catch (Exception ex) { error = "Ce fichier ne peut pas être lu comme un classeur Excel pris en charge."; }
DataFormatter fmt = new DataFormatter(Locale.FRANCE);
FormulaEvaluator eval = book == null ? null : book.getCreationHelper().createFormulaEvaluator();
SimpleDateFormat date = new SimpleDateFormat("dd/MM/yyyy HH:mm", Locale.FRANCE);
%>
<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Atelier Excel | ONG</title><link rel="stylesheet" href="assets/app.css"></head><body>
<div class="layout"><aside><a class="brand" href="index.jsp"><b>A</b><span>Atelier Excel<small>Outils de données ONG</small></span></a><nav><a class="active" href="#importer">＋　Importer un fichier</a><a href="#archives">▤　Archives</a></nav><div class="storage">● Stockage local actif</div></aside><main>
<header><div><small>ESPACE DE TRAVAIL</small><h1>Gestion des classeurs</h1></div><span class="count">●　<%= imports.size() %> fichier(s) importé(s)</span></header>
<% if (message != null) { %><div class="notice"><%= esc(message) %></div><% } if (error != null) { %><div class="notice bad"><%= esc(error) %></div><% } %>
<section id="importer"><div class="section-title"><div><small>01 / CHARGER</small><h2>Importer un classeur</h2></div><span class="muted">XLS · XLSX · XLSM · XLTX · XLTM</span></div>
<form class="upload" action="import.jsp" method="post" enctype="multipart/form-data"><label class="drop" for="excel-file"><b class="upload-icon">↑</b><span><strong>Choisir un fichier Excel</strong><small>Formats pris en charge · 50 Mo maximum</small></span><i>Parcourir</i><input id="excel-file" type="file" name="excelFile" accept=".xls,.xlsx,.xlsm,.xltx,.xltm" required></label><div class="submit-row"><span id="file-name">Aucun fichier sélectionné</span><button>Importer le classeur →</button></div></form></section>
<% if (source != null && book != null && sheet != null) { %>
<section id="apercu"><div class="section-title"><div><small>02 / EXPLORER</small><h2><%= esc(displayName(source)) %></h2></div><a class="link" href="download.jsp?kind=import&id=<%= selectedId %>">Télécharger l’original ↗</a></div>
<div class="tabs"><% for (int i=0;i<book.getNumberOfSheets();i++) { %><a class="<%= i==sheetIndex?"chosen":"" %>" href="index.jsp?file=<%= selectedId %>&sheet=<%= i %>"><%= esc(book.getSheetName(i)) %></a><% } %></div>
<form action="export.jsp" method="post"><input type="hidden" name="file" value="<%= selectedId %>"><input type="hidden" name="sheet" value="<%= sheetIndex %>"><div class="picker"><div class="picker-head"><div><h3>Colonnes à conserver</h3><p>Sélectionnez les champs du fichier personnalisé.</p></div><button type="button" class="light" id="toggle">Tout désélectionner</button></div><div class="columns">
<% Row header=sheet.getRow(0); for(int c=0;c<cols;c++){String label=text(header,c,fmt,eval);if(label.isBlank())label="Colonne "+(c+1); %><label><input type="checkbox" name="columns" value="<%= c %>" checked><span><%= String.format("%02d",c+1) %></span><%= esc(label) %></label><% } %>
</div><div class="export-row"><span>Exporter la feuille complète avec les colonnes choisies</span><div><button class="excel" name="format" value="xlsx">Créer le fichier Excel ↓</button><button name="format" value="pdf">Créer le PDF ↓</button></div></div></div></form>
<div class="table-head"><div><h3>Aperçu des données</h3><p>Feuille « <%= esc(book.getSheetName(sheetIndex)) %> » · lignes <%= last<0?0:start+1 %> à <%= Math.min(last+1,start+100) %></p></div><div><% if(start>0){ %><a class="light" href="index.jsp?file=<%= selectedId %>&sheet=<%= sheetIndex %>&start=<%= Math.max(0,start-100) %>">← Précédent</a><% } if(start+100<=last){ %><a class="light" href="index.jsp?file=<%= selectedId %>&sheet=<%= sheetIndex %>&start=<%= start+100 %>">Suivant →</a><% } %></div></div>
<% if(cols==0 || last<0){ %><div class="empty">Cette feuille ne contient pas encore de données.</div><% }else{ %><div class="table-wrap"><table><thead><tr><th>Ligne</th><% for(int c=0;c<cols;c++){String label=text(sheet.getRow(0),c,fmt,eval);if(label.isBlank())label="Colonne "+(c+1); %><th><%= esc(label) %></th><% } %></tr></thead><tbody><% for(int r=start;r<=Math.min(last,start+99);r++){Row row=sheet.getRow(r); %><tr><td><%= r+1 %></td><% for(int c=0;c<cols;c++){String value=text(row,c,fmt,eval); %><td title="<%= esc(value) %>"><%= esc(value) %></td><% } %></tr><% } %></tbody></table></div><% } %>
</section><% } %>
<section id="archives"><div class="section-title"><div><small>03 / RETROUVER</small><h2>Archives des fichiers</h2></div></div><div class="archives"><div class="archive"><div class="archive-title"><h3>Imports</h3><span><%= imports.size() %></span></div><% if(imports.isEmpty()){ %><p class="muted">Aucun fichier importé pour le moment.</p><% } for(Path file:imports){ %><div class="file"><b>X</b><div><a href="index.jsp?file=<%= fileId(file) %>"><%= esc(displayName(file)) %></a><small><%= date.format(new Date(file.toFile().lastModified())) %> · <%= String.format(Locale.FRANCE,"%.1f Ko",file.toFile().length()/1024.0) %></small></div><a class="download" href="download.jsp?kind=import&id=<%= fileId(file) %>">↓</a></div><% } %></div>
<div class="archive"><div class="archive-title"><h3>Exports</h3><span><%= exports.size() %></span></div><% if(exports.isEmpty()){ %><p class="muted">Aucun fichier exporté pour le moment.</p><% } for(Path file:exports){ %><div class="file"><b class="pdf"><%= displayName(file).toLowerCase(Locale.ROOT).endsWith(".pdf")?"P":"X" %></b><div><a href="download.jsp?kind=export&id=<%= fileId(file) %>"><%= esc(displayName(file)) %></a><small><%= date.format(new Date(file.toFile().lastModified())) %> · <%= String.format(Locale.FRANCE,"%.1f Ko",file.toFile().length()/1024.0) %></small></div><a class="download" href="download.jsp?kind=export&id=<%= fileId(file) %>">↓</a></div><% } %></div></div></section>
<footer>Atelier Excel ONG <span>Les fichiers sont conservés sur le serveur.</span></footer></main></div>
<script>const f=document.getElementById('excel-file');if(f)f.addEventListener('change',()=>document.getElementById('file-name').textContent=f.files.length?f.files[0].name:'Aucun fichier sélectionné');const t=document.getElementById('toggle');if(t)t.addEventListener('click',()=>{const boxes=[...document.querySelectorAll('.columns input')],check=boxes.some(b=>!b.checked);boxes.forEach(b=>b.checked=check);t.textContent=check?'Tout désélectionner':'Tout sélectionner'});</script>
<% if(book!=null)book.close(); %></body></html>