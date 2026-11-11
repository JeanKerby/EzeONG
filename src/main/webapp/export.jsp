<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="java.io.*,java.nio.file.*,java.net.URLEncoder,java.nio.charset.StandardCharsets,java.util.*,java.text.Normalizer,java.text.SimpleDateFormat,java.awt.Color,org.apache.poi.ss.usermodel.*,org.apache.poi.xssf.usermodel.XSSFWorkbook,org.apache.pdfbox.pdmodel.*,org.apache.pdfbox.pdmodel.common.PDRectangle,org.apache.pdfbox.pdmodel.font.*" %>
<%@ include file="/WEB-INF/jspf/common.jspf" %>
<%!
private String pdfText(String input) {
    StringBuilder result=new StringBuilder();
    for(int i=0;i<input.length();i++){char c=input.charAt(i);result.append(c>=32&&c<=255&&c!=173?c:'?');}
    return result.toString();
}
private void pdfCell(PDPageContentStream stream, PDFont font, String value, float x, float y, float width, float size, Color color) throws IOException {
    String text=pdfText(value);
    while(!text.isEmpty()&&font.getStringWidth(text)*size/1000f>width)text=text.substring(0,text.length()-1);
    if(text.length()<pdfText(value).length()&&text.length()>2)text=text.substring(0,text.length()-2)+"..";
    stream.beginText();stream.setFont(font,size);stream.setNonStrokingColor(color);stream.newLineAtOffset(x,y);stream.showText(text);stream.endText();
}
%>
<%
request.setCharacterEncoding("UTF-8"); Path outputFile=null; String redirect;
try {
    String id=checkedId(request.getParameter("file")), format=request.getParameter("format");
    int sheetIndex=Integer.parseInt(request.getParameter("sheet")); String[] raw=request.getParameterValues("columns");
    if(id==null||!("xlsx".equals(format)||"pdf".equals(format))||raw==null||raw.length==0)throw new IllegalArgumentException("Choisissez au moins une colonne et un format d’export.");
    Path root=storage(application), source=findFile(root.resolve("imports"),id);
    if(source==null)throw new IllegalArgumentException("Le fichier source est introuvable.");
    String base=displayName(source);int dot=base.lastIndexOf('.');if(dot>0)base=base.substring(0,dot);base=cleanName(base);
    String stamp=new SimpleDateFormat("yyyyMMdd-HHmmss").format(new Date()), ext="xlsx".equals(format)?".xlsx":".pdf";
    String exportName=base+"-personnalise-"+stamp+ext, exportId=UUID.randomUUID().toString();
    Path folder=root.resolve("exports");outputFile=folder.resolve(exportId+"__"+exportName).normalize();
    try(Workbook input=WorkbookFactory.create(source.toFile())){
        if(sheetIndex<0||sheetIndex>=input.getNumberOfSheets())throw new IllegalArgumentException("Feuille invalide.");
        Sheet sheet=input.getSheetAt(sheetIndex);int count=0;for(Row row:sheet)count=Math.max(count,Math.max(0,row.getLastCellNum()));
        SortedSet<Integer> selected=new TreeSet<>();for(String value:raw){int col=Integer.parseInt(value);if(col<0||col>=count)throw new IllegalArgumentException("Une colonne sélectionnée est invalide.");selected.add(col);}
        if(selected.isEmpty())throw new IllegalArgumentException("Choisissez au moins une colonne.");
        List<Integer> columns=new ArrayList<>(selected);DataFormatter formatter=new DataFormatter(Locale.FRANCE);FormulaEvaluator evaluator=input.getCreationHelper().createFormulaEvaluator();
        if("xlsx".equals(format)){
            try(Workbook result=new XSSFWorkbook()){
                Sheet target=result.createSheet("Données personnalisées");Map<String,CellStyle> styles=new HashMap<>();
                for(int r=0;r<=sheet.getLastRowNum();r++){Row from=sheet.getRow(r);if(from==null)continue;Row to=target.createRow(r);
                    for(int c=0;c<columns.size();c++){Cell sourceCell=from.getCell(columns.get(c),Row.MissingCellPolicy.RETURN_BLANK_AS_NULL), dest=to.createCell(c);if(sourceCell==null)continue;
                        switch(sourceCell.getCellType()){
                            case STRING -> dest.setCellValue(sourceCell.getStringCellValue());
                            case NUMERIC -> {dest.setCellValue(sourceCell.getNumericCellValue());String fmt=sourceCell.getCellStyle().getDataFormatString();if(fmt!=null&&!fmt.isBlank()&&!"General".equalsIgnoreCase(fmt)){CellStyle style=styles.computeIfAbsent(fmt,key->{CellStyle made=result.createCellStyle();made.setDataFormat(result.createDataFormat().getFormat(key));return made;});dest.setCellStyle(style);}}
                            case BOOLEAN -> dest.setCellValue(sourceCell.getBooleanCellValue());
                            case FORMULA, ERROR -> dest.setCellValue(formatter.formatCellValue(sourceCell,evaluator));
                            default -> dest.setBlank();
                        }
                    }
                }
                for(int c=0;c<columns.size();c++)target.setColumnWidth(c,18*256);
                try(OutputStream outputStream=Files.newOutputStream(outputFile)){result.write(outputStream);}
            }
        }else{
            try(PDDocument pdf=new PDDocument()){
                PDFont regular=new PDType1Font(Standard14Fonts.FontName.HELVETICA), bold=new PDType1Font(Standard14Fonts.FontName.HELVETICA_BOLD);
                float width=PDRectangle.A4.getHeight(),height=PDRectangle.A4.getWidth(),margin=32,usable=width-2*margin,colWidth=usable/columns.size();
                PDPageContentStream stream=null;int dataRow=0;
                try{
                    for(int r=1;r<=sheet.getLastRowNum();r++,dataRow++){
                        int pageRow=dataRow%24;
                        if(pageRow==0){if(stream!=null)stream.close();PDPage pdfPage=new PDPage(new PDRectangle(width,height));pdf.addPage(pdfPage);stream=new PDPageContentStream(pdf,pdfPage);
                            pdfCell(stream,bold,"Export personnalisé",margin,height-34,usable,14,new Color(32,51,45));
                            pdfCell(stream,regular,displayName(source)+" - "+input.getSheetName(sheetIndex),margin,height-50,usable,8,new Color(105,120,110));
                            float hy=height-78;stream.setNonStrokingColor(new Color(23,104,79));stream.addRect(margin,hy,usable,22);stream.fill();
                            for(int c=0;c<columns.size();c++){String label=text(sheet.getRow(0),columns.get(c),formatter,evaluator);if(label.isBlank())label="Colonne "+(columns.get(c)+1);pdfCell(stream,bold,label,margin+c*colWidth+4,hy+7,colWidth-8,8,Color.WHITE);}
                        }
                        float y=height-78-22-(pageRow+1)*17;
                        if(pageRow%2==1){stream.setNonStrokingColor(new Color(241,246,242));stream.addRect(margin,y,usable,17);stream.fill();}
                        Row row=sheet.getRow(r);for(int c=0;c<columns.size();c++)pdfCell(stream,regular,text(row,columns.get(c),formatter,evaluator),margin+c*colWidth+4,y+5,colWidth-8,7,new Color(40,54,47));
                    }
                }finally{if(stream!=null)stream.close();}
                if(pdf.getNumberOfPages()==0){PDPage pdfPage=new PDPage(new PDRectangle(width,height));pdf.addPage(pdfPage);}
                pdf.save(outputFile.toFile());
            }
        }
    }
    redirect="index.jsp?message="+URLEncoder.encode("Export archivé : "+exportName,StandardCharsets.UTF_8);
}catch(IllegalArgumentException ex){if(outputFile!=null)try{Files.deleteIfExists(outputFile);}catch(IOException ignored){}redirect="index.jsp?error="+URLEncoder.encode(ex.getMessage(),StandardCharsets.UTF_8);}
catch(Exception ex){if(outputFile!=null)try{Files.deleteIfExists(outputFile);}catch(IOException ignored){}redirect="index.jsp?error="+URLEncoder.encode("La création du fichier personnalisé a échoué.",StandardCharsets.UTF_8);}
response.sendRedirect(redirect);
%>