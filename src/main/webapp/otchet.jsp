<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%@ page isELIgnored = "false" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>
<%@ page import="java.net.URLEncoder" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
    "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
   <link rel="stylesheet" href="styles/w3.css">
   <link rel="stylesheet" href="styles/list.css">
   <link rel="stylesheet" href="styles/divtable.css">
   <title>${requestScope.title}</title>
</head>
<body>
    <center>
        <h2>${requestScope.message}</h2>

        <div class="divTable" style="width: 600px;" >
        <div class="divTableBody">
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №1. Фактические медосмотры (по датам).</div>
        <div class="divTableCell">
            <%
                String name = (String)request.getAttribute("docxName");
                String dir = (String)request.getAttribute("reportsDir");
                String fName1 = "";
                try {
                      String URLEncodedFileName = URLEncoder.encode(name, "UTF-8");
                      String ResultFileName = URLEncodedFileName.replace('+', ' ');
                      fName1 = ResultFileName;
                    } catch (UnsupportedEncodingException e) {
                          e.printStackTrace();
                    }
                out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName1.substring(3, fName1.length()) +"\" download=\"\">СКАЧАТЬ</a>");
            %>
        </div>
        </div>
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №2. Детализация медосмотров (по ФИО водителей).</div>
        <div class="divTableCell">
            <%
                        String name2 = (String)request.getAttribute("docx2Name");
                        String fName2 = "";
                        try {
                                    String URLEncodedFileName = URLEncoder.encode(name2, "UTF-8");
                                    String ResultFileName = URLEncodedFileName.replace('+', ' ');
                                    fName2 = ResultFileName;
                        } catch (UnsupportedEncodingException e) {
                                    e.printStackTrace();
                        }

                        out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName2.substring(3, fName2.length()) +"\" download=\"\">СКАЧАТЬ</a>");
            %>
        </div>
        </div>
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №3. Детализация медосмотров (по ФИО медработников).</div>
        <div class="divTableCell">
            <%
                             String name3 = (String)request.getAttribute("docx3Name");
                             String fName3 = "";
                                try {
                                            String URLEncodedFileName = URLEncoder.encode(name3, "UTF-8");
                                            String ResultFileName = URLEncodedFileName.replace('+', ' ');
                                            fName3 = ResultFileName;
                                } catch (UnsupportedEncodingException e) {
                                            e.printStackTrace();
                                }

                             out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName3.substring(3, fName3.length()) +"\" download=\"\">СКАЧАТЬ</a>");
            %>
        </div>
        </div>
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №4. Детализация медосмотров (по точкам выпуска).</div>
        <div class="divTableCell">
            <%
                                  String name4 = (String)request.getAttribute("docx4Name");
                                  String fName4 = "";
                                  try {
                                        String URLEncodedFileName = URLEncoder.encode(name4, "UTF-8");
                                        String ResultFileName = URLEncodedFileName.replace('+', ' ');
                                        fName4 = ResultFileName;
                                       } catch (UnsupportedEncodingException e) {
                                                e.printStackTrace();
                                       }

                    			  out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName4.substring(3, fName4.length()) +"\" download=\"\">СКАЧАТЬ</a>");
            %>
        </div>
        </div>
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №5. Реестр медосмотров (предр., послер. и линейный).</div>
        <div class="divTableCell">
            <%
                                          String name5 = (String)request.getAttribute("docx5Name");
                                          String fName5 = "";
                                          try {
                                                String URLEncodedFileName = URLEncoder.encode(name5, "UTF-8");
                                                String ResultFileName = URLEncodedFileName.replace('+', ' ');
                                                fName5 = ResultFileName;
                                               } catch (UnsupportedEncodingException e) {
                                                        e.printStackTrace();
                                               }

                            			  out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName5.substring(3, fName5.length()) +"\" download=\"\">СКАЧАТЬ</a>");

            %>
        </div>
        </div>
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №6.  Причины недопусков (статистика).</div>
        <div class="divTableCell">
            <%
                                                  String name6 = (String)request.getAttribute("docx6Name");
                                                  String fName6 = "";
                                                  try {
                                                        String URLEncodedFileName = URLEncoder.encode(name6, "UTF-8");
                                                        String ResultFileName = URLEncodedFileName.replace('+', ' ');
                                                        fName6 = ResultFileName;
                                                       } catch (UnsupportedEncodingException e) {
                                                                e.printStackTrace();
                                                       }

                                    			  out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName6.substring(3, fName6.length()) +"\" download=\"\">СКАЧАТЬ</a>");

            %>
        </div>
        </div>
        <div class="divTableRow">
        <div class="divTableCell" align="left">Отчёт №7.  Группы риска (давление и пульс).</div>
        <div class="divTableCell">
            <%
                                    String name7 = (String)request.getAttribute("docx7Name");
                                    String fName7 = "";
                                    try {
                                                String URLEncodedFileName = URLEncoder.encode(name7, "UTF-8");
                                                String ResultFileName = URLEncodedFileName.replace('+', ' ');
                                                fName7 = ResultFileName;
                                    } catch (UnsupportedEncodingException e) {
                                                e.printStackTrace();
                                    }

                                    out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName7.substring(3, fName7.length()) +"\" download=\"\">СКАЧАТЬ</a>");
            %>
        </div>
        </div>
        </div>
        </div>

        <br><br>
        <p align="center">==============================================</p>
        <a class="w3-button w3-ripple w3-teal" href="./" >Повторить</a>
    </center>
    <br>

    <br>
    <div class="w3-container w3-left-align">
        <jsp:include page="/resources/support.html" />
    </div>
</body>
</html>