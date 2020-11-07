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
   <title>${requestScope.title}</title>
</head>
<body>
    <center>
        <h2>${requestScope.message}</h2>

        <p align="center">Таблица №1. Фактические медосмотры (по датам).</p>

        <%
               String name = (String)request.getAttribute("docxName");
               String name2 = (String)request.getAttribute("docx2Name");
               String dir = (String)request.getAttribute("reportsDir");
               String link = dir+name;
               String link2 = dir+name2;

               String fName0 = "";
               try {
                  String URLEncodedFileName = URLEncoder.encode(name, "UTF-8");
                  String ResultFileName = URLEncodedFileName.replace('+', ' ');
                  fName0 = ResultFileName;
                 } catch (UnsupportedEncodingException e) {
                        e.printStackTrace();
                   }         
            out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName0.substring(3, fName0.length()) +"\" download=\"\">СКАЧАТЬ</a>");
       
        %>

        <br><br>
        <p align="center">Таблица №2. Детализация медосмотров (по фамилиям).</p>

        <%
            String fName = "";
            try {
                  String URLEncodedFileName = URLEncoder.encode(name2, "UTF-8");
                  String ResultFileName = URLEncodedFileName.replace('+', ' ');
                  fName = ResultFileName;
                 } catch (UnsupportedEncodingException e) {
                        e.printStackTrace();
                   }         
            out.println("<a class=\"w3-button w3-ripple w3-teal\" href=\"."+ File.separator + dir + File.separator + fName.substring(3, fName.length()) +"\" download=\"\">СКАЧАТЬ</a>");
        %>

        <br><br><br><br>
        <p align="center">==============================================</p>
        <a class="w3-button w3-ripple w3-teal" href="./" >Ещё разок</a>
    </center>
    <br>

    <br>
    <div class="w3-container w3-left-align">
        <jsp:include page="/resources/support.html" />
    </div>
</body>
</html>
