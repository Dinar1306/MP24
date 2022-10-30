<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%@ page isELIgnored = "false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
    "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset="UTF-8">
   <link rel="stylesheet" href="styles/w3.css">
   <link rel="stylesheet" href="styles/list.css">
   <title>${requestScope.title}</title>
</head>
<body>
    <center>
        <h2>${requestScope.message}</h2>

        <a class="w3-button w3-ripple w3-teal" href="./" >Попробовать снова</a>

        <br><br><br><br>
                <%
                   String debug = (String)request.getAttribute("debug");
                   if (!(debug==null)){
                        out.println("<h0><b>Служебная информация: </b></h0>" + System.lineSeparator() + debug);
                   }

                %>

    </center>
    <br><br>

    <div class="w3-container w3-left-align">
         <!--<jsp:include page="/resources/support.html" />-->
         <a  href="${requestScope.dev}" >Support is here</a> ;)
    </div>
</body>
</html>