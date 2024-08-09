<%@page import="java.util.ArrayList"%>
<%@page import="online.ITmed.ReportsTable"%>
<%@page import="java.util.List"%>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@ taglib prefix="display" uri="http://displaytag.sf.net" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
    "http://www.w3.org/TR/html4/loose.dtd">

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Все отчеты.</title>
        <link rel="stylesheet" href="styles/displaytag.css" type="text/css">
        <link rel="stylesheet" href="styles/screen.css" type="text/css">
        <link rel="stylesheet" href="styles/site.css" type="text/css">
        <link rel="stylesheet" href="styles/w3.css">
        <link rel="stylesheet" href="styles/list.css">

        <script>
                function goToSite() {
                      window.location.replace('list');
                }
        </script>

    </head>
    <body>


        <%
            List<ReportsTable> spisokOtchetov_v2 = (List<ReportsTable>) request.getAttribute("spisokOtchetov_v2");
        %>
        <center>
        <div id='tab0' class="tab_content" style="display: block; width: 100%">
            <h3>Список отчетов</h3>
            <p><b>ИНФО:</b> клик по названию столбца для сортировки.<br>
               <span style="color:red">ВНИМАНИЕ!</span> Удаление происходит без подтверждения, т.е. сразу по нажатию кнопки "Удалить"</p>
            <display:table name="spisokOtchetov_v2" pagesize="50" keepStatus="true" export="false" sort="list" uid="zero">

                <display:column property="orgName"    title="Название организации" sortable="true" headerClass="sortable" />
                <display:column property="tipOtcheta" title="Тип отчета"           sortable="true" headerClass="sortable" />
                <display:column property="period"     title="Период"               sortable="true" headerClass="sortable" />
                <display:column property="dataVremya" title="Время создания"       sortable="true" headerClass="sortable" />
                <display:column property="downloadLink" title="Скачать"/>
                <display:column property="removeLink" title="Удалить" />

            </display:table>
        </div>
        <p><a href="/">Назад</a></p>
        </center>
                      <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      <div class="w3-container w3-left-align">
                               <!--<jsp:include page="/resources/support.html" />
                               <a  href="${requestScope.dev}" >Support is here</a> ;)-->
                               <%
                                  String dev = (String)request.getAttribute("dev");
                                  out.println("<a class=\"w3-button w3-ripple w3-teal\" target=\"_blank\" href="+ dev + ">Get support</a><br>");
                               %>
                      </div>


    </body>
</html>
