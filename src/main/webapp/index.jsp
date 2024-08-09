<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset="UTF-8">
    <link rel="stylesheet" href="styles/w3.css">
	<link rel="stylesheet" href="styles/list.css">
    <noscript>
        <!-- Check that JavaScript is enabled. -->
        <meta http-equiv="refresh" content="0; url=resources/browser_requirements.html" />
    </noscript>

</head>
<body>
   <script>
     document.title = "Сервис подготовки отчётов по меджурналам.";
   </script>

      <div class="w3-container w3-center">
        <h3>
            Подготовка отчетов по медосмотрам работников.
        </h3>

        <br>
        <div class="w3-content w3-center">
          <form action="otchet" enctype="multipart/form-data" method="POST">
            <fieldset>
                            <legend>=====&nbsp;Загрузите файл меджурнала, отметьте необходимые отчеты, укажите необходимые настройки&nbsp;=====</legend>
            				<!--<p>Меджурнал из системы MedPoint24</p>-->
                            <p><input name="file" type="file" id="file" accept=".xlsx, .xls" ></p>
            				<br>
                            <table style="margin-left: auto; margin-right: auto; font-size: small; line-height: 1.3" cellpadding="3" border="1">
                             <tbody>
                             <tr>
                             <td align="left">
 							<p align="center">МЕДЖУРНАЛ:</p>
                             <p align="center"><input id="mp24" type="radio" name="radio" value="1" checked>
                                <label for="mp24">MedPoint24&nbsp;&nbsp;&nbsp;&nbsp;|</label>
 							   <input id="dimeco" type="radio" name="radio" value="2" disabled>
                                <label for="dimeco">Dimeco&nbsp;&nbsp;&nbsp;&nbsp;|</label>
 							   <input id="medcontrol" type="radio" name="radio" value="3" disabled>
                                <label for="medcontrol">MedControl</label></p>
                             <p></p>
                             <p></p>
                             </td>
                             </tr>
                             <tr>
                             <td align="left">
                             <p align="center">ОТЧЁТЫ:</p>

 							<input type="checkbox" id="adchss" name="adchss" checked />
 							<label for="adchss" alt="Группы риска (АД, ЧСС, возраст)" title="Группы риска (АД, ЧСС, возраст)">Группы риска №1 (АД от 140/90, ЧСС от 100, возраст от 55)</label><br>

 							<input type="checkbox" id="smena" name="smena" disabled />
 							<label for="smena" alt="Группы риска короткий перерыв между сменами)" title="Группы риска (короткий перерыв между сменами)">Группы риска №2 (короткий перерыв между сменами)</label><br>

 							<input type="checkbox" id="reestr" name="reestr" checked />
                            <label for="reestr" alt="Реестр медосмотров (предр., послер. и линейный)." title="Реестр медосмотров (предр., послер. и линейный).">Реестр медосмотров (предр., послер. и линейный)</label><br>

 							<input type="checkbox" id="nedopuski" name="nedopuski" />
 							<label for="nedopuski" alt="Причины отстранений (статистика недопусков)" title="Причины отстранений (статистика недопусков)">Причины отстранений (статистика недопусков)</label><br>

                            <input type="checkbox" id="facticheski" name="facticheski" />
                            <label for="facticheski" alt="Фактические медосмотры (по датам)" title="Фактические медосмотры (по датам)">Фактические медосмотры (по датам)</label><br>

                            <input type="checkbox" id="rabotniki" name="rabotniki" />
                            <label for="rabotniki" alt="Детализация медосмотров (по ФИО осматриваемых)" title="Детализация медосмотров (по ФИО осматриваемых)">Детализация медосмотров (по ФИО осматриваемых)</label><br>

                            <input type="checkbox" id="mediki" name="mediki" />
                            <label for="mediki" alt="Детализация медосмотров (по ФИО медработников)" title="Детализация медосмотров (по ФИО медработников)">Детализация медосмотров (по ФИО медработников)</label><br>

                            <input type="checkbox" id="tochki" name="tochki" />
                            <label for="tochki" alt="Детализация медосмотров (по точкам выпуска)" title="Детализация медосмотров (по точкам выпуска)">Детализация медосмотров (по точкам выпуска)</label><br>

                             </td>
                             </tr>
                             <tr>
                             <td>
 							<p align="center">НАСТРОЙКИ:<br>
                             --------------------------------------------------------<br>
                             от <select name="select_ad"  >
                                 <option value="1">1</option>
                                 <option value="2">2</option>
                                 <option selected value="3">3</option>
                                 <option value="4">4</option>
                                 <option value="5">5</option>
                                 <option value="6">6</option>
                                 <option value="7">7</option>
                                 <option value="8">8</option>
                                 <option value="9">9</option>
                                 <option value="10">10</option>
                             </select> превышений АД или ЧСС для включения в <br> &laquo;Группу риска&raquo;<br>
                             --------------------------------------------------------<br>
 							менее <select name="select_time"  >
                                 <option value="1">1</option>
                                 <option value="2">2</option>
                                 <option value="3">3</option>
                                 <option value="4">4</option>
                                 <option value="5">5</option>
                                 <option value="6">6</option>
                                 <option value="7">7</option>
                                 <option value="8">8</option>
                                 <option value="9">9</option>
                                 <option selected value="10">10</option>
 								<option value="9">11</option>
 								<option value="9">12</option>
                             </select> часов между сменами для включения в <br> &laquo;Группу риска&raquo;<br>
                             --------------------------------------------------------<br>
                             <input type="checkbox" id="unfinished" name="unfinished" checked />
                             <label for="unfinished" alt="Незавершённые осмотры, подписанные медработником" title="Незавершённые осмотры, подписанные медработником">учитывать незавершённые осмотры в меджурнале</label>
                             </p>
                             </td>
                             </tr>
                             </tbody>
                             </table>

            </fieldset>
            <p><button type="submit" >Сформировать</button></p>
            <p>Отчеты, сформированные ранее, доступны <a href="list">здесь</a>.</p>
            <br>
            <p><a href="/resources/04-2021-primer.xlsx" target="_blank">Скачать</a> пример выгрузки меджурнала из distmed.com (для оценки формируемых отчётов).</p>
          </form>
        </div>
                      <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      <div class="w3-container w3-left-align">
                               <!--<jsp:include page="/resources/support.html" />-->
                               <!--<a  href=${requestScope.dev} >Support is here</a> ;)-->
                               <%
                                  String dev = (String)request.getAttribute("dev");
                                  out.println("<a target=\"_blank\" href="+ dev + ">Get support</a>");
                               %>

                      </div>
      </div>
</body>
</html>
