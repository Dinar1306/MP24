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
    <!--<script type="text/javascript">
	            // Make sure cookies are enabled.
	            document.cookie = 'leapforce_cookie_check=enabled; path=/';
				if (document.cookie == '') {
	                document.location='resources/browser_requirements.html';
	            }
	</script>-->
</head>
<body>
   <script>
     document.title = "Сервис подготовки отчётов по меджурналу Medpoint24";
   </script>

      <div class="w3-container w3-center">
        <h3>
            Подготовка отчетов по медосмотрам водителей.
        </h3>
        <!--<span class="w3-left">Заполните поля формы и прикрепите файл в формате DOC или DOCX:</span>-->
        <br>
        <div class="w3-content w3-center">
          <form action="otchet" enctype="multipart/form-data" method="POST">
            <fieldset>
                            <legend>=====&nbsp;Загрузите файл и отметьте вид отчёта&nbsp;=====</legend>
            				<p>Меджурнал из системы MedPoint24</p>
                            <p><input name="file" type="file" id="file" accept=".xlsx" ></p>
            				<br>
            				<!--<p>Отчет, подготовленный вручную</p>
                            <p><input name="file_p" type="file" id="file_p" accept=".xls"></p>-->
                            <p>Вид меджурнала:</p>
                            <table style="margin-left: auto; margin-right: auto; font-size: small;">
                            <tbody>
                            <tr>
                            <td>
                            <p><input id="radio-1" type="radio" name="radio" value="1">
                               <label for="radio-1">из distmed.com</label></p>
                            <p><input id="radio-2" type="radio" name="radio" value="2">
                               <label for="radio-2">из V3 (старого образца)</label></p>
                            <p><input id="radio-3" type="radio" name="radio" value="3">
                               <label for="radio-3">из V3 (универсальный)</label></p>
                            <p><input id="radio-4" type="radio" name="radio" value="4" checked>
                               <label for="radio-4">из V3 (меджурнал)</label></p>
                            <select name="select"  >
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
                            </select> превышения АД 139/89 <br> учитывать в группе риска (табл. 8)
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
                               <a  href="${requestScope.dev}" >Support is here</a> ;)
                      </div>
      </div>
</body>
</html>
