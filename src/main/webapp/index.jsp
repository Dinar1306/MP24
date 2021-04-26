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
     document.title = "Начало работы";
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
                            <legend>Загрузите данные</legend>
            				<p>Меджурнал из системы Medpoint24</p>
                            <p><input name="file" type="file" id="file" accept=".xlsx" ></p>
            				<br>
            				<!--<p>Отчет, подготовленный вручную</p>
                            <p><input name="file_p" type="file" id="file_p" accept=".xls"></p>-->
            </fieldset>
            <p><button type="submit" >Сформировать</button></p>
            <p>Отчеты, сформированные ранее, доступны <a href="list">здесь</a>.</p>
            <br>
            <p><a href="/resources/04-2021-primer.xlsx" target="_blank">Скачать</a> пример выгрузки меджурнала (для оценки формируемых отчётов).</p>
          </form>
        </div>
                      <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      <div class="w3-container w3-left-align">
                            <jsp:include page="/resources/support.html" />
                      </div>
      </div>
</body>
</html>
