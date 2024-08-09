<!DOCTYPE html>
<html lang="en">
<head>

    <meta http-equiv="Content-Type"; content="text/html"; charset="UTF-8">
    <link rel="stylesheet" href="styles/w3.css">
    <link rel="stylesheet" href="styles/list.css">
    <link rel="stylesheet" href="styles/divtable.css">

    <#if title??>
        <title>${title}</title>
        <#else><title>Заголовок страницы не найден ((</title>
    </#if>

</head>
<body>
 <center>
   <h2>Отчёты сформированы успешно!</h2>
   <br>


   <div class="divTable" style="width: 600px;" >
      <div class="divTableBody">

         <#list table_name_and_link as row>
            <div class="divTableRow">
                <#list row as field>
                  <div class="divTableCell" align="left">${field}</div>
                </#list>
            </div>
         </#list>
      </div>
   </div>





   <br><br>
   <p align="center">☆ ====================  ʕ ᵔᴥᵔ ʔ  ==================== ☆</p>
   <a class="w3-button w3-ripple w3-teal" href="./" >Повторить</a>
 </center>
</body>
</html>