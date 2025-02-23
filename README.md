## Краткое описание
Веб-приложение (single page application) для подготовки отчетных данных о применении автоматизированной системы предрейсовых медосмотров (АПРМО) [MedPoint24](https://medpoint24.ru/) клиентами (организациями) Заказчика путем преобразования электронного медицинского журнала в формате файла MS Excel, выгружаемого Заказчиком из личного кабинета АПРМО, в электронные документы офисного приложения MS Word.

## Создано с помощью
Java™ SE Development Kit 10.0.1<br/>
Git - управление версиями<br/>
GitHub - репозиторий<br/>
[JSP](https://projects.eclipse.org/projects/ee4j.jsp) - Java(Jakarta) Server Pages (JSP) — технология, позволяющая создавать веб-страницы и Java-приложения со статическим и динамическим содержимым<br/>
[Apache Maven](https://maven.apache.org/) - сборка, управление зависимостями<br/>
[Apache POI](https://poi.apache.org/) - создание файлов Word и Excel<br/>
[Apache Tomcat](https://tomcat.apache.org/) - контейнер сервлетов (платформа для запуска веб-приложения)<br/>
[Heroku](https://www.heroku.com/) - деплой, хостинг<br/>
! Полный список зависимостей и используемые версии компонентов можно найти в ```pom.xml```

## Примеры
Конечный результат работы веб-приложения представляет собой страницу со ссылками на MS Word файлы, содержащие один из следующих видов отчетов:

<table border=1>
<caption>Отчёт №1</caption>
<tr><th colspan=8>Отчет по {название организации} за фактически проведенные предрейсовые и послерейсовые медицинские осмотры за {месяц буквами} {год цифрами} года</th></tr>
<tr>
  <td>№ п/п</td>
  <td>Число отчетного месяца</td>
  <td>Общее количество мед.осмотров</td>
  <td>Количество предрейсовых мед.осмотров</td>
  <td>Количество мед.осмотров "Допуск"</td>
  <td>Количество мед.осмотров "Не допуск"</td>
  <td>Количество послерейсовых мед.осмотров</td>
  <td>% недопусков</td>  
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  
</tr>
<tr>
  <td> </td>  <td>Итого:</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  
</tr>  
</table>

<table border=1>
<caption>Отчёт №2</caption>
<tr><th colspan=8>Группы риска по артериальному давлению за {месяц буквами} {год цифрами} года</th></tr>
<tr>
  <td>№ п/п</td>
  <td>ФИО сотрудника</td>
  <td>Дата рождения</td>
  <td>Возраст, полных лет</td>
  <td>Среднее АД (все измерения)</td>
  <td>Кол-во повышенных АД из всех измерений</td>
  <td>Среднее АД (измерения с превышением нормы)</td>
  <td>% повышенных АД</td>  
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  
</tr>
</table>
Всего обследовано - {число} чел., проведено исследований - {число} шт. Выявлены факторы риска:<br/> 
- среднее артериальное давление выше 139/89 мм.рт.ст. ({число} чел.),<br/> 
- среднее ЧСС выше 100 уд./мин. ({число} чел.).<br/> 
- возраст 55 лет и старше ({число} чел.).<br/> 
Факторы риска отсутствуют: {число} чел. ({число}% от общего числа сотрудников).<br/><br/>

<table border=1>
<caption>Отчёт №3</caption>
<tr><th colspan=5>Статистика причин недопусков за {месяц буквами} {год цифрами} года</th></tr>
<tr>
  <td>№ п/п</td>
  <td>Причина недопуска</td>
  <td>Количество недопусков</td>
  <td>% от всех недопусков</td>
  <td>% от всех осмотров</td>   
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>    
</tr>
</table>
Всего не допусков: {число} ({число}% от всех осмотров) в т.ч. по мед.причинам: {число} ({число}% от всех осмотров)<br/><br/>


<table border=1>
<caption>Отчёт №4</caption>
<tr><th colspan=5>Реестр медицинских осмотров сотрудников {название организации} за {месяц буквами} {год цифрами} года
</th></tr>
<tr>
  <td>№ п/п</td>
  <td>ФИО сотрудника</td>
  <td>Дата осмотра</td>
  <td>Время осмотра</td>
  <td>Результат</td>   
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>    
</tr>
</table>

<table border=1>
<caption>Отчёт №5</caption>
<tr><th colspan=18>Детализация (по ФИО сотрудников) предрейсовых(предсменных)/послерейсовых(послесменных) медицинских осмотров автоматизированным способом {название организации} за {месяц буквами} {год цифрами} года</th></tr>
<tr>
  <td>№ п/п</td>
  <td>ФИО сотрудника / День месяца</td>
  <td>Общее количество мед.осмотров</td>
  <td>1</td>
  <td>2</td>
  <td>3</td>
  <td>4</td>
  <td>5</td>
  <td>6</td>
  <td>7</td>
  <td>...</td>
  <td>25</td>
  <td>26</td>
  <td>27</td>
  <td>28</td>
  <td>29</td>
  <td>30</td>
  <td>31</td>  
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td> <td> </td>  <td> </td>  <td> </td>
</tr>
<tr>
  <td> </td>  <td>Итого:</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td> <td> </td>  <td> </td>  <td> </td>
</tr>  
</table>

<table border=1>
<caption>Отчёт № 6</caption>
<tr><th colspan=18>Детализация (по ФИО медработников) предрейсовых(предсменных)/послерейсовых(послесменных) медицинских осмотров автоматизированным способом {название организации} за {месяц буквами} {год цифрами} года</th></tr>
<tr>
  <td>№ п/п</td>
  <td>ФИО медработника / День месяца</td>
  <td>Общее количество мед.осмотров</td>
  <td>1</td>
  <td>2</td>
  <td>3</td>
  <td>4</td>
  <td>5</td>
  <td>6</td>
  <td>7</td>
  <td>...</td>
  <td>25</td>
  <td>26</td>
  <td>27</td>
  <td>28</td>
  <td>29</td>
  <td>30</td>
  <td>31</td>  
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td> <td> </td>  <td> </td>  <td> </td>
</tr>
<tr>
  <td> </td>  <td>Итого:</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td> <td> </td>  <td> </td>  <td> </td>
</tr>  
</table>

<table border=1>
<caption>Отчёт № 7</caption>
<tr><th colspan=18>Детализация (по адресам осмотров) предрейсовых(предсменных)/послерейсовых(послесменных) медицинских осмотров автоматизированным способом {название организации} за {месяц буквами} {год цифрами} года</th></tr>
<tr>
  <td>№ п/п</td>
  <td>Адрес установки медоборудования / День месяца</td>
  <td>Общее количество мед.осмотров</td>
  <td>1</td>
  <td>2</td>
  <td>3</td>
  <td>4</td>
  <td>5</td>
  <td>6</td>
  <td>7</td>
  <td>...</td>
  <td>25</td>
  <td>26</td>
  <td>27</td>
  <td>28</td>
  <td>29</td>
  <td>30</td>
  <td>31</td>  
</tr>
<tr>
  <td>&nbsp;</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td> <td> </td>  <td> </td>  <td> </td>
</tr>
<tr>
  <td> </td>  <td>Итого:</td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td>  <td> </td> <td> </td>  <td> </td>  <td> </td>
</tr>  
</table>
