package online.ITmed;

import com.ibm.icu.text.Transliterator;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;   //для xlsx
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.http.*;

import java.io.*;
import java.math.BigInteger;
import java.sql.Array;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.*;
import java.time.temporal.ChronoUnit;
import java.util.*;

import static org.apache.poi.xwpf.usermodel.TableRowAlign.CENTER;

@MultipartConfig //запрос может содержать несколько параметров
        (fileSizeThreshold=1024*1024*5, // 5MB
         maxFileSize=1024*1024*10,      // 10MB
         maxRequestSize=1024*1024*50)   // 50MB

public class MainServlet extends HttpServlet {

    //static final String REPORTS_DIR = "otchety";
    static final String copyright = "\u00a9";
    //static final String arrow = "\u21E8";
    //static final String arrow = "\u279C";
    static final String arrow = "\u1F310";
    static final Set<String> vidOsmotra = new HashSet<String>(){{
        add("Предрейсовый (предсменный)");
        add("Послерейсовый (послесменный)");
        add("Линейный");
        add("Профилактический");
    }};
    static String REPORTS_DIR, DEV_NAME, DEV_LINK;
    private static List<String> filesList = new ArrayList<>();
    private List<ReportsTable> spisokOtchetov_v2 = new ArrayList<>();     // список отчетов из списка файлов в папке отчетов
    private String organization = "";
    private String period = "";
    private String god = "";
    private boolean failed = false;
    private int errorStringNumber;
    private String debug = "";
    private String message = "";
    private String[] radiobutton; //вид меджурнала
    //private int radio; //значение переключателя (1-2-3)
    private ArrayList<Integer> allVozrasts = new ArrayList<>(); //таблица возрастов всех водителей
    private String CYRILLIC_TO_LATIN = "Russian-Latin/BGN";


    private class FactTable {
        int obscheeChisloMO;
        int kolichPredreisMO;
        int kolichDopuskov;
        int kolichNedopuskov;
        int kolichPoslerMO;
        float procentNedopuska;

        void setProcentNedopuska() {
            this.procentNedopuska = this.kolichNedopuskov / (float)this.obscheeChisloMO;
        }

        //конструктор
        FactTable(int obscheeChisloMO, int kolichPredreisMO, int kolichDopuskov, int kolichNedopuskov, int kolichPoslerMO) {
            this.obscheeChisloMO = obscheeChisloMO;
            this.kolichPredreisMO = kolichPredreisMO;
            this.kolichDopuskov = kolichDopuskov;
            this.kolichNedopuskov = kolichNedopuskov;
            this.kolichPoslerMO = kolichPoslerMO;
            //this.procentNedopuska = kolichNedopuskov/(float)obscheeChisloMO;
        }

        //конструктор по умолчанию
        FactTable() {
        }
    }

    @Override
    public void init() throws ServletException
    {
        // Загрузка настроек

        Properties prop = new Properties();
        try {
            InputStream input = getServletContext().getResourceAsStream(File.separator + "resources"+ File.separator + "application.properties");
            if (input == null) {
                System.out.println("Sorry, unable to find application.properties");

                //настройки по-умолчанию
                REPORTS_DIR = "otchety";
                DEV_NAME = "MDF-Lab";
                DEV_LINK ="https://u.to/S4tUHA";
                return;
            }
            prop.load(input);

            //get the property's values
            REPORTS_DIR = prop.getProperty("directory.reports");
            DEV_NAME = prop.getProperty("dev.name");
            DEV_LINK = prop.getProperty("url.link");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Override
    public void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        RequestDispatcher requestDispatcher;
        //List<ArrayList<String>> spisokOtchetov = new ArrayList<>();     // список отчетов из списка файлов в папке отчетов

        // gets absolute path of the web application
        String applicationPath = request.getServletContext().getRealPath("");

        // constructs path of the directory to save uploaded file
        String uploadFilePath = applicationPath + File.separator + REPORTS_DIR;

        // Раскладываем адрес на составляющие
        String[] list = request.getRequestURI().split("/");
        //забираем команду
        String action = list[list.length-1];

        //выбираем необходимый JSP в зависимости что нажато
        switch (action) {
            case "list":
                //Получаем список файлов-отчетов в папке с отчетами
                filesList = getFileTree(uploadFilePath);

                if((!filesList.isEmpty())&(filesList.get(0)=="empty")){
                    response.setContentType("text/html");
                    request.setCharacterEncoding ("UTF-8");
                    response.setCharacterEncoding("UTF-8");
                    request.setAttribute("message", "Отчеты отсутствуют.");
                    request.setAttribute("debug", "Папка с отчетами пуста.");
                    request.setAttribute("dev", DEV_LINK);
                    requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                    requestDispatcher.forward(request, response);
                }else{
                    //Готовим таблицу из списка
                    // Назв.орг. | Тип.отч | Период(месяц) |  Дата/время создания | Скачать | Удалить
                    //spisokOtchetov = makeTableFromFilelist(filesList);
                    spisokOtchetov_v2 = makeTableFromFilelist_v2(filesList);
                    response.setContentType("text/html");
                    request.setCharacterEncoding ("UTF-8");
                    response.setCharacterEncoding("UTF-8");
                    request.setAttribute("spisokOtchetov_v2", spisokOtchetov_v2);
                    request.setAttribute("dev", DEV_LINK);
                    requestDispatcher = request.getRequestDispatcher("list.jsp");
                    requestDispatcher.forward(request, response);

                }
                break;
            case "delete":
                //получаем номер отчета для удаления
                Integer id = Integer.valueOf(request.getParameter("id"));
                //если список отчетов не пустой, приступаем к удалению
                if ((spisokOtchetov_v2!=null)&(spisokOtchetov_v2.size()!=0)){
                    try {
                        ReportsTable reportForDelete = spisokOtchetov_v2.get(id); //получаем запись об отчете
                        String downloadLink = reportForDelete.getDownloadLink(); //ссылка для скачивания файла и его название
                        String fileNameForDelete = downloadLink.substring(downloadLink.lastIndexOf(File.separator), downloadLink.length()-12);
                        //удаление самого файла
                        File delFile = new File(uploadFilePath+File.separator+fileNameForDelete);
                        boolean deleted = delFile.delete();
                        if (deleted) {
                            spisokOtchetov_v2.remove(id);
                        } else {
                            request.setAttribute("message", "Отчет удалить не удалось((");
                            request.setAttribute("debug", "-");
                            request.setAttribute("dev", DEV_LINK);
                            requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                            requestDispatcher.forward(request, response);
                            return;
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else {
                    request.setAttribute("message", "Список очетов пуст((");
                    request.setAttribute("debug", "-");
                    request.setAttribute("dev", DEV_LINK);
                    requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                    requestDispatcher.forward(request, response);
                    return;
                    }
                response.setContentType("text/html");
                request.setCharacterEncoding ("UTF-8");
                response.setCharacterEncoding("UTF-8");
                request.setAttribute("spisokOtchetov_v2", spisokOtchetov_v2);
                request.setAttribute("dev", DEV_LINK);
                requestDispatcher = request.getRequestDispatcher("list.jsp");
                requestDispatcher.forward(request, response);
                //String[] delString = allRows.get(id);
                break;
            default:
                request.setAttribute("dev", DEV_LINK);
                request.setAttribute("title", "MP24 Reports Service");
                requestDispatcher = request.getRequestDispatcher("index.jsp");
                requestDispatcher.forward(request, response);
                break;

//        RequestDispatcher requestDispatcher = request.getRequestDispatcher("index.jsp");
//        requestDispatcher.forward(request, response);
        }
    }

    private List<ReportsTable> makeTableFromFilelist_v2(List<String> filesList) {
        List<ReportsTable> result = new ArrayList<>();
        int count = 0;
        for (String stroka: filesList) {
            try {
                ReportsTable reportsTable = new ReportsTable(stroka, count);
                result.add(reportsTable);
            }
            catch (StringIndexOutOfBoundsException e) {
                e.printStackTrace();
            }
            count++;
        }
        return result;
    }

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        //инициализируем переменные и объекты
        String table1FileName = "";                 // название файла Word с отчетной таблицей 1 по датам (для скачивания)
        String table2FileName = "";                 // название файла Word с отчетной таблицей 2 по водителям (для скачивания)
        String table3FileName = "";                 // название файла Word с отчетной таблицей 3 по медсестрам (для скачивания)
        String table4FileName = "";                 // название файла Word с отчетной таблицей 4 по точкам (для скачивания)
        String table5FileName = "";                 // название файла Word с отчетной таблицей 5 реестр осмотров (для скачивания)
        String table6FileName = "";                 // название файла Word с отчетной таблицей 6 причины недопусков (для скачивания)
        String table7FileName = "";                 // название файла Word с отчетной таблицей 7 группы риска (для скачивания)
        //InputStream inputStream;                  // поток чтения для загружаемого файла
        XSSFWorkbook workBookXLSX;                  // объект книги эксель xlsx

        List<ArrayList<String>> list = new ArrayList<>();     // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listPosleReis = new ArrayList<>(); // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listLine = new ArrayList<>(); // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listProf = new ArrayList<>(); // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listPosleAndLine = new ArrayList<>(); // для объединения послерейса и линейного
        List<ArrayList<String>> listPredreis = new ArrayList<>();    // массив строк листа (каждая строка - массив строк) для списка предрейсовых осмотров
        TreeMap<Integer, Integer[]> medOsmotryByDatesPredReis = new TreeMap<Integer, Integer[]>(); //итоговые данные отсортированы по дате
        //т.е. здесть Integer Key - дата мед.осм.
        //Integer[] Value - таблица допущено / не допущено (в эту дату)
        TreeMap<Integer, Integer[]> medOsmotryByDatesPosleReis;
        TreeMap<Integer, FactTable> medOsmotryByDatesFacticheskie = new TreeMap<Integer, FactTable>();; // Таблица 1 для ворда
        //т.е. здесть Integer Key - дата мед.осм.
        //FactTable Value - таблица: общ.число.МО|кол.предр.МО|допусков|недопусков|кол.послер.МО|%невыпуска (в эту дату)

        TreeMap<String, DriverRiskData> gruppyRiskaByFIO = new TreeMap<String, DriverRiskData>();
        //т.е. здесть Integer Key - дата мед.осм.
        //int[] Value - таблица: общ.кол|предр|допущ|недоущ|послер| (в эту дату) --> добавить столбец %невыпуска

        int chisloMO = 0; //общее число медосмотров из трех списков(предр, послер и линейный)
        int chisloPredr = 0; //общее число предр.медосмотров
        int chisloPosler = 0; //общее число послер.медосмотров
        int chisloLine = 0; //общее число линейн.медосмотров

        //Массив дат медосмотров (для Табл.№2)
        ArrayList<Integer> dates = new ArrayList<>();

        //итоговые данные отсортированы по фамилиям и дате
        TreeMap<String, int[]> medOsmotryByFIOXLS;
        TreeMap<String, int[]> medOsmotryByFIO = new TreeMap<String, int[]>();
        TreeMap<String, int[]> medRabotnikByFIO = new TreeMap<String, int[]>();
        TreeMap<String, int[]> medOsmByHost = new TreeMap<String, int[]>();
        TreeMap<String, int[]> medOsmByNepoduski;
        // здесь key   это ФИО водителя - String
        // здесь value это таблица с суммарным значением предрейса и послерейса в каждой ячейке,
        // причем длина массива равна длине массива дат dates

        //инициализия завершена

        ///////////////WORK///////////////////
        //получаем части (нужные нам файлы)
        Part part = request.getPart("file");
        long size = part.getSize(); //файл медпойнта


        //получаем radiobutton (вид меджурнала: 1 - из дистмед, 2 - старый из V3, 3 - из V3)
        radiobutton = request.getParameterValues("radio");


        //проверям загруженли файл меджурнала:
        //ничего
        if (size == 0){
            request.setAttribute("title", "Error :((");
            request.setAttribute("message", "Загрузите файл меджурнала!");
            request.setAttribute("dev", DEV_LINK);
            RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
            requestDispatcher.forward(request, response);
            return;
        }
        // меджурнал medpoint24 загружен
        else {
            //получаем объект книги XLSX из формы
            workBookXLSX = XLSXFromPart(part);

            //выбираем вид меджурнала
            int radio = Integer.parseInt(radiobutton[0]); //значение переключателя (1-2-3)
            switch (radio){
                case 1 : { //меджурнал из distmed
                    try {
                        //разбираем первый лист файла medpoint24 на объектную модель
                        listPredreis = getListFromSheet(workBookXLSX, 0); //получаем лист предрейса
                        listPosleReis = getListFromSheet(workBookXLSX, 1); //получаем лист послерейса
                        listLine = getListFromSheet(workBookXLSX, 5); //получаем лист линейного
                        ArrayList<String> pervayaStroka = listPredreis.get(0); //первая строка (заголовок)
                        organization = getOrganizationName_v2(pervayaStroka); //достаем из первой строки (заголовка) название компании.
                        period = getMonth_v2(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
                        god = getGod_v2(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.
                    } catch (Exception e) {
                        e.printStackTrace();
                        request.setAttribute("title", "Error!");
                        request.setAttribute("message", "При обработке файла произошла ошибка.");
                        request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                        request.setAttribute("dev", DEV_LINK);
                        RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                        requestDispatcher.forward(request, response);
                        return;
                    }

                    //Причесываем списки:
                    // убираем заголовок таблицы, убираем шапку таблицы, убирем последние 5 и 7 ненужных строк из предрейса и послерейса соответственно
                    listPredreis = listPredreis.subList(2, listPredreis.size()-5);
                    listPosleReis = listPosleReis.subList(2, listPosleReis.size()-7);
                    //причесываем линейный
                    listLine = listLine.subList(2, listLine.size()-5);

                    //считаем общее число медосмотров
                    if (!listPredreis.isEmpty()) {
                        chisloPredr = listPredreis.size();
                        chisloMO = chisloMO + chisloPredr;
                    }
                    if (!listPosleReis.isEmpty()){
                        chisloPosler = listPosleReis.size();
                        chisloMO = chisloMO + chisloPosler;
                    }
                    if (!listLine.isEmpty()){
                        chisloLine = listLine.size();
                        chisloMO = chisloMO + chisloLine;
                    }

                    ////соединяем послерейс и линейные МО
                    //Объединение двух списков в третий:
                    listPosleAndLine.addAll(listPosleReis);
                    listPosleAndLine.addAll(listLine);

                    break;
                } /////////////case 1
                case 2 : {  //меджурнал старого образца
                    try {
                        //разбираем первый(единственный) лист файла medpoint24 на объектную модель
                        list = getListFromSheet(workBookXLSX, 0); //получаем лист всех видов осмотра
                        //listPosleReis = getListFromSheet(workBookXLSX, 1); //получаем лист послерейса
                        //listLine = getListFromSheet(workBookXLSX, 5); //получаем лист линейного
                        ArrayList<String> pervayaStroka = list.get(0); //первая строка (заголовок)
                        organization = getOrganizationName(pervayaStroka); //достаем из первой строки (заголовка) название компании.
                        period = getMonth_v3(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
                        god = getGod_v3(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.
                    } catch (Exception e) {
                        e.printStackTrace();
                        request.setAttribute("message", "При обработке файла произошла ошибка.");
                        request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                        request.setAttribute("dev", DEV_LINK);
                        RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                        requestDispatcher.forward(request, response);
                        return;
                    }
                    //Причесываем списки:
                    // убираем заголовок таблицы, убираем шапку и последние 3 ненужные строки
                    list = list.subList(2, list.size()-3); //общий список со всеми видами осмотров
                    //считаем общее число медосмотров
                    if (!list.isEmpty()) {
                        chisloMO = list.size();
                    }
                    //получаем список предрейса
                    listPredreis = getPredreisList(list);
                    listPosleReis = getPoslereisList(list);
                    listLine = getLineList(list);

                    if (!listPosleReis.isEmpty()){
                        chisloPosler = listPosleReis.size();
                    }
                    if (!listLine.isEmpty()){
                        chisloLine = listLine.size();
                    }

                    ////соединяем послерейс и линейные МО
                    //Объединение двух списков в третий:
                    listPosleAndLine.addAll(listPosleReis);
                    listPosleAndLine.addAll(listLine);

                    break;
                } /////////////case 2
                case 3 : { //меджурнал V3 (Универсальный)
                    //количество листов в книге
                    int sheets = workBookXLSX.getNumberOfSheets();

                    //проходим по всем листам
                    for (int i = 0; i<sheets; i++){
                        try {
                            String sheetName = workBookXLSX.getSheetName(i);
                            switch (sheetName) {
                                case "Предрейсовый (предсменный)":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист предрейса
                                    listPredreis = getPredreisList_V3(list); //
                                    break;
                                case "Послерейсовый (послесменный)":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист послерейса
                                    listPosleReis = getPoslereisList(list); //
                                    break;
                                case "Линейный":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист послерейса
                                    listLine = getLineList(list); //
                                    break;
                                case "Профилактический":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист послерейса
                                    listProf = getProfList(list); //ИСПРАВИТЬ на новый геттер из list
                                    break;
                                default: // любое др. назв. (пропуск)
                                    //tempStringArray.add("");
                                    break;
                            } //switch

                        } catch (Exception e){
                            request.setAttribute("message", "При обработке файла произошла ошибка.");
                            request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                            request.setAttribute("dev", DEV_LINK);
                            RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                            requestDispatcher.forward(request, response);
                            return;
                        }

                    } // for i

                    ////соединяем послерейс и линейные МО
                    //Объединение двух списков в третий:
                    listPosleAndLine.addAll(listPosleReis);
                    listPosleAndLine.addAll(listLine);

                    // достать оргу, год, месяц из не пустых списков
                    ArrayList<String> pervayaStroka = list.get(0); //первая строка (заголовок)
                    organization = getOrganizationName(pervayaStroka); //достаем из первой строки (заголовка) название компании.
                    period = getMonth_v3(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
                    god = getGod_v3(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.

                    break;
                } /////////////case 3
                case 4 : { //меджурнал V3 (Меджурнал XLS)
                    //количество листов в книге
                    int sheets = workBookXLSX.getNumberOfSheets();

                    //проходим по всем листам
                    for (int i = 0; i<sheets; i++){
                        try {
                            String sheetName = workBookXLSX.getSheetName(i);
                            switch (sheetName) {
                                case "Предрейсовый (предсменный)":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист предрейса
                                    listPredreis = getPredreisList_V3(list); //
                                    break;
                                case "Послерейсовый (послесменный) вн":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист послерейса
                                    listPosleReis = getPoslereisList(list); //
                                    break;
                                case "Линейный":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист послерейса
                                    listLine = getLineList(list); //
                                    break;
                                case "Профилактический":
                                    list = getListFromSheet(workBookXLSX, i); //получаем лист послерейса
                                    listProf = getProfList(list); //ИСПРАВИТЬ на новый геттер из list
                                    break;
                                default: // любое др. назв. (пропуск)
                                    //tempStringArray.add("");
                                    break;
                            } //switch

                        } catch (Exception e){
                            request.setAttribute("message", "При обработке файла произошла ошибка.");
                            request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                            request.setAttribute("dev", DEV_LINK);
                            RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                            requestDispatcher.forward(request, response);
                            return;
                        }

                    } // for i

                    ////соединяем послерейс и линейные МО
                    //Объединение двух списков в третий:
                    listPosleAndLine.addAll(listPosleReis);
                    listPosleAndLine.addAll(listLine);

                    // достать оргу, год, месяц из не пустых списков
                    ArrayList<String> pervayaStroka = list.get(0); //первая строка (заголовок)
                    organization = getOrganizationName(pervayaStroka); //достаем из первой строки (заголовка) название компании.
                    period = getMonth_v3(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
                    god = getGod_v3(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.

                    break;
                } /////////////case 4
            }

            //производим подсчёт по предрейсовым
            medOsmotryByDatesPredReis = prepare(listPredreis);

            //производим подсчёт по объединенному послерейсу и линейному (новый вариант)
            medOsmotryByDatesPosleReis = prepare(listPosleAndLine); //новый вариант

            // (Табл.1 Фактические медосмотры)
            medOsmotryByDatesFacticheskie = prepareTable1(medOsmotryByDatesPredReis, medOsmotryByDatesPosleReis);

            // считаем проценты недопусков в табл.1
            for (Map.Entry<Integer, FactTable> entry: medOsmotryByDatesFacticheskie.entrySet()) {
                entry.getValue().setProcentNedopuska();
            }

            //получаем массив дат
            //for ( Integer keys:medOsmotryByDatesALL.keySet() ) {
            for ( Integer keys:medOsmotryByDatesFacticheskie.keySet() ) {
                dates.add(keys);
            }
            // (Табл.2 Детализация, по водителям) предрейс+послерейс, нужен 6й столбец
            medOsmotryByFIO = prepareTable2(listPredreis, listPosleAndLine, dates, 6);

            // (Табл.3 Детализация, по медсестрам) предрейс+послерейс, нужен 18й столбец
            medRabotnikByFIO = prepareTable2(listPredreis, listPosleAndLine, dates, 18);

            // (Табл.4 Детализация, по точкам) предрейс+послерейс, нужен 4й столбец
            medOsmByHost = prepareTable2(listPredreis, listPosleAndLine, dates, 4);

            // (Табл.7 Группы риска)
            try {
                gruppyRiskaByFIO = prepareTableGruppyRiska(listPredreis, listPosleAndLine);
            } catch (Exception e) {
                e.printStackTrace();
                request.setAttribute("message", "При обработке файла произошла ошибка.");
                request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                request.setAttribute("dev", DEV_LINK);
                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                requestDispatcher.forward(request, response);
                return;
            }

            // gets absolute path of the web application
            String applicationPath = request.getServletContext().getRealPath("");
            // constructs path of the directory to save uploaded file
            String uploadFilePath = applicationPath + File.separator + REPORTS_DIR;

            //Создаем папку для формируемых отчетов Word если ее нет
            File uploadFolder = new File(uploadFilePath);
            if (!uploadFolder.exists()) {  //если папки не существует, то создаем
                uploadFolder.mkdirs();
            }

            try {   //заменить на суммарый с послерейсом +(готово)
                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table1FileName) в JSP
                //table1FileName = makeWordDocumentTable1(medOsmotryByDatesALL, uploadFilePath, medOsmotryByDatesAllProcent);
                table1FileName = makeWordDocumentTable1XLS(medOsmotryByDatesFacticheskie, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table2FileName)
                table2FileName = makeWordDocumentTable2XLS("водит.", dates, medOsmotryByFIO, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table3FileName)
                table3FileName = makeWordDocumentTable2XLS("медраб.", dates, medRabotnikByFIO, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table4FileName)
                table4FileName = makeWordDocumentTable2XLS("точки осм.", dates, medOsmByHost, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table5FileName)
                table5FileName = makeWordDocumentReestr(listPredreis, listPosleReis, listLine, listProf, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table6FileName)
                table6FileName = makeWordDocumentStatNedopuskov(listPredreis, listPosleAndLine, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table7FileName)
                table7FileName = makeWordDocumentGruppaRiska(gruppyRiskaByFIO, uploadFilePath);

            } catch (XmlException e) {
                e.printStackTrace();
                request.setAttribute("message", "При обработке файла произошла ошибка.");
                request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                request.setAttribute("dev", DEV_LINK);
                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                requestDispatcher.forward(request, response);
                return;
                //response.setContentType("text/html");
            } catch (InterruptedException e) {
                e.printStackTrace();
                request.setAttribute("message", "При обработке файла произошла ошибка.");
                request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                request.setAttribute("dev", DEV_LINK);
                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                requestDispatcher.forward(request, response);
                return;
            } catch (IOException e) {
                e.printStackTrace();
                request.setAttribute("message", "Слишком длинное название конечного файла.");
                request.setAttribute("debug", ExceptionUtils.getStackTrace(e));
                request.setAttribute("dev", DEV_LINK);
                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                requestDispatcher.forward(request, response);
                return;
            }

            response.setContentType("text/html");
            response.setCharacterEncoding("UTF-8");
            request.setCharacterEncoding("UTF-8");
            request.setAttribute("docxName", table1FileName);
            request.setAttribute("docx2Name", table2FileName);
            request.setAttribute("docx3Name", table3FileName);
            request.setAttribute("docx4Name", table4FileName);
            request.setAttribute("docx5Name", table5FileName);
            request.setAttribute("docx6Name", table6FileName);
            request.setAttribute("docx7Name", table7FileName);
            request.setAttribute("reportsDir", REPORTS_DIR);
            request.setAttribute("message", "Отчёты сформированы успешно!");
            request.setAttribute("dev", DEV_LINK);
            RequestDispatcher requestDispatcher = request.getRequestDispatcher("otchet.jsp");
            requestDispatcher.forward(request, response);
            return;
        }

    }



    ////////////////////////////////////////////////////////////////////////
    //                      ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ                        //
    ////////////////////////////////////////////////////////////////////////

    //получаем список предрейсовых медосмотров
    private List<ArrayList<String>> getPredreisList(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        for (ArrayList strArr: list ) {
            if ((strArr.get(7).equals("Предрейсовый осмотр"))|(strArr.get(7).equals("Предсменный осмотр"))){
                res.add(strArr);
            }
        }
        return convertToOldFormat(res);
    }

    //получаем список предрейсовых медосмотров из Универсального отчета
    private List<ArrayList<String>> getPredreisList_V3(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        List<ArrayList<String>> out = new ArrayList<>();
        for (ArrayList strArr: list ) {
            //выбор вида отчета: универсальный V3 или меджурнал V3
            switch (Integer.parseInt(radiobutton[0])) { //значение переключателя вида файла меджурнала
                case 3: //универсальный
                    if ((strArr.get(4).equals("Предрейсовый осмотр"))|(strArr.get(4).equals("Предсменный осмотр"))){
                        res.add(strArr);
                        out = convertToOldFormat_V3(res);
                    }
                    break;
                case 4: //меджурнал
                    if ((strArr.get(3).equals("Предрейсовый осмотр"))|(strArr.get(3).equals("Предсменный осмотр"))){
                        res.add(strArr);
                        out = convertToOldFormat_V3_M(res);
                    }
                    break;
                default: // любое др. назв. (пропуск)
                    //пусто
                    break;
            } //switch

        }
        return out;
    }

    private List<ArrayList<String>> getPoslereisList(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        List<ArrayList<String>> out = new ArrayList<>();
        int nomerStolbtca;
        int r = Integer.parseInt(radiobutton[0]);
        switch (r){
            case 3:
                nomerStolbtca = 4;
                break;
            case 4:
                nomerStolbtca = 3;
                break;
            default:
                nomerStolbtca = 7;
        }

        for (ArrayList strArr: list ) {
            if ((strArr.get(nomerStolbtca).equals("Послерейсовый осмотр"))|(strArr.get(nomerStolbtca).equals("Послесменный осмотр"))){
                res.add(strArr);
            }
        }

        switch (r){
            case 3:
                out = convertToOldFormat_V3(res);
                break;
            case 4:
                out = convertToOldFormat_V3_M(res);
                break;
            default:
                out = convertToOldFormat_V3(res);
        }

        return  out;
    }

    private List<ArrayList<String>> getLineList(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        List<ArrayList<String>> out = new ArrayList<>();
        int nomerStolbtca;
        int r = Integer.parseInt(radiobutton[0]);
        switch (r){
            case 3:
                nomerStolbtca = 4;
                break;
            case 4:
                nomerStolbtca = 3;
                break;
            default:
                nomerStolbtca = 7;
        }

        for (ArrayList strArr: list ) {
            if (strArr.get(nomerStolbtca).equals("Линейный осмотр")){
                res.add(strArr);
            }
        }

        switch (r){
            case 3:
                out = convertToOldFormat_V3(res);
                break;
            case 4:
                out = convertToOldFormat_V3_M(res);
                break;
            default:
                out = convertToOldFormat_V3(res);
        }

        return  out;
    }

    private List<ArrayList<String>> getProfList(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        for (ArrayList strArr: list ) {
            if (strArr.get(4).equals("Алкотестирование")){ //могут быть и др. виды проф.осмотров, например "Термометрия"
                res.add(strArr);
            }
        }
        return res;
    }

    private List<ArrayList<String>> convertToOldFormat(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        ArrayList<String> converted = new ArrayList<>();
        String zakl = "";
        for (ArrayList<String> strArr: list ) {
            converted.add(0, strArr.get(0)); // № п/п
            converted.add(1, convertDate(strArr.get(1))); //Дата и время осмотра
            converted.add(2, strArr.get(2)); // Длительность осмотра (на терминале)
            converted.add(3, strArr.get(7)); // Тип осмотра
            converted.add(4, strArr.get(19)); // Место осмотра
            converted.add(5, strArr.get(6)); // Табельный номер
            converted.add(6, strArr.get(3)); // ФИО работника
            converted.add(7, strArr.get(4)); // Пол
            converted.add(8, convertBD(strArr.get(5))); // Дата рождения
            converted.add(9, strArr.get(12));  // Жалобы
            converted.add(10, "пусто"); // Осмотр
            converted.add(11, strArr.get(8)); // АД
            converted.add(12, strArr.get(9)); // ЧСС
            converted.add(13, strArr.get(11)); // температура
            converted.add(14, strArr.get(10)); // Проба на наличие алкоголя
            if (strArr.get(14).equals("Допущен")||strArr.get(14).equals("Прошёл")) zakl = "О"; else zakl = "Н";
            converted.add(15, zakl); // Заключение (Н или О)*
            converted.add(16, strArr.get(14)); // Результат
            converted.add(17, strArr.get(15)); // Комментарий
            converted.add(18, strArr.get(16)); // ФИО медицинского работника
            converted.add(19, strArr.get(17)); // Подпись медицинского работника
            converted.add(20, strArr.get(18)); // Подпись работника

            //конвертировано, добавляем (кроме закрытых ботом)
            if (!converted.get(18).contains("Бот оповещения")){
                res.add((ArrayList<String>)converted.clone());
            }
            converted.clear();
        }
        return res;
    }

    private List<ArrayList<String>> convertToOldFormat_V3(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        ArrayList<String> converted = new ArrayList<>();
        String zakl = "";
        for (ArrayList<String> strArr: list ) {
            converted.add(0, strArr.get(0)); // № п/п
            converted.add(1, convertDate_V3(strArr.get(2))); //Дата и время осмотра
            converted.add(2, strArr.get(3)); // Длительность осмотра (на терминале)
            converted.add(3, strArr.get(4)); // Тип осмотра
            converted.add(4, strArr.get(5)); // Место осмотра
            converted.add(5, strArr.get(6)); // Табельный номер
            converted.add(6, strArr.get(7)); // ФИО работника
            converted.add(7, strArr.get(8)); // Пол
            converted.add(8, convertBD_V3(strArr.get(9))); // Дата рождения
            converted.add(9, strArr.get(10));  // Жалобы
            converted.add(10, strArr.get(11)); // Осмотр
            converted.add(11, convertAD_V3(strArr.get(12),strArr.get(13))); // АД
            converted.add(12, convertPuls_V3(strArr.get(14))); // ЧСС
            converted.add(13, strArr.get(15)); // температура
            converted.add(14, strArr.get(16)); // Проба на наличие алкоголя
            //if (strArr.get(14).equals("Допущен")||strArr.get(14).equals("Прошёл")) zakl = "О"; else zakl = "Н";
            converted.add(15, strArr.get(17)); // Заключение (Н или О)*
            converted.add(16, strArr.get(18)); // Результат
            converted.add(17, strArr.get(19)); // Комментарий
            converted.add(18, strArr.get(20)); // ФИО медицинского работника
            converted.add(19, strArr.get(21)); // Подпись медицинского работника
            converted.add(20, strArr.get(22)); // Подпись работника

            //конвертировано, добавляем (кроме закрытых ботом)
            if (!converted.get(18).contains("Бот оповещения")){
                res.add((ArrayList<String>)converted.clone());
            }
            converted.clear();
        }
        return res;
    }


    private List<ArrayList<String>> convertToOldFormat_V3_M(List<ArrayList<String>> list) {
        List<ArrayList<String>> res = new ArrayList<>();
        ArrayList<String> converted = new ArrayList<>();
        String zakl = "";
        for (ArrayList<String> strArr: list ) {
            converted.add(0, strArr.get(0)); // № п/п
            converted.add(1, strArr.get(2)); //Дата и время осмотра
            converted.add(2, ""); // Длительность осмотра (на терминале)
            converted.add(3, strArr.get(3)); // Тип осмотра
            converted.add(4, strArr.get(4)); // Место осмотра
            converted.add(5, strArr.get(5)); // Табельный номер
            converted.add(6, strArr.get(6)); // ФИО работника
            converted.add(7, strArr.get(7)); // Пол
            converted.add(8, convertBD_V3(strArr.get(8))); // Дата рождения
            converted.add(9, strArr.get(9));  // Жалобы
            converted.add(10, strArr.get(10)); // Осмотр
            converted.add(11, convertAD_V3(strArr.get(11),strArr.get(12))); // АД
            converted.add(12, convertPuls_V3(strArr.get(13))); // ЧСС
            converted.add(13, strArr.get(14)); // температура
            converted.add(14, strArr.get(15)); // Проба на наличие алкоголя
            //if (strArr.get(14).equals("Допущен")||strArr.get(14).equals("Прошёл")) zakl = "О"; else zakl = "Н";
            converted.add(15, strArr.get(16)); // Заключение (Н или О)*
            converted.add(16, strArr.get(17)); // Результат
            converted.add(17, strArr.get(18)); // Комментарий
            converted.add(18, strArr.get(19)); // ФИО медицинского работника
            converted.add(19, strArr.get(20)); // Подпись медицинского работника
            converted.add(20, strArr.get(21)); // Подпись работника

            //конвертировано, добавляем (кроме закрытых ботом и незавершенных)
            if (
                   ! ( (converted.get(18).contains("Бот оповещения"))
                            |
                    (converted.get(17).contains("Незавершенный")))
            ){
                res.add((ArrayList<String>)converted.clone());
            }
            converted.clear();
        }
        return res;
    }

    //конвертирование даты с "29.07.2022 15:27:46" на "2022-07-29 15:27"
    private String convertDate (String s){
        String res = "";
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
        SimpleDateFormat newFormatter = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        //newFormatter.setTimeZone(TimeZone.getTimeZone("UTC+5")); //перевод на местное время
        try {
            Date date = formatter.parse(s);
            ZonedDateTime d = ZonedDateTime.ofInstant(date.toInstant(),  ZoneId.systemDefault()); //не важно какой часовой пояс на хосте программы, т.к. ко времени из
            // журнала надо прибавить два часа (в журнале московское время)
            LocalDateTime ldt = d.toLocalDateTime();
            ldt = ldt.plusHours(2); //+2 часа к московскому времени из журнала
            res = newFormatter.format(Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant()));

        } catch (ParseException e) {
            e.printStackTrace();
        }

        return res;
    }

    private String convertDate_V3 (String st_dt){
        String res = "";
        String s = st_dt.substring(0,19); // обрезка " (+5)"
        String timeZone = st_dt.substring(21,23);
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
        SimpleDateFormat newFormatter = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        //newFormatter.setTimeZone(TimeZone.getTimeZone("UTC+5")); //перевод на местное время
        try {
            Date date = formatter.parse(s);
            ZonedDateTime d = ZonedDateTime.ofInstant(date.toInstant(),  ZoneId.systemDefault()); //не важно какой часовой пояс на хосте программы, т.к. ко времени из
            // журнала надо прибавить два часа (в журнале московское время)
            LocalDateTime ldt = d.toLocalDateTime();
            if (timeZone.equals("+3")) ldt = ldt.plusHours(2); //+2 часа к московскому времени из журнала
            res = newFormatter.format(Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant()));

        } catch (ParseException e) {
            e.printStackTrace();
        }

        return res;
    }

    //конвертирование даты рождения с "2022-07-29" на "29-07-2022"
    private String convertBD (String s){
        String res = "";
        String[] tempArray = s.split("-");
        res=tempArray[2]+"-"+tempArray[1]+"-"+tempArray[0];
        return res;
    }

    //конвертирование даты рождения с "29.07.2022" на "2022-07-29"
    private String convertBD_V3 (String s){
        String res = "";
        String[] tempArray = s.split("\\.");
        res=tempArray[0]+"-"+tempArray[1]+"-"+tempArray[2];
        return res;
    }

    private String convertAD_V3 (String sad, String dad){
        String res = "";
        CharSequence dot = ".";
        if (sad.contains(dot)){
            String[] tempo = sad.split("\\.");
            sad = tempo[0];
        }
        if (dad.contains(dot)){
            String[] tempo = dad.split("\\.");
            dad = tempo[0];
        }
        res = sad+"/"+dad;
        return res;
    }

    private String convertPuls_V3 (String puls){
        String res = "";
        CharSequence dot = ".";
        if (puls.contains(dot)){
            String[] tempo = puls.split("\\.");
            res = tempo[0];
        } else res = puls;
        return res;
    }

    //получаем объект книги xlsx
    private XSSFWorkbook XLSXFromPart(Part part){
        InputStream inputStream;
        XSSFWorkbook workBook = new XSSFWorkbook();
        try {
            inputStream = part.getInputStream();
            workBook = new XSSFWorkbook(inputStream);
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
        }
        return workBook;
    }

    //получаем лист из книги xlsx
    private List<ArrayList<String>> getListFromSheet(XSSFWorkbook workBook, int num) throws IOException { //разбираем первый лист входного файла на объектную модель
        List<ArrayList<String>> res = new ArrayList<>();

        Sheet sheet = workBook.getSheetAt(num);

        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
            while (it.hasNext()) {
                ArrayList<String> tempStringArray = new ArrayList<>();
                Row row = it.next();
                Iterator<Cell> cells = row.iterator();
                while (cells.hasNext()) {

                    Cell cell = cells.next();
                    CellType cellType = cell.getCellType();

                    switch (cellType) {
                        case /*Cell.CELL_TYPE_STRING*/STRING:
                            tempStringArray.add(cell.getStringCellValue());
                            break;
                        case /*Cell.CELL_TYPE_NUMERIC*/NUMERIC:
                            tempStringArray.add(Double.toString(cell.getNumericCellValue()));
                            break;
                        case /*Cell.CELL_TYPE_FORMULA*/FORMULA:
                            tempStringArray.add(Double.toString(cell.getNumericCellValue()));
                            break;
                        default: // пустая ячейка
                            tempStringArray.add("");
                            break;
                    }
                }
                res.add(tempStringArray);
            }
            workBook.close();

            //chisloMO = 0; //общее число медосмотров из трех списков(предр, послер и линейный)
        return res;
    }

    private Integer getDate (String data){
        int r = Integer.parseInt(radiobutton[0]);
        String[] s1 = data.split(" "); // 2020-08-31 18:06 делим по пробелу
        Integer res = 0;

        switch (r){
            case 4:
                String[] s4 = s1[0].split("\\."); // 31.08.2020  делим дату месяц год
                res = Integer.parseInt(s4[0]); //забираем дату
                break;
            default:
                String[] s3 = s1[0].split("-"); // 2020-08-31  делим дату месяц год
                res = Integer.parseInt(s3[2]); //забираем дату
        }

        return res;
    }

    private TreeMap<Integer, Integer[]> prepare (List<ArrayList<String>> spisokVes){
        //заготовка для результата
        TreeMap<Integer, Integer[]> result = new TreeMap<Integer, Integer[]>();
        int r = Integer.parseInt(radiobutton[0]);

        switch (r){
            case 4:
                String control = new String("НЕ допущен");
                String control2 = new String("выявлены");

                // foreach
                for (ArrayList<String> stroka : spisokVes) { //пробегаемся по строкам
                    Integer data = getDate(stroka.get(1)); // получаем дату из второй ячейки строки

                    //определяем Допущен или Не допущен и увеличиваем счетчик в соответствующей ячейке (первой или второй)


                        if (stroka.get(16).toLowerCase().contains(control.toLowerCase()) | stroka.get(16).toLowerCase().contains(control2.toLowerCase())){ //если не допущен
                            //нашелся Не допуск -> увеличиваем значение во второй ячейке
                            if ((result.get(data)==null))       // если эта дата еще не внесена
                            {
                                result.put(data, new Integer[] {0, 1}); //добавляем текущую строку (ключ) и счетчик (первое нахождение)
                            } else {
                                Integer[] v = result.get(data); //получаем значение счетчика допущенных (нужна будет первая ячейка)
                                v[1]++;                         // и увеличиваем
                                result.put(data, v);            // перезаписываем счетчик
                            }
                        }
                        else {
                            //нашелся допуск -> увеличиваем значение в первой ячейке
                            if ((result.get(data)==null))       // если эта дата еще не внесена
                            {
                                result.put(data, new Integer[] {1, 0}); //добавляем текущую строку (ключ) и счетчик (первое нахождение)
                            } else {
                                Integer[] v = result.get(data); //получаем значение счетчика допущенных (нужна будет первая ячейка)
                                v[0]++;                         // и увеличиваем
                                result.put(data, v);            // перезаписываем счетчик
                            }
                        }


                }
                break;
            default:
                // foreach
                for (ArrayList<String> stroka : spisokVes) { //пробегаеся по строкам
                    Integer data = getDate(stroka.get(1)); // получаем дату из второй ячейки строки

                    //определяем Допущен или Не допущен и увеличиваем счетчик в соответствующей ячейке (первой или второй)
                    switch (stroka.get(16)){ //было 15
                        case "Допущен":
                        case "Прошел":
                        case "Прошёл":
                            //нашелся допуск -> увеличиваем значение в первой ячейке
                            if (stroka.get(15).equals("О")){
                                if ((result.get(data)==null))       // если эта дата еще не внесена
                                {
                                    result.put(data, new Integer[] {1, 0}); //добавляем текущую строку (ключ) и счетчик (первое нахождение)
                                } else {
                                    Integer[] v = result.get(data); //получаем значение счетчика допущенных (нужна будет первая ячейка)
                                    v[0]++;                         // и увеличиваем
                                    result.put(data, v);            // перезаписываем счетчик
                                }
                            } else { //нашелся Не допуск
                                if ((result.get(data)==null))       // если эта дата еще не внесена
                                {
                                    result.put(data, new Integer[] {0, 1}); //добавляем текущую строку (ключ) и счетчик (первое нахождение)
                                } else {
                                    Integer[] v = result.get(data); //получаем значение счетчика допущенных (нужна будет первая ячейка)
                                    v[1]++;                         // и увеличиваем
                                    result.put(data, v);            // перезаписываем счетчик
                                }
                            }
                            break;
                        case "Не допущен":
                        case "Не прошел":
                        case "Не прошёл":
                            //нашелся Не допуск -> увеличиваем значение во второй ячейке    daNet.equals("прошел") & zapis.get(15).equals("Н")
                            if ((result.get(data)==null))       // если эта дата еще не внесена
                            {
                                result.put(data, new Integer[] {0, 1}); //добавляем текущую строку (ключ) и счетчик (первое нахождение)
                            } else {
                                Integer[] v = result.get(data); //получаем значение счетчика допущенных (нужна будет первая ячейка)
                                v[1]++;                         // и увеличиваем
                                result.put(data, v);            // перезаписываем счетчик
                            }
                            break;
                        default:
                            //nothing to do
                            break;
                    }
                }
        }

        return result;
    }

    private int getIntFromFloatString (String floatString){
        float f = Float.parseFloat(floatString);
        return (int) f; // int
    }

    private TreeMap<Integer, FactTable> prepareTable1(TreeMap<Integer, Integer[]> pred, TreeMap<Integer, Integer[]> posl){
        TreeMap<Integer, FactTable> res = new TreeMap<Integer, FactTable>();

        //проходим по предрейсу
        //результат пуст, поэтому сразу добавляем без проверки наличия добавляемой даты
        for (Map.Entry<Integer, Integer[]> entry: pred.entrySet()) {
            Integer key = entry.getKey(); //получаем дату
            Integer[] dopuskNedopusk = entry.getValue();                //значения допуска/недопуска в эту дату
            res.put(key,
                    new FactTable(
                            dopuskNedopusk[0]+dopuskNedopusk[1], //общ.число.МО
                            dopuskNedopusk[0]+dopuskNedopusk[1], //число.предрейс
                            dopuskNedopusk[0], //кол-во допусков
                            dopuskNedopusk[1], //кол-во недопусков
                            0)//ноль т.к. лист предрейса)
            );
        }

        // проходим по послерейсу (совмещенному с линейным) и добавляем, если такой даты нет, обновляем, если такая дата есть
        for (Map.Entry<Integer, Integer[]> entry: posl.entrySet()) {
            FactTable temp = new FactTable(); // времянка для доставаемых значений
            int vsegoMO = 0;
            int predrVsego = 0;
            int dopuskov = 0;
            int nedopuskov = 0;
            int poslerVsego = 0;

            Integer key = entry.getKey();                //получаем дату
            Integer[] dopuskNedopusk = entry.getValue(); //значения допуска/недопуска в эту дату

            //если дата уже внесена - достаем значения, добавляем новые и перезаписываем значения новыми
            //иначе добавляем дату и значения на эту дату
            if(res.containsKey(key)){ //если дата имеется
                temp = res.get(key);  // достаем значения (объект FactTable)
                //берем значения из полученного FactTable
                vsegoMO = temp.obscheeChisloMO;
                predrVsego = temp.kolichPredreisMO;
                dopuskov = temp.kolichDopuskov;
                nedopuskov = temp.kolichNedopuskov;
                poslerVsego = temp.kolichPoslerMO;
                //перезаписываем значения
                vsegoMO = vsegoMO+dopuskNedopusk[0]+dopuskNedopusk[1];
                dopuskov = dopuskov + dopuskNedopusk[0];
                nedopuskov = nedopuskov + dopuskNedopusk[1];
                poslerVsego = poslerVsego + dopuskNedopusk[0]+dopuskNedopusk[1];
                //обновляем объект FactTable
                res.put(key, new FactTable(vsegoMO, predrVsego, dopuskov, nedopuskov, poslerVsego));
            }
            else { //такой даты нет в общей таблице
                res.put(key,
                        new FactTable(
                                dopuskNedopusk[0]+dopuskNedopusk[1], //общ.число.МО
                                0, ////ноль т.к. лист послерейса)
                                dopuskNedopusk[0], //кол-во допусков
                                dopuskNedopusk[1], //кол-во недопусков
                                dopuskNedopusk[0]+dopuskNedopusk[1])
                );

            }
        }
        return res;
    }

    private TreeMap<String, int[]> prepareTable2 (List<ArrayList<String>> spisokVesPred,
                                                  List<ArrayList<String>> spisokVesPosl,
                                                  ArrayList<Integer> alldates,
                                                  int stolbec) { // stolbec это столбец откуда берем данные
        //заготовка для результата
        TreeMap<String, int[]> result = new TreeMap<>();
        int vsegoDat = alldates.size();

        // foreach для предрейса, потом для послерейса
        for (ArrayList<String> stroka : spisokVesPred) { //пробегаемся по строкам
            int[] calendDates = new int[vsegoDat]; //готовим таблицу дат осмотров для каждой фамилии
            String fio = stroka.get(stolbec); // получаем ФИО из 6й ячейки строки -> было "получаем ФИО из 5й ячейки строки"
            Integer data = getDate(stroka.get(1)); // получаем дату из второй ячейки строки
            int dataPosition = getDataPosition(alldates, data);
            int counter = 1; // значение предрейса = 1, т.к. ФИО взята из списка предрейса в эту дату
            //int poslereis = getIntFromFloatString(stroka.get(4)); // значение послерейса (0 или 1)
            if ((result.get(fio)==null))       // если эта фамилия еще не внесена
            {
                //По позиции даты в календаре alldates определяем номер ячейки, в которую пишем значение счетчика
                calendDates[dataPosition] = counter;
                //добавляем фамилию (ключ) и начальные счетчики его осмотров по датам
                result.put(fio, calendDates);
            } else {       // если эта фамилия уже внесена
                // получаем значения ячеек согласно календаря
                calendDates = result.get(fio);
                //По позиции даты в календаре alldates определяем номер ячейки, в которую добавляем сумму предрейса и послерейса
                calendDates[dataPosition] = calendDates[dataPosition] + counter;
                //calendDates .add(Integer.valueOf(stroka.get(1)), (predreis+poslereis));
                //добавляем фамилию (ключ) и новые счетчики его осмотров по дате
                result.put(fio, calendDates);
            }

        }
        // повторяем foreach и для послерейса
        for (ArrayList<String> stroka : spisokVesPosl) { //пробегаемся по строкам
            int[] calendDates = new int[vsegoDat]; //готовим таблицу дат осмотров для каждой фамилии
            String fio = stroka.get(stolbec); // получаем ФИО из 6й ячейки строки -> было "получаем ФИО из 5й ячейки строки"
            Integer data = getDate(stroka.get(1)); // получаем дату из второй ячейки строки
            int dataPosition = getDataPosition(alldates, data);
            int counter = 1; // значение предрейса = 1, т.к. ФИО взята из списка предрейса в эту дату
            //int poslereis = getIntFromFloatString(stroka.get(4)); // значение послерейса (0 или 1)
            if ((result.get(fio)==null))       // если эта фамилия еще не внесена
            {
                //По позиции даты в календаре alldates определяем номер ячейки, в которую пишем значение счетчика
                calendDates[dataPosition] = counter;
                //добавляем фамилию (ключ) и начальные счетчики его осмотров по датам
                result.put(fio, calendDates);
            } else {       // если эта фамилия уже внесена
                // получаем значения ячеек согласно календаря
                calendDates = result.get(fio);
                //По позиции даты в календаре alldates определяем номер ячейки, в которую добавляем сумму предрейса и послерейса
                calendDates[dataPosition] = calendDates[dataPosition] + counter;
                //calendDates .add(Integer.valueOf(stroka.get(1)), (predreis+poslereis));
                //добавляем фамилию (ключ) и новые счетчики его осмотров по дате
                result.put(fio, calendDates);
            }

        }
        return result;
    }

    private TreeMap<String,DriverRiskData> prepareTableGruppyRiska(List<ArrayList<String>> pred, List<ArrayList<String>> posle) throws Exception {
        //заготовка для результата
        TreeMap<String, DriverRiskData> result = new TreeMap<>();

        String vidNedopuska = "";
        String FIO = "";
        String denRogdeniya = "";
        DriverRiskData temp;

        List<ArrayList<String>> listVseMO = new ArrayList<>(); // для объединения предрейса и послерейса
        //объединяем списки
        listVseMO.addAll(pred);
        listVseMO.addAll(posle);

        //формируем список по ФИО водителей
        for (ArrayList<String> zapis: listVseMO) {
            //получаем водителя и пустую заготовку данных
            FIO = zapis.get(6);
            denRogdeniya = zapis.get(8);
            temp = new DriverRiskData(denRogdeniya, 0, 0, 0, new ArrayList<Integer>(), new ArrayList<Integer>(), new ArrayList<Integer>());

            //если водителя нет в списке - добавляем все данные из записи
            //если водитель есть в списке - обновляем данные
            if ((result.get(FIO)==null))       // если эта фамилия еще не внесена
            {
                //получаем данные
                String daNet = zapis.get(16); // допущен / не допущен
                String[] bloodPressure = zapis.get(11).trim().split("/"); //[0]-САД [1]-ДАД
                //temp.setDataRojdeniya(zapis.get(8)); //дата рожд.
                temp.setOsmotrovVsego(1); //начальное значение общего числа осмотров по данному сотруднику

               // if (daNet.equals("Не допущен")){
                    vidNedopuska = zapis.get(17).trim();
                    if (vidNedopuska.contains("АД")|(vidNedopuska.contains("ЧСС")))  { //недопуск по мед.причинам
                        temp.setNedopuskov(1);    //начальное значение числа недопусков
                        temp.setDopuskov(0);      //начальное значение числа допусков
                    }
                /*}*/ else { //допущен
                    temp.setNedopuskov(0);    //начальное значение числа недопусков
                    temp.setDopuskov(1);      //начальное значение числа допусков
                }

                // Незавершенный осмотр не учитывается
                if (!vidNedopuska.contains("Незавершенный осмотр")){
                    temp.srednSAD.add(Integer.parseInt(bloodPressure[0]));
                    temp.srednDAD.add(Integer.parseInt(bloodPressure[1]));
                    temp.srednCHSS.add(Integer.parseInt(zapis.get(12).trim()));
                }

                //добавляем фамилию (ключ) и начальные счетчики его осмотра
                result.put(FIO, temp);
            } else {       // если эта фамилия уже внесена
                // получаем значения для обновления
                temp = result.get(FIO);
                //получаем данные
                String daNet = zapis.get(16); // допущен / не допущен
                String[] bloodPressure = zapis.get(11).trim().split("/"); //[0]-САД [1]-ДАД
                temp.setOsmotrovVsego(temp.getOsmotrovVsego()+1); //обновляем значение общего числа осмотров по данному сотруднику

               // if (daNet.equals("Не допущен")){
                    vidNedopuska = zapis.get(17).trim();
                    if (vidNedopuska.contains("АД")|(vidNedopuska.contains("ЧСС")))  { //недопуск по мед.причинам
                        temp.setNedopuskov(temp.getNedopuskov()+1);    //увеличиваем значение числа недопусков
                    }
               /* }*/ else { //допущен
                    temp.setDopuskov(temp.getDopuskov()+1);      //увеличиваем значение числа допусков
                }

                // Незавершенный осмотр не учитывается
                if (!vidNedopuska.contains("Незавершенный осмотр")){
                    temp.srednSAD.add(Integer.parseInt(bloodPressure[0]));
                    temp.srednDAD.add(Integer.parseInt(bloodPressure[1]));
                    temp.srednCHSS.add(Integer.parseInt(zapis.get(12).trim()));
                }


                //добавляем фамилию (ключ) и новые счетчики его осмотрa
                result.put(FIO, temp);
            }
        }
        return result;
    }

    private int getDataPosition(ArrayList<Integer> alldates, Integer data) {
        int res = -1;
        for (int i = 0; i < alldates.size(); i++) {
            if (alldates.get(i)==data){
                res = i;
            }
        }
        return res;
    }

    private String makeWordDocumentTable1XLS (TreeMap<Integer, FactTable> preparedList, String uploadFilePath) throws IOException, XmlException {

        String res = File.separator+organization+" (фактич.) ["+period.toLowerCase()+"] "
                + makeFileNameByDateAndTimeCreated()+".docx";

        //For writing the Document in file system
        FileOutputStream out = new FileOutputStream(new File(uploadFilePath
                + res));

        //Blank Document
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
        // создаем верхний колонтитул Word файла
        CTP ctpHeaderModel = createHeaderModel(copyright+" "+DEV_NAME+"   "+ arrow + "  "+ DEV_LINK);
        // устанавливаем сформированный верхний
        // колонтитул в модель документа Word
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
        headerFooterPolicy.createHeader(
                XWPFHeaderFooterPolicy.DEFAULT,
                new XWPFParagraph[]{headerParagraph}
        );

        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Set alignment paragraph to CENTER
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run.setFontFamily("Times New Roman");
        run.setFontSize(14);
        run.setText("Отчет по "+organization);                  run.addCarriageReturn();
        run.setText("за фактически проведенные предрейсовые и");run.addCarriageReturn();
        run.setText("послерейсовые медицинские осмотры");       run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" "+god+" года"); //todo: год тоже надо вытаскивать из эксель +
        run.addCarriageReturn(); //возможно убрать пустую строку

        //create table
        XWPFTable table = document.createTable();
        table.setCellMargins(10,50,10,50);

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);

        tableRowOne.getCell(0).setParagraph(fillParagraph(document, "№ п/п"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(1).setParagraph(fillParagraph(document, "Число отчетного месяца"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(2).setParagraph(fillParagraph(document, "Общее кол-во мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(fillParagraph(document, "Кол-во предр. мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(4).setParagraph(fillParagraph(document, "Кол-во мед.осмотров \"Допуск\""));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(5).setParagraph(fillParagraph(document, "Кол-во мед.осмотров \"Недопуск\""));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(6).setParagraph(fillParagraph(document, "Кол-во послер. мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(7).setParagraph(fillParagraph(document, "% отстранения"));

        //table.getRow(0).getCell(0).addParagraph();

        Iterator iterator = preparedList.keySet().iterator();
        int count = 0;          //счетчик строк таблицы
        int countMedOsm = 0;    //счетчик мед.осмотров общий
        int countPredr = 0;     //счетчик предрейсовых МО
        int countDopusk = 0;    //счетчик допусков
        int countNoDopusk = 0;  //счетчик не допусков
        int countPoslereis = 0; //счетчик послерейс.мед.осмотров

        while(iterator.hasNext()) {
            count++;
            Integer key   =(Integer) iterator.next();
            FactTable value = preparedList.get(key);

            countMedOsm = countMedOsm + value.obscheeChisloMO;
            countPredr = countPredr + value.kolichPredreisMO;  //всего предр. осмотров за этот день
            countDopusk = countDopusk + value.kolichDopuskov;
            countNoDopusk = countNoDopusk + value.kolichNedopuskov;
            countPoslereis = countPoslereis + value.kolichPoslerMO;

            //create next rows
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(fillParagraph(document, Integer.toString(count)));
            tableRowNext.getCell(1).setParagraph(fillParagraph(document, Integer.toString(key)));      //день месяца
            tableRowNext.getCell(2).setParagraph(fillParagraph(document, Integer.toString(value.obscheeChisloMO)));    //всего мед.осм.
            tableRowNext.getCell(3).setParagraph(fillParagraph(document, Integer.toString(value.kolichPredreisMO))); //предрейс.
            tableRowNext.getCell(4).setParagraph(fillParagraph(document, Integer.toString(value.kolichDopuskov))); //допущ.
            tableRowNext.getCell(5).setParagraph(fillParagraph(document, Integer.toString(value.kolichNedopuskov))); //не допущ.
            tableRowNext.getCell(6).setParagraph(fillParagraph(document, Integer.toString(value.kolichPoslerMO))); //послерейс.
            tableRowNext.getCell(7).setParagraph(fillParagraph(document, String.format("%.2f",(value.procentNedopuska*100)))); //%.недопуска
        }

        //добавляем последнюю строку с итоговыми счетчиками
        XWPFTableRow tableRowLast = table.createRow();
        //tableRowLast.getCell(0).setParagraph(paragraph);
        tableRowLast.getCell(0).setParagraph(fillParagraph(document,"")); //№ п/п
        //tableRowLast.getCell(1).setParagraph(paragraph);
        tableRowLast.getCell(1).setParagraph(fillParagraph(document,"Итого:"));   //день месяца
        //tableRowLast.getCell(2).setParagraph(paragraph);
        tableRowLast.getCell(2).setParagraph(fillParagraph(document, Integer.toString(countMedOsm)));   //всего мед.осм.
        //tableRowLast.getCell(3).setParagraph(paragraph);
        tableRowLast.getCell(3).setParagraph(fillParagraph(document, Integer.toString(countPredr)));   //предрейс.
        //tableRowLast.getCell(4).setParagraph(paragraph);
        tableRowLast.getCell(4).setParagraph(fillParagraph(document, Integer.toString(countDopusk)));   //допущ.
        tableRowLast.getCell(5).setParagraph(fillParagraph(document, Integer.toString(countNoDopusk))); //не допущ.
        tableRowLast.getCell(6).setParagraph(fillParagraph(document, Integer.toString(countPoslereis))); //послерейс.
        tableRowLast.getCell(7).setParagraph(fillParagraph(document, String.format("%.2f", (countNoDopusk/(float)countMedOsm)*100))); //%невыпуска итоговый

        //List<XWPFParagraph> allParagraphs = document.getParagraphs();

        //костыль: удаляем весь мусор после таблицы - т.е. оставляем только первые два элемента документа (параграф и таблица)
        List<IBodyElement> elements = document.getBodyElements();
        for ( int i = elements.size()-1; i >= 2; i-- ) {
            //System.out.println( "removing bodyElement #" + i );
            document.removeBodyElement( i );
        }

        document.write(out); //сохраняем файл отчета в Word
        out.close();
        document.close();

        return res;
    }

    private String makeWordDocumentTable2XLS(String vidOtcheta,
                                             ArrayList<Integer> alldates,
                                             TreeMap<String, int[]> medOsmotryByFIOXLS,
                                             String uploadFilePath) throws IOException, XmlException, InterruptedException {
        String type = "";

        String res = File.separator+organization+" ("+vidOtcheta+") ["+period.toLowerCase()+"] "
                + makeFileNameByDateAndTimeCreated()+".docx";
        if (vidOtcheta.equals("точки осм.")) type = "Точка осмотра"; else type = "ФИО";

        //For writing the Document in file system
        FileOutputStream out = new FileOutputStream(new File(uploadFilePath + res));

        //Blank Document
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
        // создаем верхний колонтитул Word файла
        CTP ctpHeaderModel = createHeaderModel(copyright+" "+DEV_NAME+"   "+ arrow + "  "+ DEV_LINK);
        // устанавливаем сформированный верхний
        // колонтитул в модель документа Word
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
        headerFooterPolicy.createHeader(
                XWPFHeaderFooterPolicy.DEFAULT,
                new XWPFParagraph[]{headerParagraph}
        );

        //установка альбомной ориентации
        CTBody body = document.getDocument().getBody();
        if (!body.isSetSectPr()) {
            body.addNewSectPr();
        }
        CTSectPr section = body.getSectPr();

        if(!section.isSetPgSz()) {
            section.addNewPgSz();
        }
        CTPageSz pageSize = section.getPgSz();

        //для ландшафтной бумаги типа LETTER
        //pageSize.setW(BigInteger.valueOf(15840));
        //pageSize.setH(BigInteger.valueOf(12240));
        // --> https://overcoder.net/q/1168121/как-установить-ориентацию-страницы-для-документа-word
        pageSize.setW(BigInteger.valueOf(16840));
        pageSize.setH(BigInteger.valueOf(11900));

        pageSize.setOrient(STPageOrientation.LANDSCAPE);
        //ориентация страницы установлена

        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Set alignment paragraph to CENTER
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run.setFontFamily("Times New Roman");
        run.setFontSize(14);
        run.setText("Детализация"); run.addCarriageReturn();
        run.setText("предрейсовых(предсменных)/послерейсовых(послесменных)"); run.addCarriageReturn();
        run.setText("медицинских осмотров автоматизированным способом"); run.addCarriageReturn();
        run.setText(organization);run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" "+god+" года");run.addCarriageReturn();
        //run.addCarriageReturn(); //возможно убрать пустую строку

        //create table
        XWPFTable table = document.createTable();
        table.setCellMargins(10,50,10,50);
        table.setTableAlignment(CENTER);

        //**************************************************
        //https://stackoverflow.com/questions/27209863/apache-poi-merge-cells-from-a-table-in-a-word-document
        //**************************************************

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);

        tableRowOne.getCell(0).setParagraph(fillParagraphBold(document, "№ п/п"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(1).setParagraph(fillParagraphBold(document, type+ " \\ День месяца"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(2).setParagraph(fillParagraphBold(document, "∑"));
        tableRowOne.getCell(2).setWidth("400"); //устанавливаем ширину 20

        //выводим даты в шапку таблицы
        int vsegoDat = alldates.size();
        for (int i = 0; i < vsegoDat; i++) {
            tableRowOne.addNewTableCell();
            tableRowOne.getCell(i+3).setWidth("400"); //устанавливаем ширину 20
            tableRowOne.getCell(i+3).setParagraph(fillParagraphBold(document, alldates.get(i).toString()));
        }
        //шапка готова, заполняем таблицу
        int count = 0;   //счетчик для номеров п/п (строк)
        int countMO = 0; //Общий счетчик мед.осм (всех)
        int[] countMODaily = new int[vsegoDat]; //счетчик мед.осм. по каждой дате
        for (String fio : medOsmotryByFIOXLS.keySet()
                ) {
            count++;
            int[] temp =  medOsmotryByFIOXLS.get(fio);  //получаем массив осмотров по датам у водителя(фамилии)
            int driversAllMO = countDriversAllMO(temp);//получаем кол-во осмотров по водителю за месяц
            countMO = countMO+driversAllMO;           //подсчет общего кол-ва МО
            //create next rows
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(fillParagraph(document, Integer.toString(count)));       // № п/п
            tableRowNext.getCell(1).setParagraph(fillParagraphLeft(document, fio));                       // ФИО водителя
            tableRowNext.getCell(2).setParagraph(fillParagraph(document, Integer.toString(driversAllMO)));// ∑ по водителю
            //заполняем ячейки в каждой дате
            for (int j = 0; j < vsegoDat; j++) {
                tableRowNext.getCell(3+j).setParagraph(fillParagraph(document, Integer.toString(temp[j])));  //кол-во мед.осм. в эту дату
                countMODaily[j] = countMODaily[j]+temp[j]; //подсчет кол-ва МО в каждом календарном дне
            }
        }

        //добавляем последнюю строку с итоговыми счетчиками
        XWPFTableRow tableRowLast = table.createRow();
        tableRowLast.getCell(0).setParagraph(fillParagraph(document,""));                  //№ п/п
        tableRowLast.getCell(1).setParagraph(fillParagraphRight(document,"Всего:"));       //под ФИО
        tableRowLast.getCell(2).setParagraph(fillParagraph(document,Integer.toString(countMO)));//общее кол-во МО
        //заполняем итоговые ячейки по каждой дате
        for (int j = 0; j < vsegoDat; j++) {
            tableRowLast.getCell(3+j).setParagraph(fillParagraph(document, Integer.toString(countMODaily[j])));  //кол-во мед.осм. в эту дату
        }

        //костыль: удаляем весь мусор после таблицы - т.е. оставляем только первые два элемента документа (параграф и таблица)
        List<IBodyElement> elements = document.getBodyElements();
        for ( int i = elements.size()-1; i >= 2; i-- ) {
            //System.out.println( "removing bodyElement #" + i );
            document.removeBodyElement( i );
        }

        document.write(out); //сохраняем файл отчета в Word
        document.close();
        out.close();

        return res;
    }

    private String makeWordDocumentGruppaRiska(TreeMap<String, DriverRiskData> spisok,
                                               String uploadFilePath) throws IOException {
        TreeMap<Float, DriverRiskData> riskGroup = new TreeMap<>();

        String res = File.separator + organization + " (гр.риска) [" + period.toLowerCase() + "] "
                + makeFileNameByDateAndTimeCreated() + ".docx";

        //For writing the Document in file system
        FileOutputStream out = new FileOutputStream(new File(uploadFilePath
                + res));

        //отбираем в группу риска сотрудников с тремя и более осмотрами + недопусками от 20% и сортируем по %недопуска
        for (String s: spisok.keySet()) {
            spisok.get(s).setProcentNedopuskov(); //считаем % недопуски
            spisok.get(s).setFIO(s);    //устанавливаем фамилию - дубляж :/
            if ((spisok.get(s).getOsmotrovVsego()>=3)&(spisok.get(s).getProcentNedopuskov()>=0.25)){
                riskGroup.put(spisok.get(s).getProcentNedopuskov(), spisok.get(s));
            }
        }

        //Blank Document
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
        // создаем верхний колонтитул Word файла
        CTP ctpHeaderModel = createHeaderModel(copyright+" "+DEV_NAME+"   "+ arrow + "  "+ DEV_LINK);
        // устанавливаем сформированный верхний
        // колонтитул в модель документа Word
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
        headerFooterPolicy.createHeader(
                XWPFHeaderFooterPolicy.DEFAULT,
                new XWPFParagraph[]{headerParagraph}
        );

        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Set alignment paragraph to CENTER
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        //run.setBold(true);
        run.setText("Группы риска по артериальному давлению и пульсу");   run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" "+god+" года");

        if (riskGroup.isEmpty()){
            run.addCarriageReturn();
            run.addCarriageReturn();
            run.setText("Группы риска не сформированы, т.к. отсутствуют сотрудники, имеющие от 25% недопусков на основании не менее трёх осмотров");
        } else {
            //подготовка форматирования ячеек
            XWPFParagraph paragraphTableCell = document.createParagraph();
//            XWPFRun runCalibri = paragraphTableCell.createRun();
//            runCalibri.setFontFamily("Calibri");
//            runCalibri.setFontSize(9);
            paragraphTableCell.setAlignment(ParagraphAlignment.CENTER);
            paragraphTableCell.setSpacingAfter(0);
            paragraphTableCell.setSpacingBetween(1.00);
            //paragraphTableCell.setStyle();

            XWPFParagraph paragraphTableCellL = document.createParagraph();
            paragraphTableCellL.setAlignment(ParagraphAlignment.LEFT);
            paragraphTableCellL.setSpacingAfter(0);
            paragraphTableCellL.setSpacingBetween(1.00);
            //paragraphTableCellL.createRun().setFontSize(9);
            //cellrun.setFontFamily("Calibri");
            //cellrun.setFontSize(9);

            //create table
            XWPFTable table = document.createTable();
            table.setCellMargins(10,50,10,50);
            table.setTableAlignment(TableRowAlign.valueOf("CENTER"));

            //create first row
            XWPFTableRow tableRowOne = table.getRow(0);

            tableRowOne.getCell(0).setParagraph(paragraphTableCell);
            tableRowOne.getCell(0).setText("№ п/п");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(1).setParagraph(paragraphTableCell);
            tableRowOne.getCell(1).setText("ФИО сотрудника");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(2).setParagraph(paragraphTableCell);
            tableRowOne.getCell(2).setText("Дата рождения");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(3).setParagraph(paragraphTableCell);
            tableRowOne.getCell(3).setText("Возраст, полных лет");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(4).setParagraph(paragraphTableCell);
            tableRowOne.getCell(4).setText("Осмотров всего");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(5).setParagraph(paragraphTableCell);
            tableRowOne.getCell(5).setText("Мед. показатели: норма");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(6).setParagraph(paragraphTableCell);
            tableRowOne.getCell(6).setText("Мед. показатели: вне нормы");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(7).setParagraph(paragraphTableCell);
            tableRowOne.getCell(7).setText("% не допусков");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(8).setParagraph(paragraphTableCell);
            tableRowOne.getCell(8).setText("Ср. знач. САД");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(9).setParagraph(paragraphTableCell);
            tableRowOne.getCell(9).setText("Ср. знач. ДАД");

            tableRowOne.addNewTableCell();
            tableRowOne.getCell(10).setParagraph(paragraphTableCell);
            tableRowOne.getCell(10).setText("Ср. знач. ЧСС");

            //добавляем остальные строки (сортированы по %недопуска)
            int i = 0;
            for (Float fl:riskGroup.keySet()) {
                XWPFTableRow tableRowNext = table.createRow();
                tableRowNext.getCell(0).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(0).setText(Integer.toString(++i));         // № п/п
                tableRowNext.getCell(1).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(1).setText(riskGroup.get(fl).getFIO());    //ФИО
                tableRowNext.getCell(2).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(2).setText(riskGroup.get(fl).getDataRojdeniya());         // Дата рождения
                tableRowNext.getCell(3).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(3).setText(Long.toString(riskGroup.get(fl).getVozrast())); // Возраст
                tableRowNext.getCell(4).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(4).setText(Integer.toString(riskGroup.get(fl).getOsmotrovVsego()));   // Осмотров всего
                tableRowNext.getCell(5).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(5).setText(Integer.toString(riskGroup.get(fl).getDopuskov()));   // Допусков
                tableRowNext.getCell(6).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(6).setText(Integer.toString(riskGroup.get(fl).getNedopuskov()));   // Недопусков
                tableRowNext.getCell(7).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(7).setText(String.format("%.2f", fl*100));                      // % недопусков
                tableRowNext.getCell(8).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(8).setText(Integer.toString(riskGroup.get(fl).setSrednSAD()));   // Ср.САД
                tableRowNext.getCell(9).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(9).setText(Integer.toString(riskGroup.get(fl).setSrednDAD()));   // Ср.ДАД
                tableRowNext.getCell(10).setParagraph(paragraphTableCellL);
                tableRowNext.getCell(10).setText(Integer.toString(riskGroup.get(fl).setSrednCHSS())); // Ср.ЧСС
            }
        }
        document.write(out); //сохраняем файл отчета в Word
        out.close();
        document.close();
        return res;
    }

    private String makeWordDocumentStatNedopuskov(List<ArrayList<String>> pred,
                                          List<ArrayList<String>> posle,
                                          String uploadFilePath) throws IOException {

        String res = File.separator + organization + " (недопуски) [" + period.toLowerCase() + "] "
                + makeFileNameByDateAndTimeCreated() + ".docx";

        //For writing the Document in file system
        FileOutputStream out = new FileOutputStream(new File(uploadFilePath
                + res));

        //Blank Document
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
        // создаем верхний колонтитул Word файла
        CTP ctpHeaderModel = createHeaderModel(copyright+" "+DEV_NAME+"   "+ arrow + "  "+ DEV_LINK);
        // устанавливаем сформированный верхний
        // колонтитул в модель документа Word
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
        headerFooterPolicy.createHeader(
                XWPFHeaderFooterPolicy.DEFAULT,
                new XWPFParagraph[]{headerParagraph}
        );

        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Set alignment paragraph to CENTER
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        //run.setBold(true);
        run.setText("Статистика причин отстранения");   run.addCarriageReturn();
        run.setText(organization);                  run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" "+god+" года");            //run.addCarriageReturn();

        //подготовка форматирования ячеек
        XWPFParagraph paragraphTableCell = document.createParagraph();
        paragraphTableCell.setAlignment(ParagraphAlignment.CENTER);
        paragraphTableCell.setSpacingAfter(0);
        paragraphTableCell.setSpacingBetween(1.00);

        XWPFParagraph paragraphTableCellL = document.createParagraph();
        paragraphTableCellL.setAlignment(ParagraphAlignment.LEFT);
        paragraphTableCellL.setSpacingAfter(0);
        paragraphTableCellL.setSpacingBetween(1.00);

        List<ArrayList<String>> listVseMO = new ArrayList<>(); // для объединения
        //объединяем списки
        listVseMO.addAll(pred);
        listVseMO.addAll(posle);

        Integer[] itog = makeStatNedopuskov(document, listVseMO, paragraphTableCell, paragraphTableCellL); //формирование таблицы в документе ворд
        //itog[] = всего осм, кол-во недопусков, в т.ч. по мед.причинам
        //добавляем итоговые записи вида:
        //Всего недопусков:           79 (21,5% от всех осмотров)
        //в т.ч. по мед.причинам:     58 (15,8% от всех осмотров)
        XWPFParagraph paragraphText = document.createParagraph();
        paragraphText.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun runText = paragraphText.createRun();
        runText.setFontFamily("Times New Roman");
        runText.setFontSize(11);
        runText.setText("Всего не допусков: "+itog[1]+" ("+String.format("%.1f", (itog[1]/(float)itog[0])*100)+"% от всех осмотров)");   runText.addCarriageReturn();
        runText.setText("в т.ч. по мед.причинам: "+itog[2]+" ("+String.format("%.1f", (itog[2]/(float)itog[0])*100)+"% от всех осмотров)");   //run.addCarriageReturn();

        document.write(out); //сохраняем файл отчета в Word
        out.close();
        document.close();
        return res;
    }

    private Integer[] makeStatNedopuskov(XWPFDocument wordDoc, List<ArrayList<String>> listVseMO, XWPFParagraph par1, XWPFParagraph par2) {
        int vsegoMO = listVseMO.size(); //общее число медосмотров
        int countNedopuskiMO = 0; //счетчик для недопусков по мед.причинам
        int chisloNedopuskov = 0; //суммарное число недопусков
        String vidNedopuska; //вид недопуска
        String daNet;
        TreeMap<String, Integer> vseNedopuskiStat = new TreeMap<String, Integer>(); //счетчик недопусков
        //здесь
        //  key - вид недопуска (String)
        //  value - кол-во недопусков по данному виду
        //проходимся по списку и считаем недопуски по каждому виду
        for (ArrayList<String> zapis: listVseMO) {
            switch (radiobutton[0]){
                case "4":
                    daNet = zapis.get(17);
                    if (!daNet.equals("") ){ //если ячейка с комментарием недопуска не пуста, т.е. есть описание причины недопуска
                        vidNedopuska = zapis.get(17).trim();
                        if (vidNedopuska.contains("АД")|
                                (vidNedopuska.contains("ЧСС"))|
                                (vidNedopuska.contains("жалоб"))|
                                (vidNedopuska.contains("паров"))|
                                (vidNedopuska.contains("Температура"))|
                                (vidNedopuska.contains("Цвет"))|
                                (vidNedopuska.contains("Травма"))) countNedopuskiMO++; //подсчет недопусков по мед.причинам
                        if(vseNedopuskiStat.containsKey(vidNedopuska)){ //такой недопуск есть
                            int count = vseNedopuskiStat.get(vidNedopuska); // получаем число недопусков
                            count++;                                    // увеличиваем счетчик
                            vseNedopuskiStat.put(vidNedopuska, count);  // обновляем инфу
                        }
                        else {
                            vseNedopuskiStat.put(vidNedopuska, 1);      // добавляем первый недопуск данного вида
                        }
                    }
                    break;
                default:
                    daNet = zapis.get(16).toLowerCase();
                    if (daNet.equals("не допущен") |
                            daNet.equals("не прошел") |
                            daNet.equals("не прошёл") |
                            (daNet.equals("прошел") & zapis.get(15).equals("Н")) ){
                        vidNedopuska = zapis.get(17).trim();
                        if (vidNedopuska.contains("АД")|
                                (vidNedopuska.contains("ЧСС"))|
                                (vidNedopuska.contains("жалоб"))|
                                (vidNedopuska.contains("паров"))|
                                (vidNedopuska.contains("Температура"))|
                                (vidNedopuska.contains("Цвет"))|
                                (vidNedopuska.contains("Травма"))) countNedopuskiMO++; //подсчет недопусков по мед.причинам
                        if(vseNedopuskiStat.containsKey(vidNedopuska)){ //такой недопуск есть
                            int count = vseNedopuskiStat.get(vidNedopuska); // получаем число недопусков
                            count++;                                    // увеличиваем счетчик
                            vseNedopuskiStat.put(vidNedopuska, count);  // обновляем инфу
                        }
                        else {
                            vseNedopuskiStat.put(vidNedopuska, 1);      // добавляем первый недопуск данного вида
                        }
                    }
                    break;
            }

        }

        //считаем недопуски все
        for (Integer k: vseNedopuskiStat.values()) {
            chisloNedopuskov = chisloNedopuskov + k;
        }

        //create table
        XWPFTable table = wordDoc.createTable();
        table.setCellMargins(10,50,10,50);
        table.setTableAlignment(TableRowAlign.valueOf("CENTER"));

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);

        tableRowOne.getCell(0).setParagraph(par1);
        tableRowOne.getCell(0).setText("№ п/п");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(1).setParagraph(par2);
        tableRowOne.getCell(1).setText("Комментарий") /*.setParagraph(fillParagraphBold(document, "ФИО сотрудника"))*/;

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(2).setParagraph(par1);
        tableRowOne.getCell(2).setText("Количество не допусков");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(par1);
        tableRowOne.getCell(3).setText("% от всех не допусков");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(4).setParagraph(par1);
        tableRowOne.getCell(4).setText("% от всех осмотров");

        //добавляем остальные строки (начальные даты месяца в конце списка)
        int i = 0;
        for (String st:vseNedopuskiStat.keySet()) {
            int num = vseNedopuskiStat.get(st); //число недопусков по данному виду недопусков
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(par1);
            tableRowNext.getCell(0).setText(Integer.toString(++i)); // № п/п
            tableRowNext.getCell(1).setParagraph(par2);
            tableRowNext.getCell(1).setText(st);                       // Комментарий (вид недопуска)
            tableRowNext.getCell(2).setParagraph(par1);
            tableRowNext.getCell(2).setText(Integer.toString(num));    // Количество недопусков данного вида
            tableRowNext.getCell(3).setParagraph(par1);
            tableRowNext.getCell(3).setText(String.format("%.1f", (num/(float)chisloNedopuskov)*100));    // % от всех недопусков
            tableRowNext.getCell(4).setParagraph(par1);
            tableRowNext.getCell(4).setText(String.format("%.2f", (num/(float)vsegoMO)*100));             // % от всех осмотров
        }
        //возвращаем данные для последних итоговых строк
        Integer[] res = new Integer[3];
        res[0] = vsegoMO;
        res[1] = chisloNedopuskov;
        res[2] = countNedopuskiMO;
        return res;
    }

    private String makeWordDocumentReestr(List<ArrayList<String>> pred,
                                          List<ArrayList<String>> posle,
                                          List<ArrayList<String>> line,
                                          List<ArrayList<String>> prof,
                                          String uploadFilePath) throws IOException, XmlException {

        String res = File.separator + organization + " (реестр) [" + period.toLowerCase() + "] "
                + makeFileNameByDateAndTimeCreated() + ".docx";

        //For writing the Document in file system
        FileOutputStream out = new FileOutputStream(new File(uploadFilePath
                + res));

        //считаем кол-во осмотров, допусков, недопусков и сколько водителей осмотрели + общий процент недопуска
        int vsegoOsm = countOsm(pred, posle, line);
        int dopuskov = 0;
        switch (radiobutton[0]){
            case "4":
                dopuskov = countDopusk_M(pred, posle, line);
                break;
            default:
                dopuskov = countDopusk(pred, posle, line);
                break;
        }
        int nedopuskov = /*countNedopusk(pred, posle, line);*/ vsegoOsm-dopuskov;
        float procentNedopuskov = nedopuskov/(float)vsegoOsm;
        int chisloVoditelei = countVod(pred, posle, line); //ОК
        float srednVozrast = summaVosrastov()/(float)chisloVoditelei;
        allVozrasts.clear(); //очистка списка всех возрастов для следующего запуска процедуры (для обработки очередного меджурнала)
        int before = 0;
        int after = 0;
        int regular = 0;
        int profControl = 0;
        if (!pred.isEmpty()) before=pred.size();
        if (!posle.isEmpty()) after=posle.size();
        if (!line.isEmpty()) regular=line.size();
        if (!prof.isEmpty()) profControl=prof.size();
        //int tablesCounter = 0; //счетчик номеров таблиц
        int [] tablesCounter = new int [1]; //счетчик номеров таблиц
        String fraza1 = "Всего осмотров: "+vsegoOsm+", в т.ч. предрейсовых – "+before;
        String fraza2 = "Допусков, всего – "+dopuskov+", не допусков – "+nedopuskov+", что составило "+String.format("%.1f", procentNedopuskov*100)+"% от общего числа медосмотров.";
        String fraza3 = "Всего осмотрено сотрудников: "+chisloVoditelei+" чел., средний возраст по группе: "+String.format("%.1f",srednVozrast)+" лет.";
        String dobavka ="";
        String fraza4 = ""; //фраза профконтроля

        if (after>0) dobavka = dobavka+", послерейсовых – "+after;
        if (regular>0) dobavka = dobavka+", линейных – "+regular;
        dobavka = dobavka+".";
        fraza1 = fraza1+dobavka;
        // посчитали и подготовили текст (три фразы в три строки)

        //готовим фразу профконтроля
        if (profControl != 0) {
            int chisloSotr = countSotr(prof);
            //  Профконтроль: алкотестирование - 12 шт. (5 чел.)
            fraza4 = "Профконтроль: алкотестирование - "+profControl+" шт. ("+chisloSotr+" чел.)";
        }

        //Blank Document
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
        // создаем верхний колонтитул Word файла
        CTP ctpHeaderModel = createHeaderModel(copyright+" "+DEV_NAME+"   "+ arrow + "  "+ DEV_LINK);
        // устанавливаем сформированный верхний
        // колонтитул в модель документа Word
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
        headerFooterPolicy.createHeader(
                XWPFHeaderFooterPolicy.DEFAULT,
                new XWPFParagraph[]{headerParagraph}
        );

        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Set alignment paragraph to CENTER
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run.setFontFamily("Times New Roman");
        run.setFontSize(12);
        run.setBold(true);
        run.setText("Отчет по медицинским осмотрам сотрудников");   run.addCarriageReturn();
        run.setText(organization);   run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" "+god+" года");            //run.addCarriageReturn();

        XWPFParagraph paragraphText = document.createParagraph();
        paragraphText.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun runText = paragraphText.createRun();
        runText.setFontFamily("Times New Roman");
        runText.setFontSize(12);
        //runText.addCarriageReturn(); //возможно убрать пустую строку
        runText.setText(fraza1); runText.addCarriageReturn();
        runText.setText(fraza2); runText.addCarriageReturn();
        runText.setText(fraza3);
        if (profControl != 0){
            runText.addCarriageReturn();
            runText.setText(fraza4);
        }
        //
        //  до табличных реестров выводим надпись вида
        //  Всего осмотров: 652, в т.ч. предрейсовых – 636, послерейсовых – 16.
        //  Допусков, всего – 547, не допусков – 105, что составило 16,1% от общего числа медосмотров.
        //  Всего осмотрено водителей: 209 чел.
        //  Профконтроль: алкотестирование - 12 шт. (5 чел.)

        //подготовка форматирования ячеек
        XWPFParagraph paragraphTableCell = document.createParagraph();
        paragraphTableCell.setAlignment(ParagraphAlignment.CENTER);
        paragraphTableCell.setSpacingAfter(0);
        paragraphTableCell.setSpacingBetween(1.00);

        XWPFParagraph paragraphTableCellL = document.createParagraph();
        paragraphTableCellL.setAlignment(ParagraphAlignment.LEFT);
        paragraphTableCellL.setSpacingAfter(0);
        paragraphTableCellL.setSpacingBetween(1.00);


        if (!pred.isEmpty()) makeReestr(document, "предрейсовых", pred, tablesCounter, paragraphTableCell, paragraphTableCellL);
        if (!posle.isEmpty()) makeReestr(document, "послерейсовых", posle, tablesCounter, paragraphTableCell, paragraphTableCellL);
        if (!line.isEmpty()) makeReestr(document, "линейных", line, tablesCounter, paragraphTableCell, paragraphTableCellL);
        if (!prof.isEmpty()) makeProfcontrolTable(document, "Профконтроль: алкотестирование", prof, tablesCounter, paragraphTableCell, paragraphTableCellL);

        document.write(out); //сохраняем файл отчета в Word
        out.close();
        document.close();
        return res;
    }

    private void makeProfcontrolTable (XWPFDocument wordDoc,
                             String vid,
                             List<ArrayList<String>> spisok,
                             int[] num,
                             XWPFParagraph par1,
                             XWPFParagraph par2){
        int size = spisok.size();
        int n = num[0];
        XWPFParagraph paragraphTableNum = wordDoc.createParagraph();
        paragraphTableNum.setAlignment(ParagraphAlignment.RIGHT);
        paragraphTableNum.setSpacingAfter(0);
        paragraphTableNum.setSpacingBetween(1.00);
        XWPFRun runTableNum = paragraphTableNum.createRun();
        runTableNum.setFontFamily("Times New Roman");
        runTableNum.setFontSize(12);
        runTableNum.addCarriageReturn();
        runTableNum.setText("Табл. "+(++n)); //runTableNum.addCarriageReturn();
        // Табл. 1
        //обновляем значение номера таблицы для следующего реестра
        num[0] = n;
        XWPFParagraph paragraphTableName = wordDoc.createParagraph();
        paragraphTableName.setAlignment(ParagraphAlignment.CENTER);
        paragraphTableName.setSpacingAfter(0);
        paragraphTableName.setSpacingBetween(1.00);
        XWPFRun runTableName = paragraphTableName.createRun();
        runTableName.setFontFamily("Times New Roman");
        runTableName.setFontSize(12);
        runTableName.setText(vid); // "Профконтроль: алкотестирование"

        //create table
        XWPFTable table = wordDoc.createTable();
        table.setCellMargins(10,50,10,50);
        table.setTableAlignment(TableRowAlign.valueOf("CENTER"));

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);

        tableRowOne.getCell(0).setParagraph(par1);
        tableRowOne.getCell(0).setText("№ п/п");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(1).setParagraph(par1);
        tableRowOne.getCell(1).setText("Дата и время");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(2).setParagraph(par1);
        tableRowOne.getCell(2).setText("ФИО сотрудника");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(par1);
        tableRowOne.getCell(3).setText("Алкоголь");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(4).setParagraph(par1);
        tableRowOne.getCell(4).setText("Результат");

        //добавляем остальные строки (начальные даты месяца в конце списка)
        for (int i = size; i>0; i--){
            //String[] timestamp = spisok.get(i-1).get(1).split(" ");
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(par1);
            tableRowNext.getCell(0).setText(Integer.toString(size-i+1)); // № п/п
            tableRowNext.getCell(1).setParagraph(par2);
            tableRowNext.getCell(1).setText(spisok.get(i-1).get(2));   // Дата и время осмотра
            tableRowNext.getCell(2).setParagraph(par2);
            tableRowNext.getCell(2).setText(spisok.get(i-1).get(7));   // ФИО сотрудника
            tableRowNext.getCell(3).setParagraph(par1);
            tableRowNext.getCell(3).setText(spisok.get(i-1).get(11));   // Алкоголь
            tableRowNext.getCell(4).setParagraph(par2);
            tableRowNext.getCell(4).setText(spisok.get(i-1).get(12));   // Результат
        }
    }

    private void makeReestr (XWPFDocument wordDoc,
                             String vid,
                             List<ArrayList<String>> spisok,
                             int[] num,
                             XWPFParagraph par1,
                             XWPFParagraph par2){
        int size = spisok.size();
        int n = num[0];
        XWPFParagraph paragraphTableNum = wordDoc.createParagraph();
        paragraphTableNum.setAlignment(ParagraphAlignment.RIGHT);
        paragraphTableNum.setSpacingAfter(0);
        paragraphTableNum.setSpacingBetween(1.00);
        XWPFRun runTableNum = paragraphTableNum.createRun();
        runTableNum.setFontFamily("Times New Roman");
        runTableNum.setFontSize(12);
        runTableNum.addCarriageReturn();
        runTableNum.setText("Табл. "+(++n)); //runTableNum.addCarriageReturn();
        // Табл. 1
        //обновляем значение номера таблицы для следующего реестра
        num[0] = n;
        XWPFParagraph paragraphTableName = wordDoc.createParagraph();
        paragraphTableName.setAlignment(ParagraphAlignment.CENTER);
        paragraphTableName.setSpacingAfter(0);
        paragraphTableName.setSpacingBetween(1.00);
        XWPFRun runTableName = paragraphTableName.createRun();
        runTableName.setFontFamily("Times New Roman");
        runTableName.setFontSize(12);
        runTableName.setText("Реестр "+vid+" медицинских осмотров."); //runTableName.addCarriageReturn();

        //create table
        XWPFTable table = wordDoc.createTable();
        table.setCellMargins(10,50,10,50);
        table.setTableAlignment(TableRowAlign.valueOf("CENTER"));

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);

        tableRowOne.getCell(0).setParagraph(par1);
        tableRowOne.getCell(0).setText("№ п/п");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(1).setParagraph(par1);
        tableRowOne.getCell(1).setText("ФИО сотрудника") /*.setParagraph(fillParagraphBold(document, "ФИО сотрудника"))*/;

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(2).setParagraph(par1);
        tableRowOne.getCell(2).setText("Дата осмотра");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(par1);
        tableRowOne.getCell(3).setText("Время осмотра");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(4).setParagraph(par1);
        tableRowOne.getCell(4).setText("Результат");

        //добавляем остальные строки (начальные даты месяца в конце списка)
        for (int i = size; i>0; i--){
            String[] timestamp = spisok.get(i-1).get(1).split(" ");
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(par1);
            tableRowNext.getCell(0).setText(Integer.toString(size-i+1)); // № п/п
            tableRowNext.getCell(1).setParagraph(par2);
            tableRowNext.getCell(1).setText(spisok.get(i-1).get(6));   // ФИО сотрудника
            tableRowNext.getCell(2).setParagraph(par1);
            tableRowNext.getCell(2).setText(timestamp[0]);   // Дата осмотра
            tableRowNext.getCell(3).setParagraph(par1);
            tableRowNext.getCell(3).setText(timestamp[1]);   // Время осмотра
            tableRowNext.getCell(4).setParagraph(par2);
            tableRowNext.getCell(4).setText(spisok.get(i-1).get(16));   // Результат
        }
    }

    private int countVod(List<ArrayList<String>> pred, List<ArrayList<String>> posle, List<ArrayList<String>> line) {
        int res;
        Set<String> Voditeli = new HashSet<>();
        if (!pred.isEmpty()){
            for (ArrayList<String> st0: pred) {
                boolean isAdded = Voditeli.add(st0.get(6));
                if (isAdded) {
                    int vozrast = countDriversAge(st0.get(8));
                    allVozrasts.add(vozrast);
                }
            }
        }
        if (!posle.isEmpty()){
            for (ArrayList<String> st1: posle) {
                boolean isAdded = Voditeli.add(st1.get(6));
                if (isAdded) {
                    int vozrast = countDriversAge(st1.get(8));
                    allVozrasts.add(vozrast);
                }
            }
        }
        if (!line.isEmpty()){
            for (ArrayList<String> st2: line) {
                boolean isAdded = Voditeli.add(st2.get(6));
                if (isAdded) {
                    int vozrast = countDriversAge(st2.get(8));
                    allVozrasts.add(vozrast);
                }
            }
        }
        res = Voditeli.size();
        return res;
    }

    private int countSotr(List<ArrayList<String>> prof){
        int res = 0;
        Set<String> sotrudniki = new HashSet<>();
        if (!prof.isEmpty()){
            for (ArrayList<String> st0: prof) {
                sotrudniki.add(st0.get(7)); //столбец с фамилией
            }
        }
        res = sotrudniki.size();
        return res;
    }

    private int summaVosrastov(){
        int res = 0;
        for (Integer vozrast : allVozrasts) {
            res = res + vozrast;
        }
        return res;
    };

    private int countOsm(List<ArrayList<String>> pred, List<ArrayList<String>> posle, List<ArrayList<String>> line) {
        int res = 0;
        if (!pred.isEmpty()){
            res = res + pred.size();
        }
        if (!posle.isEmpty()) {
            res = res + posle.size();
        }
        if (!line.isEmpty()) {
            res = res + line.size();
        }
        return res;
    }

    private int countDopusk(List<ArrayList<String>> pred, List<ArrayList<String>> posle, List<ArrayList<String>> line) {
        int res = 0;
        if (!pred.isEmpty()){
            for (ArrayList<String> st0: pred) {
                if (st0.get(16).equals("Допущен")){
                    res++;
                }
            }
        }
        if (!posle.isEmpty()) {
            for (ArrayList<String> st1 : posle) {
                if (st1.get(16).equals("Допущен") | st1.get(16).equals("Прошёл") | st1.get(16).equals("Прошел")) {
                    res++;
                }
            }
        }
        if (!line.isEmpty()) {
            for (ArrayList<String> st2 : line) {
                if (st2.get(16).equals("Допущен") | st2.get(16).equals("Прошёл") | st2.get(16).equals("Прошел")) {
                    res++;
                }
            }
        }
        return res;
    }

    private int countDopusk_M(List<ArrayList<String>> pred, List<ArrayList<String>> posle, List<ArrayList<String>> line) {
        int res = 0;
        if (!pred.isEmpty()){
            for (ArrayList<String> st0: pred) {
                if (st0.get(16).toLowerCase().contains("не допущен")){
                    //do nothing
                } else res++;
            }
        }
        if (!posle.isEmpty()) {
            for (ArrayList<String> st1 : posle) {
                if (st1.get(16).toLowerCase().contains("выявлены")){
                    //do nothing
                } else res++;
            }
        }
        if (!line.isEmpty()) {
            for (ArrayList<String> st2 : line) {
                if (st2.get(16).toLowerCase().contains("не допущен")) {
                    //do nothing
                } else res++;
            }
        }
        return res;
    }


    private int countDriversAge(String st) {
        int res = 0;
        if (!(st==null)){


                String[] dr = st.split("-"); // 31-08-2020  делим по дефису
                Integer y;
                Integer m;
                Integer d;
                LocalDate bday = null;
                if (dr.length==3) {
                    y = Integer.parseInt(dr[2]);
                    m = Integer.parseInt(dr[1]);
                    d = Integer.parseInt(dr[0]);
                    bday = LocalDate.of(y, m, d);
                }
                LocalDate today = LocalDate.now(); ///of(2010, 5, 17); //
                 //
                int vozrast = calculateAge(bday, today);
                res = res + vozrast;

        }

        return res;
    }

    //подсчет количества медосмотров у водителя
    int countDriversAllMO (int[] allDates){
        int res = 0;
        for (int c:allDates) {
            res = res+c;
        }
        return res;
    }

    // создаем хедер или верхний колонтитул
    private static CTP createHeaderModel(String headerContent) {

        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();

        cttHeader.setStringValue(headerContent);
        return ctpHeaderModel;
    }

    //create Paragraph For Cells (текст по центру)
    private XWPFParagraph fillParagraph(XWPFDocument document, String text) {
        XWPFParagraph paragraphForCells = document.createParagraph();
        paragraphForCells.setAlignment(ParagraphAlignment.CENTER);
        paragraphForCells.setSpacingAfter(0);
        paragraphForCells.setSpacingBefore(0);
        XWPFRun run = paragraphForCells.createRun();
        run.setFontSize(11);
        run.setFontFamily("Times New Roman");
        run.setText(text);
        return paragraphForCells;
    }

    //create Paragraph For Cells (полужирный для шапки)
    private XWPFParagraph fillParagraphBold(XWPFDocument document, String text) {
        XWPFParagraph paragraphForCells = document.createParagraph();
        paragraphForCells.setAlignment(ParagraphAlignment.CENTER);
        paragraphForCells.setSpacingAfter(0);
        paragraphForCells.setSpacingBefore(0);
        XWPFRun run = paragraphForCells.createRun();
        run.setBold(true);
        run.setFontSize(11);
        run.setFontFamily("Times New Roman");
        run.setText(text);
        return paragraphForCells;
    }

    //create Paragraph For Cells (Влево)
    private XWPFParagraph fillParagraphLeft(XWPFDocument document, String text) {
        XWPFParagraph paragraphForCells = document.createParagraph();
        paragraphForCells.setAlignment(ParagraphAlignment.LEFT);
        paragraphForCells.setSpacingAfter(0);
        paragraphForCells.setSpacingBefore(0);
        XWPFRun run = paragraphForCells.createRun();
        run.setFontSize(11);
        run.setFontFamily("Times New Roman");
        run.setText(text);
        return paragraphForCells;
    }

    //create Paragraph For Cells (Вправо)
    private XWPFParagraph fillParagraphRight(XWPFDocument document, String text) {
        XWPFParagraph paragraphForCells = document.createParagraph();
        paragraphForCells.setAlignment(ParagraphAlignment.RIGHT);
        paragraphForCells.setSpacingAfter(0);
        paragraphForCells.setSpacingBefore(0);
        XWPFRun run = paragraphForCells.createRun();
        run.setFontSize(11);
        run.setFontFamily("Times New Roman");
        run.setText(text);
        return paragraphForCells;
    }

    //подготовка названия файла
    private String makeFileNameByDateAndTimeCreated(){
        String dateTimeAdded = "";
        Calendar calendar = Calendar.getInstance(TimeZone.getDefault(), Locale.getDefault());
        calendar.setTime(new Date());
        //int year = calendar.get(Calendar.YEAR); //текущий год

        DateFormat formatter = new SimpleDateFormat("YYYY.MM.dd__HH-mm-ss");

        dateTimeAdded = formatter.format(calendar.getTime()); //время добавления документа
        return dateTimeAdded;
    }






    //получение из первой строки Excel название компании
    private String getOrganizationName (ArrayList<String> firsRow){
        String res = "orgName";
        String row = firsRow.get(0);

        //разбиваем строку по пробелам
//        String[] tempArray = row.split(" ");
//        //собираем название организации (без начальных и без последних четырех элементов временного массива)
//        for (int i=5; i<tempArray.length-4; i++){
//            res = res+" "+tempArray[i];
//        } //устаревший метод, не применяется

        int r = Integer.parseInt(radiobutton[0]);
        if (r == 4) { //если выбран меджурнал
            int end = row.indexOf(" с ");
            int begin = row.lastIndexOf("организации");
            res = row.substring(begin+11, end); //Название компании
        }
        if (r == 3) { //если выбран универсальный отчет
            int end = row.indexOf('(');
            res = row.substring(0, end); //Название компании c самого начала и до первой скобки
        }
        if (r == 2) { //если журнал старого образца
            int begin = row.lastIndexOf("осмотра");
            if (begin!=(-1)){
                String temp1 = row.substring(begin);
                int end = temp1.indexOf('(');
                if (end!=(-1)){
                    res = temp1.substring(8, end); //Название компании после слова "осмотра" и до первой скобки
                    if (res.length()>100){
                        Transliterator toLatinTrans = Transliterator.getInstance(CYRILLIC_TO_LATIN);
                        String result = toLatinTrans.transliterate(res);
                        res = result;
                    }
                }
            }
        }

        //меняем в названии кавычки вида  " на такие: « в начале и в конце »
        StringBuilder strBuilder = new StringBuilder(res);
        int marker = 0;
        for (int i = 0; i < res.length(); i++) {
            if (res.charAt(i)=='\"'){

                if (marker==0) {
                    strBuilder.replace(i, i+1, "«");
                    marker++;
                } else {
                    strBuilder.replace(i, i+1, "»");
                    i = res.length();
                }

            }
        }
        res = strBuilder.toString();

        return res.trim();
    }

    //получение из первой строки Excel название компании
    private String getOrganizationName_v2 (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);

        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        //собираем название организации (без начальных и без последних четырех элементов временного массива)
        for (int i=5; i<tempArray.length-6; i++){
            res = res+" "+tempArray[i];
        }

        //меняем в названии кавычки вида  " на такие: « в начале и в конце »
        StringBuilder strBuilder = new StringBuilder(res);
        int marker = 0;
        for (int i = 0; i < res.length(); i++) {
            if (res.charAt(i)=='\"'){

                if (marker==0) {
                    strBuilder.replace(i, i+1, "«");
                    marker++;
                } else {
                    strBuilder.replace(i, i+1, "»");
                    i = res.length();
                }

            }
        }
        res = strBuilder.toString();
        return res.trim();
    }

    //получение из первой строки Excel отчетного месяца
    private String getMonth_v3 (ArrayList<String> firsRow){
        String res = "month";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        //String[] tempArray = row.split(" ");
        String tempStr = "";
        Locale rLocale = new Locale("ru"); //русская локаль
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy", rLocale);
        SimpleDateFormat newFormatter = new SimpleDateFormat("MMMM", rLocale);

        int dot = row.indexOf('.');
        if (dot!=(-1)){
            tempStr = row.substring(dot-2, dot+8); //от первой точки два символа слева и 8 справа
        }

        try {
            //Date date = formatter.parse(tempArray[tempArray.length-1]);
            Date date = formatter.parse(tempStr);
            res = newFormatter.format(date);

        } catch (ParseException e) {
            e.printStackTrace();
            return res.trim();
        }
        return res.trim();
    }

    //получение из первой строки Excel отчетного месяца
    private String getMonth_v2 (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        Locale rLocale = new Locale("ru"); //русская локаль
        SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy", rLocale);
        SimpleDateFormat newFormatter = new SimpleDateFormat("MMMM", rLocale);

        try {
            Date date = formatter.parse(tempArray[tempArray.length-3]);
            res = newFormatter.format(date);

        } catch (ParseException e) {
            e.printStackTrace();
        }
        return res.trim();
    }

    //получение из первой строки Excel года
    private String getGod_v3 (ArrayList<String> firsRow){
        String res = "god";
        String row = firsRow.get(0);

        //разбиваем строку по пробелам - устаревший метод - не используется
        //String[] tempArray = row.split(" ");
        //String temp = tempArray[tempArray.length-1];
        //String[] tempos = temp.split("\\.");
        //res = tempos[2];
        ////////////////////////////////////////////////////////////////////

        int dot = row.indexOf('.');
        if (dot!=(-1)){
            res = row.substring(dot+4, dot+8); //от первой точки пляшем
        }

        return res.trim();
    }

    //получение из первой строки Excel года
    private String getGod_v2 (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        String temp = tempArray[tempArray.length-3];
        String[] tempos = temp.split("-");
        res = tempos[2];
        return res.trim();
    }

    //получение списка файлов с отчетами
    static List<String> getFileTree(String root) throws IOException {
        List<String> fileTree = new ArrayList<>();
        File path = new File(root);
        Queue<File> directories = new LinkedList<>(); //папки перебираются в очереди
        directories.add(path);
        try {
            while (directories.size()!=0){
                File[] fList = ((LinkedList<File>) directories).getFirst().listFiles();
                for (File file : fList) {
                    if (file.isFile()) {            // если файл, то добавляем в список файлов
                        fileTree.add(file.getAbsolutePath());

                    } else {
                        if (file.isDirectory()) {  // если это папка, то добавляем в конец очереди (список папок)
                            ((LinkedList<File>) directories).addLast(file);
                        }
                    }
                }
                ((LinkedList<File>) directories).removeFirst(); //прошлись по всей очереди, удаляем первого
            }
        } catch (NullPointerException e) {
            e.printStackTrace();
            fileTree.add("empty");
        }

        return fileTree;
    }

    //возвращает возраст человека
    public int calculateAge(
            LocalDate birthDate,
            LocalDate currentDate) {
        // validate inputs ...
        if ((birthDate != null) && (currentDate != null)) {
            return Period.between(birthDate, currentDate).getYears();
        } else {
            return 0;
        }
    }

    private class DriverRiskData {
        String dataRojdeniya;
        String fio;
        long vozrast;
        int osmotrovVsego;
        int dopuskov;
        int nedopuskov;
        float procentNedopuskov;
        ArrayList<Integer> srednSAD;
        ArrayList<Integer> srednDAD;
        ArrayList<Integer> srednCHSS;

        public void setDataRojdeniya(String dataRojdeniya) {
            this.dataRojdeniya = dataRojdeniya;
        }

        public void setOsmotrovVsego(int osmotrovVsego) {
            this.osmotrovVsego = osmotrovVsego;
        }

        public void setDopuskov(int dopuskov) {
            this.dopuskov = dopuskov;
        }

        public void setNedopuskov(int nedopuskov) {
            this.nedopuskov = nedopuskov;
        }

        //конструктор по умолчанию
        public DriverRiskData() {}

        //конструктор
        public DriverRiskData(String dataRojdeniya,
                              int osmotrovVsego,
                              int dopuskov,
                              int nedopuskov,
                              //float procentNedopuskov,
                              ArrayList<Integer> srednSAD,
                              ArrayList<Integer> srednDAD,
                              ArrayList<Integer> srednCHSS) {
            this.dataRojdeniya = dataRojdeniya;
            this.vozrast = countVozrast(dataRojdeniya);
            this.osmotrovVsego = osmotrovVsego;
            this.dopuskov = dopuskov;
            this.nedopuskov = nedopuskov;
            //this.procentNedopuskov = procentNedopuskov;
            this.srednSAD = srednSAD;
            this.srednDAD = srednDAD;
            this.srednCHSS = srednCHSS;
        }

        void setProcentNedopuskov() {
            this.procentNedopuskov = this.nedopuskov / (float)this.osmotrovVsego;
        }

        Integer setSrednSAD(){
            int sum = 0;
            for (int davlenie : this.srednSAD) {
                sum = sum + davlenie;
            }
            return Math.round(sum/this.srednSAD.size());
        }

        Integer setSrednDAD(){
            int sum = 0;
            for (int davlenie : this.srednDAD) {
                sum = sum + davlenie;
            }
            return Math.round(sum/this.srednDAD.size());
        }

        Integer setSrednCHSS(){
            int sum = 0;
            for (int davlenie : this.srednCHSS) {
                sum = sum + davlenie;
            }
            return Math.round(sum/this.srednCHSS.size());
        }

        Long countVozrast(String birthDay){
            String[] dr = birthDay.split("-"); // 31-08-2020  делим по дефису
            Integer y = Integer.parseInt(dr[2]);
            Integer m = Integer.parseInt(dr[1]);
            Integer d = Integer.parseInt(dr[0]);

            LocalDate today = LocalDate.now(); ///of(2010, 5, 17); //
            LocalDate bday = LocalDate.of(y, m, d); //
            return ChronoUnit.YEARS.between(bday, today);
        }



        public void setFIO (String s){ this.fio=s; }

        public String getDataRojdeniya() {
            return dataRojdeniya;
        }

        public String getFIO() {
            return fio;
        }

        public long getVozrast() {
            return vozrast;
        }

        public int getOsmotrovVsego() {
            return osmotrovVsego;
        }

        public int getDopuskov() {
            return dopuskov;
        }

        public int getNedopuskov() {
            return nedopuskov;
        }

        public float getProcentNedopuskov() {
            return procentNedopuskov;
        }

    }
}

