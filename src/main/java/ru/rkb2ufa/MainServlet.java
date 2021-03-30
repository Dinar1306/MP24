package ru.rkb2ufa;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;   //для xls
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;   //для xlsx
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.http.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.*;
import static org.apache.poi.xwpf.usermodel.TableRowAlign.CENTER;

@MultipartConfig //запрос может содержать несколько параметров
        (fileSizeThreshold=1024*1024*5, // 5MB
         maxFileSize=1024*1024*10,      // 10MB
         maxRequestSize=1024*1024*50)   // 50MB

public class MainServlet extends HttpServlet {

    static final String REPORTS_DIR = "otchety";
    private static List<String> filesList = new ArrayList<>();
    private List<ReportsTable> spisokOtchetov_v2 = new ArrayList<>();     // список отчетов из списка файлов в папке отчетов
    private String organization = "";
    private String period = "";
    private String god = "";
    private boolean failed = false;
    private int errorStringNumber;
    private String debug = "";
    private String message = "";
    private int chisloMO = 0; //общее число медосмотров из трех списков(предр, послер и линейный)
    private int chisloPredr = 0; //общее число предр.медосмотров
    private int chisloPosler = 0; //общее число послер.медосмотров
    private int chisloLine = 0; //общее число линейн.медосмотров

    private class FactTable {
        int obscheeChisloMO;
        int kolichPredreisMO;
        int kolichDopuskov;
        int kolichNedopuskov;
        int kolichPoslerMO;
        float procentNedopuska;

        public void setProcentNedopuska() {
            this.procentNedopuska = this.kolichNedopuskov / (float)this.obscheeChisloMO;
        }

        //конструктор
        public FactTable(int obscheeChisloMO, int kolichPredreisMO, int kolichDopuskov, int kolichNedopuskov, int kolichPoslerMO) {
            this.obscheeChisloMO = obscheeChisloMO;
            this.kolichPredreisMO = kolichPredreisMO;
            this.kolichDopuskov = kolichDopuskov;
            this.kolichNedopuskov = kolichNedopuskov;
            this.kolichPoslerMO = kolichPoslerMO;
            //this.procentNedopuska = kolichNedopuskov/(float)obscheeChisloMO;
        }

        //конструктор по умолчанию
        public FactTable() {
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

                //Готовим таблицу из списка
                // Назв.орг. | Тип.отч | Период(месяц) |  Дата/время создания | Скачать | Удалить
                //spisokOtchetov = makeTableFromFilelist(filesList);
                spisokOtchetov_v2 = makeTableFromFilelist_v2(filesList);
                response.setContentType("text/html");
                request.setCharacterEncoding ("UTF-8");
                response.setCharacterEncoding("UTF-8");
                request.setAttribute("spisokOtchetov_v2", spisokOtchetov_v2);
                requestDispatcher = request.getRequestDispatcher("list.jsp");
                requestDispatcher.forward(request, response);
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
                    requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                    requestDispatcher.forward(request, response);
                    return;
                    }
                response.setContentType("text/html");
                request.setCharacterEncoding ("UTF-8");
                response.setCharacterEncoding("UTF-8");
                request.setAttribute("spisokOtchetov_v2", spisokOtchetov_v2);
                requestDispatcher = request.getRequestDispatcher("list.jsp");
                requestDispatcher.forward(request, response);
                //String[] delString = allRows.get(id);
                break;
            default:
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
        //InputStream inputStream;                  // поток чтения для загружаемого файла
        XSSFWorkbook workBookXLSX;                  // объект книги эксель xlsx
        HSSFWorkbook myExcelBookXLS = null;         // объект книги эксель xls
        //String[] stroka = new String[20];         // строка таблицы с листа
        //String[] customs = null;
        List<ArrayList<String>> list = new ArrayList<>();     // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listPosleReis = new ArrayList<>(); // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listLine = new ArrayList<>(); // массив строк листа (каждая строка - массив строк) для medpont24
        List<ArrayList<String>> listPosleAndLine = new ArrayList<>(); // для объединения послерейса и линейного
        List<ArrayList<String>> listP = new ArrayList<>();    // массив строк листа (каждая строка - массив строк) для поликлиники
        TreeMap<Integer, Integer[]> medOsmotryByDatesPredReis = new TreeMap<Integer, Integer[]>(); //итоговые данные отсортированы по дате
        //т.е. здесть Integer Key - дата мед.осм.
        //Integer[] Value - таблица допущено / не допущено (в эту дату)
        TreeMap<Integer, Integer[]> medOsmotryByDatesPosleReis;
        TreeMap<Integer, FactTable> medOsmotryByDatesFacticheskie = new TreeMap<Integer, FactTable>();; // Таблица 1 для ворда
        //т.е. здесть Integer Key - дата мед.осм.
        //FactTable Value - таблица: общ.число.МО|кол.предр.МО|допусков|недопусков|кол.послер.МО|%невыпуска (в эту дату)

        TreeMap<Integer, int[]> medOsmotryByDatesALL = new TreeMap<Integer, int[]>();
        //т.е. здесть Integer Key - дата мед.осм.
        //int[] Value - таблица: общ.кол|предр|допущ|недоущ|послер| (в эту дату) --> добавить столбец %невыпуска

        // %невыпуска (для добавления последнего столбца при формировании документа ворд
        TreeMap<Integer, Float> medOsmotryByDatesAllProcent = new TreeMap<Integer, Float>();

        //итоговые данные отсортированы по дате (Поликлиника)
        TreeMap<Integer, Integer[]> medOsmotryByDatesXLS;
        //т.е. здесть Integer Key - дата мед.осм.
        //Integer[] Value - таблица предрейс / послерейс (в эту дату)

        //Массив дат медосмотров (для Табл.№2)
        ArrayList<Integer> dates = new ArrayList<>();

        //итоговые данные отсортированы по фамилиям и дате
        TreeMap<String, int[]> medOsmotryByFIOXLS;
        TreeMap<String, int[]> medOsmotryByFIO;
        TreeMap<String, int[]> medRabotnikByFIO;
        TreeMap<String, int[]> medOsmByHost;
        TreeMap<String, int[]> medOsmByNepoduski;
        // здесь key   это ФИО водителя - String
        // здесь value это таблица с суммарным значением предрейса и послерейса в каждой ячейке,
        // причем длина массива равна длине массива дат dates

        //инициализия завершена

        ///////////////WORK///////////////////
        //получаем части (нужные нам файлы)
        Part part = request.getPart("file");
        long size = part.getSize(); //файл медпойнта

        //Part part_p = request.getPart("file_p");
        //long size_p = part_p.getSize(); // файл поликлиники

        //проверям загруженли файл меджурнала:
        //ничего
        if (size == 0){
            request.setAttribute("message", "Не загружен файл меджурнала :(");
            RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
            requestDispatcher.forward(request, response);
            return;
        }
        // меджурнал medpoint24 загружен
        else {
            //получаем объект книги XLSX из формы
            workBookXLSX = XLSXFromPart(part);
            //разбираем первый лист файла medpoint24 на объектную модель
            list = getListFromSheet(workBookXLSX, 0); //получаем лист предрейса
            listPosleReis = getListFromSheet(workBookXLSX, 1); //получаем лист послерейса
            listLine = getListFromSheet(workBookXLSX, 5); //получаем лист линейного
            ArrayList<String> pervayaStroka = list.get(0); //первая строка (заголовок)

            try {
                organization = getOrganizationName_v2(pervayaStroka); //достаем из первой строки (заголовка) название компании.
                period = getMonth_v2(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
                god = getGod_v2(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.
            } catch (Exception e) {
                e.printStackTrace();
            }

            //Причесываем списки:
            // убираем заголовок таблицы, убираем шапку таблицы, убирем последние 5 и 7 ненужных строк из предрейса и послерейса соответственно
            list = list.subList(2, list.size()-5);
            listPosleReis = listPosleReis.subList(2, listPosleReis.size()-7);
            //причесываем линейный
            listLine = listLine.subList(2, listLine.size()-5);

            //считаем общее число медосмотров
            if (!list.isEmpty()) {
                chisloPredr = list.size();
                chisloMO = chisloMO + list.size();
            }
            if (!listPosleReis.isEmpty()){
                chisloPosler = listPosleReis.size();
                chisloMO = chisloMO + listPosleReis.size();
            }
            if (!listLine.isEmpty()){
                chisloLine = listLine.size();
                chisloMO = chisloMO + listLine.size();
            }

            ////соединяем послерейс и линейные МО
            //Объединение двух списков в третий:
            //result.addAll(list1);
            //result.addAll(list2);
            listPosleAndLine.addAll(listPosleReis);
            listPosleAndLine.addAll(listLine);

            //производим подсчёт по предрейсовым
            medOsmotryByDatesPredReis = prepare(list);

            //производим подсчёт по послерейсовым (старый вариант)
            //medOsmotryByDatesPosleReis = prepare(listPosleReis); //старый вариант
            //производим подсчёт по объединенному послерейсу и линейному (новый вариант)
            medOsmotryByDatesPosleReis = prepare(listPosleAndLine); //новый вариант

            //производим подсчёт по линейным
            //medOsmotryByDatesLine = prepare(listLine);



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
            medOsmotryByFIO = prepareTable2(list, listPosleAndLine, dates, 6);

            // (Табл.3 Детализация, по медсестрам) предрейс+послерейс, нужен 18й столбец
            medRabotnikByFIO = prepareTable2(list, listPosleAndLine, dates, 18);

            // (Табл.4 Детализация, по точкам) предрейс+послерейс, нужен 4й столбец
            medOsmByHost = prepareTable2(list, listPosleAndLine, dates, 4);

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
                table5FileName = makeWordDocumentReestr(list, listPosleReis, listLine, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table6FileName)
                table6FileName = makeWordDocumentStatNedopuskov(list, listPosleAndLine, uploadFilePath);

            } catch (XmlException e) {
                e.printStackTrace();
                //response.setContentType("text/html");
            } catch (InterruptedException e) {
                e.printStackTrace();
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
            request.setAttribute("reportsDir", REPORTS_DIR);
            request.setAttribute("message", "Отчёты по меджурналу Medpoint24 сформированы успешно!");
            RequestDispatcher requestDispatcher = request.getRequestDispatcher("otchet.jsp");
            requestDispatcher.forward(request, response);
            return;
        }

    }


    ////////////////////////////////////////////////////////////////////////
    //                      ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ                        //
    ////////////////////////////////////////////////////////////////////////


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

    //получаем объект книги xls
    private HSSFWorkbook XLSFromPart(Part part){
        InputStream inputStream;
        HSSFWorkbook workBook = new HSSFWorkbook();
        try {
            inputStream = part.getInputStream();
            workBook = new HSSFWorkbook(inputStream);
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

    //получаем лист из книги xls
    private List<ArrayList<String>> getListFromSheetXLS (HSSFWorkbook workBook/*, int num*/) throws IOException,
                                                                                                    NullPointerException,
                                                                                                    IllegalStateException{
        List<ArrayList<String>> res = new ArrayList<>();
        //получаем лист "Лист1"
        HSSFSheet myExcelSheet = workBook.getSheet("Лист1");
        int vsegoStrok = myExcelSheet.getPhysicalNumberOfRows()-1;  //-7;

        for (int i=0; i<=vsegoStrok; i++){
            HSSFRow row = myExcelSheet.getRow(i);
            if (row==null) break;
            short vsegoYacheek = row.getLastCellNum();
            ArrayList<String> tempStringArray = new ArrayList<>();
            //System.out.println(i);
            for (int j=0; j<vsegoYacheek; j++){
                try {
                if ((j==1)&(i>2)) {
                        Date date = row.getCell(j).getDateCellValue();
                        Calendar calendar = Calendar.getInstance(TimeZone.getDefault(), Locale.getDefault());
                        calendar.setTime(date);
                        int day = calendar.get(Calendar.DATE); //получаем дату
                        tempStringArray.add(String.valueOf(day));
                } else {
                    if(row.getCell(j).getCellType() == STRING/*HSSFCell.CELL_TYPE_STRING*/){
                        tempStringArray.add(row.getCell(j).getStringCellValue());
                    }

                    if(row.getCell(j).getCellType() == NUMERIC/*HSSFCell.CELL_TYPE_NUMERIC*/){
                        tempStringArray.add(Double.toString(row.getCell(j).getNumericCellValue()));
                    }

                    if(row.getCell(j).getCellType() == FORMULA/*HSSFCell.CELL_TYPE_NUMERIC*/){
                        tempStringArray.add(Double.toString(row.getCell(j).getNumericCellValue()));
                    }
                }
                //сюда катч
                } catch (NullPointerException e) {
                    errorStringNumber = i+2;
                    workBook.close();
                    throw new NullPointerException();
                } catch (IllegalStateException e) {
                    workBook.close();
                    //errorStringNumber = getIntFromFloatString(tempStringArray.get(0));
                    errorStringNumber = i-2;
                    //System.out.println("j="+j+", i="+i);
                    throw new IllegalStateException();
                }
            }
            res.add(tempStringArray);
        }
        workBook.close();

        return res;
    }

    private Integer getDate (String data){
        String[] s1 = data.split(" "); // 2020-08-31 18:06 делим по пробелу
        String[] s2 = s1[0].split("-"); // 2020-08-31  забираем дату -> 31
        Integer res = Integer.parseInt(s2[2]);
        return res;
    }

    private TreeMap<Integer, Integer[]> prepare (List<ArrayList<String>> spisokVes){
        //заготовка для результата
        TreeMap<Integer, Integer[]> result = new TreeMap<Integer, Integer[]>();

        // foreach
        for (ArrayList<String> stroka : spisokVes) { //пробегаеся по строкам
            Integer data = getDate(stroka.get(1)); // получаем дату из второй ячейки строки

            //определяем Допущен или Не допущен и увеличиваем счетчик в соответствующей ячейке (первой или второй)
            switch (stroka.get(16)){ //было 15
                case "Допущен":
                    //нашлелся допуск -> увеличиваем значение в первой ячейке
                    if ((result.get(data)==null))       // если эта дата еще не внесена
                    {
                        result.put(data, new Integer[] {1, 0}); //добавляем текущую строку (ключ) и счетчик (первое нахождение)
                    } else {
                        Integer[] v = result.get(data); //получаем значение счетчика допущенных (нужна будет первая ячейка)
                        v[0]++;                         // и увеличиваем
                        result.put(data, v);            // перезаписываем счетчик
                    }
                    break;
                case "Не допущен":
                    //нашлелся Не допуск -> увеличиваем значение во второй ячейке
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
        return result;
    }

    private int getIntFromFloatString (String floatString){
        float f = Float.parseFloat(floatString);
        return (int) f; // int
    }

    private TreeMap<Integer, Integer[]> prepareXLS (List<ArrayList<String>> spisokVes){
        //заготовка для результата
        TreeMap<Integer, Integer[]> result = new TreeMap<Integer, Integer[]>();

        // foreach
        for (ArrayList<String> stroka : spisokVes) { //пробегаемся по строкам
            //Integer data = getDate(stroka.get(1)); // получаем дату из второй ячейки строки
            Integer data = Integer.parseInt(stroka.get(1)); // получаем дату из второй ячейки строки

            // Предрейс или Послерейс -> увеличиваем счетчик в соответствующей ячейке (первой или второй)
            if ((result.get(data)==null))       // если эта дата еще не внесена
            {
                result.put(data, new Integer[] {0, 0, 0}); //добавляем текущую дату (ключ) и начальные счетчики:
                                                                                                        //предрейсов
                                                                                                        //послерейсов
                                                                                                        //не допущ. (всегда = 0 для поликл.)
                int predreis = getIntFromFloatString(stroka.get(3)); // значение предрейса (0 или 1)
                int poslereis = getIntFromFloatString(stroka.get(4)); // значение послерейса (0 или 1)
                if (predreis == 1){
                    Integer[] v = result.get(data); //получаем значение счетчиков
                    v[0]++;                         // и увеличиваем у предрейса
                    result.put(data, v);            // перезаписываем счетчик
                }
                if (poslereis == 1){
                    Integer[] v = result.get(data); //получаем значение счетчиков
                    v[1]++;                         // и увеличиваем у послерейса
                    result.put(data, v);            // перезаписываем счетчик
                }
            } else {       // если эта дата уже не внесена
                int predreis = getIntFromFloatString(stroka.get(3)); // значение предрейса (0 или 1)
                int poslereis = getIntFromFloatString(stroka.get(4)); // значение послерейса (0 или 1)
                if (predreis == 1){
                    Integer[] v = result.get(data); //получаем значение счетчиков
                    v[0]++;                         // и увеличиваем у предрейса
                    result.put(data, v);            // перезаписываем счетчик
                }
                if (poslereis == 1){
                    Integer[] v = result.get(data); //получаем значение счетчиков
                    v[1]++;                         // и увеличиваем у послерейса
                    result.put(data, v);            // перезаписываем счетчик
                }
            }
        }
        return result;
    }

    private TreeMap<String, int[]> prepareTable2XLS (List<ArrayList<String>> spisokVes, ArrayList<Integer> alldates) {
        //заготовка для результата
        TreeMap<String, int[]> result = new TreeMap<>();
        int vsegoDat = alldates.size();

        // foreach
        for (ArrayList<String> stroka : spisokVes) { //пробегаемся по строкам
            int[] calendDates = new int[vsegoDat]; //готовим таблицу дат осмотров для каждой фамилии
            String fio = stroka.get(2); // получаем ФИО из третьей ячейки строки
            Integer data = Integer.parseInt(stroka.get(1)); // получаем дату из второй ячейки строки
            int dataPosition = getDataPosition(alldates, data);
            int predreis = getIntFromFloatString(stroka.get(3)); // значение предрейса (0 или 1)
            int poslereis = getIntFromFloatString(stroka.get(4)); // значение послерейса (0 или 1)
            if ((result.get(fio)==null))       // если эта фамилия еще не внесена
            {
                //По позиции даты в календаре alldates определяем номер ячейки, в которую пишем сумму предрейса и послерейса
                calendDates[dataPosition] = (predreis+poslereis);
                //добавляем фамилию (ключ) и начальные счетчики его осмотров по датам
                result.put(fio, calendDates);
            } else {       // если эта фамилия уже внесена
                // получаем значения ячеек согласно календаря
                calendDates = result.get(fio);
                //По позиции даты в календаре alldates определяем номер ячейки, в которую добавляем сумму предрейса и послерейса
                calendDates[dataPosition] = calendDates[dataPosition] + (predreis+poslereis);
                //calendDates .add(Integer.valueOf(stroka.get(1)), (predreis+poslereis));
                //добавляем фамилию (ключ) и новые счетчики его осмотров по дате
                result.put(fio, calendDates);
            }

        }
        return result;
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

    private int getDataPosition(ArrayList<Integer> alldates, Integer data) {
        int res = -1;
        for (int i = 0; i < alldates.size(); i++) {
            if (alldates.get(i)==data){
                res = i;
            }
        }
        return res;
    }

    // TODO: +++ 09.09.2020 сделать возврат названия файла (чтобы передать в otchet.jsp для формирования ссылки для скачивания)
    private String makeWordDocumentTable1(TreeMap<Integer, int[]> preparedList, String uploadFilePath, TreeMap<Integer, Float> preparedProcent) throws IOException, XmlException {
        String copyright = "\u00a9";
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
        CTP ctpHeaderModel = createHeaderModel("Разработано "+copyright+"MDF-lab средствами Java");
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
        run.addCarriageReturn();

        //create table
        XWPFTable table = document.createTable();
        table.setCellMargins(10,50,10,50);

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);

        tableRowOne.getCell(0).setParagraph(fillParagraph(document, "№ п/п"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(1).setParagraph(fillParagraph(document, "Число отчетного месяца"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(2).setParagraph(fillParagraph(document, "Общее количество мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(fillParagraph(document, "Количество предрейсовых мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(4).setParagraph(fillParagraph(document, "Количество мед.осмотров \"Допуск\""));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(5).setParagraph(fillParagraph(document, "Количество мед.осмотров \"Недопуск\""));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(6).setParagraph(fillParagraph(document, "Количество послерейсовых мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(7).setParagraph(fillParagraph(document, "% невыпуска"));

        //table.getRow(0).getCell(0).addParagraph();

        Iterator iterator = preparedList.keySet().iterator();
        int count = 0;          //счетчик строк таблицы
        int countPredr = 0;     //счетчик предрейса
        int countDopusk = 0;    //счетчик допусков
        int countNoDopusk = 0;  //счетчик не допусков
        int countPosler = 0;    //счетчик послерейса
        int countMedOsm = 0;    //счетчик мед.осмотров
        while(iterator.hasNext()) {
            count++;
            Integer key   =(Integer) iterator.next();
            int[] value = preparedList.get(key);
            countMedOsm = countMedOsm + value[0];
            countPredr = countPredr + value[1];  //всего осмотров за этот день (допуск + не допуск)
            countDopusk = countDopusk + value[2];
            countNoDopusk = countNoDopusk + value[3];
            countPosler = countPosler + value[4];
            float procent = preparedProcent.get(key);

            //create next rows
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(fillParagraph(document, Integer.toString(count)));
            tableRowNext.getCell(1).setParagraph(fillParagraph(document, Integer.toString(key)));   //день месяца
            tableRowNext.getCell(2).setParagraph(fillParagraph(document, Integer.toString(value[0]))); //всего мед.осм.
            tableRowNext.getCell(3).setParagraph(fillParagraph(document, Integer.toString(value[1]))); //предрейсовых.
            tableRowNext.getCell(4).setParagraph(fillParagraph(document, Integer.toString(value[2]))); //допущ.
            tableRowNext.getCell(5).setParagraph(fillParagraph(document, Integer.toString(value[3]))); //не допущ.
            tableRowNext.getCell(6).setParagraph(fillParagraph(document, Integer.toString(value[4]))); //послерейсовых.
            tableRowNext.getCell(7).setParagraph(fillParagraph(document, String.format("%.2f", procent*100))); //%невыпуска
        }

        //добавляем последнюю строку с итоговыми счетчиками
        XWPFTableRow tableRowLast = table.createRow();
        tableRowLast.getCell(0).setParagraph(fillParagraph(document,"")); //№ п/п
        tableRowLast.getCell(1).setParagraph(fillParagraph(document,"Итого:"));   //день месяца
        tableRowLast.getCell(2).setParagraph(fillParagraph(document, Integer.toString(countMedOsm))); //всего мед.осм.
        tableRowLast.getCell(3).setParagraph(fillParagraph(document, Integer.toString(countPredr))); //предрейс.
        tableRowLast.getCell(4).setParagraph(fillParagraph(document, Integer.toString(countDopusk))); //допущ.
        tableRowLast.getCell(5).setParagraph(fillParagraph(document, Integer.toString(countNoDopusk))); //не допущ.
        tableRowLast.getCell(6).setParagraph(fillParagraph(document, Integer.toString(countPosler))); //послер.
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

    private String makeWordDocumentTable1XLS (TreeMap<Integer, FactTable> preparedList, String uploadFilePath) throws IOException, XmlException {
        String copyright = "\u00a9";
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
        CTP ctpHeaderModel = createHeaderModel("Разработано "+copyright+"MDF-lab средствами Java");
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
        tableRowOne.getCell(2).setParagraph(fillParagraph(document, "Общее количество мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(fillParagraph(document, "Количество предрейсовых мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(4).setParagraph(fillParagraph(document, "Количество мед.осмотров \"Допуск\""));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(5).setParagraph(fillParagraph(document, "Количество мед.осмотров \"Не допуск\""));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(6).setParagraph(fillParagraph(document, "Количество послерейсовых мед.осмотров"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(7).setParagraph(fillParagraph(document, "% невыпуска"));

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
        String copyright = "\u00a9";
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
        CTP ctpHeaderModel = createHeaderModel("Разработано "+copyright+"MDF-lab средствами Java");
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

    private String makeWordDocumentStatNedopuskov(List<ArrayList<String>> pred,
                                          List<ArrayList<String>> posle,
                                          String uploadFilePath) throws IOException {
        String copyright = "\u00a9";
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
        CTP ctpHeaderModel = createHeaderModel("Разработано " + copyright + "MDF-lab средствами Java");
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
        run.setText("Статистика причин недопусков");   run.addCarriageReturn();
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
        runText.setText("Всего недопусков: "+itog[1]+" ("+String.format("%.2f", (itog[1]/(float)itog[0])*100)+"% от всех осмотров)");   runText.addCarriageReturn();
        runText.setText("в т.ч. по мед.причинам: "+itog[2]+" ("+String.format("%.2f", (itog[2]/(float)itog[0])*100)+"% от всех осмотров)");   //run.addCarriageReturn();

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
        TreeMap<String, Integer> vseNedopuskiStat = new TreeMap<String, Integer>(); //счетчик недопусков
        //здесь
        //  key - вид недопуска (String)
        //  value - кол-во недопусков по данному виду
        //проходимся по списку и считаем недопуски по каждому виду
        for (ArrayList<String> zapis: listVseMO) {
            String daNet = zapis.get(16);
            if (daNet.equals("Не допущен")){
                vidNedopuska = zapis.get(17).trim();
                if (vidNedopuska.contains("АД")|(vidNedopuska.contains("ЧСС"))) countNedopuskiMO++; //подсчет недопусков по мед.причинам
                if(vseNedopuskiStat.containsKey(vidNedopuska)){ //такой недопуск есть
                    int count = vseNedopuskiStat.get(vidNedopuska); // получаем число недопусков
                    count++;                                    // увеличиваем счетчик
                    vseNedopuskiStat.put(vidNedopuska, count);  // обновляем инфу
                }
                else {
                    vseNedopuskiStat.put(vidNedopuska, 1);      // добавляем первый недопуск данного вида
                }
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
        tableRowOne.getCell(2).setText("Количество недопусков");

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(3).setParagraph(par1);
        tableRowOne.getCell(3).setText("% от всех недопусков");

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
            tableRowNext.getCell(3).setText(String.format("%.2f", (num/(float)chisloNedopuskov)*100));    // % от всех недопусков
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
                                          String uploadFilePath) throws IOException, XmlException {
        String copyright = "\u00a9";
        String res = File.separator + organization + " (реестр) [" + period.toLowerCase() + "] "
                + makeFileNameByDateAndTimeCreated() + ".docx";

        //For writing the Document in file system
        FileOutputStream out = new FileOutputStream(new File(uploadFilePath
                + res));

        //считаем кол-во осмотров, допусков, недопусков и сколько водителей осмотрели + общий процент недопуска
        int vsegoOsm = countOsm(pred, posle, line);
        int dopuskov = countDopusk(pred, posle, line);
        int nedopuskov = /*countNedopusk(pred, posle, line);*/ vsegoOsm-dopuskov;
        float procentNedopuskov = nedopuskov/(float)vsegoOsm;
        int chisloVoditelei = countVod(pred, posle, line); //ОК
        int before = 0;
        int after = 0;
        int regular = 0;
        if (!pred.isEmpty()) before=pred.size();
        if (!posle.isEmpty()) after=posle.size();
        if (!line.isEmpty()) regular=line.size();
        //int tablesCounter = 0; //счетчик номеров таблиц
        int [] tablesCounter = new int [1]; //счетчик номеров таблиц
        String fraza1 = "Всего осмотров: "+vsegoOsm+", в т.ч. предрейсовых – "+before;
        String fraza2 = "Допусков, всего – "+dopuskov+", не допусков – "+nedopuskov+", что составило "+String.format("%.2f", procentNedopuskov*100)+"% от общего числа медосмотров.";
        String fraza3 = "Всего осмотрено сотрудников: "+chisloVoditelei+" чел.";
        String dobavka ="";
        if (after>0) dobavka = dobavka+", послерейсовых – "+after;
        if (regular>0) dobavka = dobavka+", линейных – "+regular;
        dobavka = dobavka+".";
        fraza1 = fraza1+dobavka;
        // посчитали и подготовили текст (три фразы в три строки)

        //Blank Document
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);
        // создаем верхний колонтитул Word файла
        CTP ctpHeaderModel = createHeaderModel("Разработано " + copyright + "MDF-lab средствами Java");
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
        run.setText("Отчет по медицинским осмотрам сотрудников "+organization);   run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" "+god+" года");            //run.addCarriageReturn();

        XWPFParagraph paragraphText = document.createParagraph();
        paragraphText.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun runText = paragraphText.createRun();
        runText.setFontFamily("Times New Roman");
        runText.setFontSize(12);
        //runText.addCarriageReturn(); //возможно убрать пустую строку
        runText.setText(fraza1); runText.addCarriageReturn();
        runText.setText(fraza2); runText.addCarriageReturn();
        runText.setText(fraza3); //runText.addCarriageReturn();
        //  до табличных реестров выводим надпись вида
        //  Всего осмотров: 652, в т.ч. предрейсовых – 636, послерейсовых – 16.
        //  Допусков, всего – 547, не допусков – 105, что составило 16,1% от общего числа медосмотров.
        //  Всего осмотрено водителей: 209 чел.

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

        document.write(out); //сохраняем файл отчета в Word
        out.close();
        document.close();
        return res;
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
                Voditeli.add(st0.get(6));
            }
        }
        if (!posle.isEmpty()){
            for (ArrayList<String> st1: posle) {
                Voditeli.add(st1.get(6));
            }
        }
        if (!line.isEmpty()){
            for (ArrayList<String> st2: line) {
                Voditeli.add(st2.get(6));
            }
        }
        res = Voditeli.size();
        return res;
    }

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
                if (st1.get(16).equals("Допущен")) {
                    res++;
                }
            }
        }
        if (!line.isEmpty()) {
            for (ArrayList<String> st2 : line) {
                if (st2.get(16).equals("Допущен")) {
                    res++;
                }
            }
        }
        return res;
    }

    private int countNedopusk(List<ArrayList<String>> pred, List<ArrayList<String>> posle, List<ArrayList<String>> line) {
        int res = 0;
        if (!pred.isEmpty()){
            for (ArrayList<String> st0: pred) {
                if (st0.get(16).equals("Не допущен")){
                    res++;
                }
            }
        }
        if (!posle.isEmpty()) {
            for (ArrayList<String> st1 : posle) {
                if (st1.get(16).equals("Не допущен")) {
                    res++;
                }
            }
        }
        if (!line.isEmpty()) {
            for (ArrayList<String> st2 : line) {
                if (st2.get(16).equals("Не допущен")) {
                    res++;
                }
            }
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
    private String getOrganizationNameFromXLS (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        res = row.replaceAll("\"", "");

//        //разбиваем строку по пробелам
//        String[] tempArray = row.split(" ");
//        //собираем название организации (из последних элементов временного массива, т.е. кроме первого)
//        for (int i=1; i<tempArray.length; i++){
//            res = res+" "+tempArray[i].replaceAll("\"", "");
//        }
        return res.trim();
    }

    //получение из первой строки Excel отчетного месяца
    private String getMonthXLS (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        res = tempArray[0];
//        Locale rLocale = new Locale("ru"); //русская локаль
//        //SimpleDateFormat formatter = new SimpleDateFormat("dd MMM yyyy", Locale.US);
//        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy", rLocale);
//        SimpleDateFormat newFormatter = new SimpleDateFormat("MMMM", rLocale);
//
//        try {
//            Date date = formatter.parse(tempArray[1]);
//            res = newFormatter.format(date);
//
//        } catch (ParseException e) {
//            e.printStackTrace();
//        }

        return res.trim();
    }

    //получение из первой строки Excel отчетного месяца
    private String getGodXLS (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        res = tempArray[1];

//        Locale rLocale = new Locale("ru"); //русская локаль
//        //SimpleDateFormat formatter = new SimpleDateFormat("dd MMM yyyy", Locale.US);
//        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy", rLocale);
//        SimpleDateFormat newFormatter = new SimpleDateFormat("yyyy", rLocale);
//
//        try {
//            Date date = formatter.parse(tempArray[1]);
//            res = newFormatter.format(date);
//
//        } catch (ParseException e) {
//            e.printStackTrace();
//        }

        return res.trim();
    }

    //получение из первой строки Excel название таблицы
    private String getTableName (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);

        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        //собираем название организации (без начальных и без последних четырех элементов временного массива)
        for (int i=5; i<tempArray.length-4; i++){
            res = res+" "+tempArray[i];
        }
        return res.trim();
    }

    //получение из первой строки Excel название компании
    private String getOrganizationName (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);

        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        //собираем название организации (без начальных и без последних четырех элементов временного массива)
        for (int i=5; i<tempArray.length-4; i++){
            res = res+" "+tempArray[i];
        }
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
        return res.trim();
    }

    //получение из первой строки Excel отчетного месяца
    private String getMonth (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        res = tempArray[tempArray.length-3];
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
    private String getGod (ArrayList<String> firsRow){
        String res = "";
        String row = firsRow.get(0);
        //разбиваем строку по пробелам
        String[] tempArray = row.split(" ");
        res = tempArray[tempArray.length-2];
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

        return fileTree;
    }
}

