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

        //инициализируем потоки
        String table1FileName = "";                 // название файла Word с отчетной таблицей 1 (для скачивания)
        String table2FileName = "";                 // название файла Word с отчетной таблицей 2 (для скачивания)
        //InputStream inputStream;                  // поток чтения для загружаемого файла
        XSSFWorkbook workBookXLSX;                  // объект книги эксель xlsx
        HSSFWorkbook myExcelBookXLS = null;         // объект книги эксель xls
        //String[] stroka = new String[20];         // строка таблицы с листа
        //String[] customs = null;
        List<ArrayList<String>> list = new ArrayList<>();     // массив строк листа (кажда строка - массив строк) для medpont24
        List<ArrayList<String>> listPosleReis = new ArrayList<>(); // массив строк листа (кажда строка - массив строк) для medpont24
        List<ArrayList<String>> listP = new ArrayList<>();    // массив строк листа (кажда строка - массив строк) для поликлиники
        TreeMap<Integer, Integer[]> medOsmotryByDatesPredReis = new TreeMap<Integer, Integer[]>(); //итоговые данные отсортированы по дате
        //т.е. здесть Integer Key - дата мед.осм.
        //Integer[] Value - таблица допущено / не допущено (в эту дату)
        TreeMap<Integer, Integer[]> medOsmotryByDatesPosleReis;

        TreeMap<Integer, int[]> medOsmotryByDatesALL = new TreeMap<Integer, int[]>();
        //т.е. здесть Integer Key - дата мед.осм.
        //int[] Value - таблица: общ.кол|предр|допущ|недоущ|послер| (в эту дату)

        //итоговые данные отсортированы по дате
        TreeMap<Integer, Integer[]> medOsmotryByDatesXLS;
        //т.е. здесть Integer Key - дата мед.осм.
        //Integer[] Value - таблица предрейс / послерейс (в эту дату)

        //Массив дат медосмотров (для Табл.№2)
        ArrayList<Integer> dates = new ArrayList<>();

        //итоговые данные отсортированы по фамилиям и дате
        TreeMap<String, int[]> medOsmotryByFIOXLS;
        TreeMap<String, int[]> medOsmotryByFIO;
        // здесь key   это ФИО водителя - String
        // здесь value это таблица с суммарным значением предрейса и послерейса в каждой ячейке,
        // причем длина массива равна длине массива дат dates

        //инициализия завершена

        ///////////////WORK///////////////////
        //получаем части (нужные нам файлы)
        Part part = request.getPart("file");
        long size = part.getSize(); //файл медпойнта

        Part part_p = request.getPart("file_p");
        long size_p = part_p.getSize(); // файл поликлиники

        //обрабатываем файлы в зависимости от того, что загружено:
        //ничего
        if (size == 0 & size_p ==0){
            request.setAttribute("message", "Не загружен ни один отчёт :(");
            RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
            requestDispatcher.forward(request, response);
            return;
        }
        // только отчет medpoint24 загружен
        if (!(size==0) & size_p ==0){
            //получаем объект книги XLSX из формы
            workBookXLSX = XLSXFromPart(part);
            //разбираем первый лист файла medpoint24 на объектную модель
            list = getListFromSheet(workBookXLSX, 0); //получаем лист предрейса
            listPosleReis = getListFromSheet(workBookXLSX, 1); //получаем лист послерейса
            ArrayList<String> pervayaStroka = list.get(0); //первая строка (заголовок)
            organization = getOrganizationName_v2(pervayaStroka); //достаем из первой строки (заголовка) название компании.
            period = getMonth_v2(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
            god = getGod_v2(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.

            //Причесываем списки:
            // убираем заголовок таблицы, убираем шапку таблицы, убирем последние 7 ненужных строк
            list = list.subList(2, list.size()-7);
            listPosleReis = listPosleReis.subList(2, listPosleReis.size()-7);

            //производим подсчёт по предрейсовым
            medOsmotryByDatesPredReis = prepare(list);

            //производим подсчёт по послерейсовым
            medOsmotryByDatesPosleReis = prepare(listPosleReis);

            //производим подсчёт по предрейсовым и послерейсовым за раз
            // (Табл.2 Детализация, по фамилиям)
            ///medOsmotryByFIO = prepareTable2(listP, dates);



            // TODO: 09.09.2020 Суммарная таблица предрейса и послерейса для формирования Word отчета
            Integer pred = medOsmotryByDatesPredReis.size();  //сколько дат предрейса
            Integer posl = medOsmotryByDatesPosleReis.size(); //сколько дат послерейса
            int hvost; //сколько дней разница
            //организуем суммирование таблиц по списку с меньшим количеством дат
            if (pred>=posl) { //подсчет по послерейсу
                hvost=pred-posl;
                for (Map.Entry<Integer, Integer[]> entry: medOsmotryByDatesPosleReis.entrySet()
                     ) {
                    Integer key = entry.getKey(); //получаем дату
                    Integer[] tempPosler = entry.getValue();                //значения послерейса в эту дату
                    Integer[] tempPredr = new Integer[]{0,0};
                    int predrVsego = 0;
                    int predrProshlo = 0;
                    int vsegoNeProshlo = 0;
                    if (medOsmotryByDatesPredReis.containsKey(key)){     //если такая дата есть в другом списке, т.е. в предрейсе
                        tempPredr = medOsmotryByDatesPredReis.get(key); //значения предрейса в эту дату
                        predrVsego = tempPredr[0]+tempPredr[1];         //всего предрейсовых в эту дату
                        predrProshlo = tempPredr[0];                    //допушено в эту дату
                        vsegoNeProshlo = tempPredr[1]+tempPosler[1];   //не прошло предр. и послер. в эту дату
                    } else {
                        //нули уже установлены
                    }

                    int poslerVsego = tempPosler[0]+tempPosler[1];     //всего послерейсовых в эту дату
                    int vsegoOsmotrov = predrVsego+poslerVsego;        // всего осмотров в эту дату

                    int[] currentValue = new int[5]; //готовим таблицу с нулями
                    //заполняем её: общ.кол|предр|допущ|недоущ|послер| (в эту дату)
                    currentValue[0] = vsegoOsmotrov;
                    currentValue[1] = predrVsego;
                    currentValue[2] = predrProshlo;
                    currentValue[3] = vsegoNeProshlo;
                    currentValue[4] = poslerVsego;

                    //заносим в итоговую мапу
                    medOsmotryByDatesALL.put(key, currentValue);
                }
                //если есть хвост то дополняем итоговую мапу
                if(!(hvost==0)){
                    //сначала достаем отсутствующие даты (которые из списка с большим количеством дат)
                    Set<Integer> notTakenDates = new HashSet<>();
                    Set<Integer> medPosleReis = new HashSet<>();
                    notTakenDates.addAll(medOsmotryByDatesPredReis.keySet());
                    medPosleReis.addAll(medOsmotryByDatesPosleReis.keySet());
                    notTakenDates.removeAll(medPosleReis); //теперь здесь не обработанные даты предрейса

                    for (Integer takeDate:notTakenDates) {
                        Integer[] predr = medOsmotryByDatesPredReis.get(takeDate); //значения предрейса в эту дату
                        //Integer[] posler = medOsmotryByDatesPosleReis.get(takeDate); //значения послерейса в эту дату
                        int predrSum = predr[0]+predr[1];         //всего предрейсовых в эту дату
                        int predrOK = predr[0];                    //допушено в эту дату
                        int vsegoNotOK = predr[1]/*+posler[1]*/;               //не прошло предр. в эту дату
                        int poslerSum =0 /*predr[1]+posler[1]*/;  //всего послерейсовых в эту дату = 0 т.к. это данные послерейса, а послерейс в эту дату не проводился
                        int vsegoSum = predrSum+poslerSum;        // всего осмотров в эту дату

                        int[] tempValues = new int[5]; //готовим таблицу с нулями
                        //заполняем её: общ.кол|предр|допущ|недоущ|послер| (в эту дату)
                        tempValues[0] = vsegoSum;
                        tempValues[1] = predrSum;
                        tempValues[2] = predrOK;
                        tempValues[3] = vsegoNotOK;
                        tempValues[4] = poslerSum;

                        //заносим в итоговую мапу
                        medOsmotryByDatesALL.put(takeDate, tempValues);
                    }
                }

            }else{ //подсчет по предрейсу, т.к. у него кол-во дат меньше
                hvost=posl-pred;
                for (Map.Entry<Integer, Integer[]> entry: medOsmotryByDatesPredReis.entrySet()
                        ) {
                    Integer key = entry.getKey(); //получаем дату
                    Integer[] tempPred = entry.getValue();                //значения предрейса в эту дату
                    Integer[] tempPosle = new Integer[]{0,0};           //значения послерейса в эту дату
                    tempPosle[0] = 0; tempPosle[1] = 0;
                    int predrVsego = tempPred[0]+tempPred[1];         //всего предрейсовых в эту дату
                    int poslerVsego = 0;                                 //всего послерейсовых в эту дату
                    int predrProshlo = tempPred[0];                    //допушено в эту дату
                    int vsegoNeProshlo =0;                                //не прошло предр. и послер. в эту дату
                    if (medOsmotryByDatesPosleReis.containsKey(key)){     //если такая дата есть в другом списке
                        tempPosle = medOsmotryByDatesPosleReis.get(key); //значения предрейса в эту дату
                        poslerVsego = tempPosle[0]+tempPosle[1];
                        vsegoNeProshlo = tempPred[1]+tempPosle[1];    //добавляем не прошедших послерейс в эту дату

                    } else {
                        //нули уже установлены
                    }
                    //int poslerVsego = tempPosler[0]+tempPosler[1];     //всего послерейсовых в эту дату
                    int vsegoOsmotrov = predrVsego+poslerVsego;        // всего осмотров в эту дату

                    int[] currentValue = new int[5]; //готовим таблицу с нулями
                    //заполняем её: общ.кол|предр|допущ|недоущ|послер| (в эту дату)
                    currentValue[0] = vsegoOsmotrov;
                    currentValue[1] = predrVsego;
                    currentValue[2] = predrProshlo;
                    currentValue[3] = vsegoNeProshlo;
                    currentValue[4] = poslerVsego;

                    //заносим в итоговую мапу
                    medOsmotryByDatesALL.put(key, currentValue);
                }
                //если есть хвост то дополняем итоговую мапу
                if(!(hvost==0)){
                    //сначала достаем отсутствующие даты (которые из списка с большим количеством дат)
                    Set<Integer> notTakenDates = new HashSet<>();
                    Set<Integer> medPredReis = new HashSet<>();
                    notTakenDates.addAll(medOsmotryByDatesPosleReis.keySet());
                    medPredReis.addAll(medOsmotryByDatesPredReis.keySet());
                    notTakenDates.removeAll(medPredReis); //теперь здесь не обработанные даты

                    for (Integer takeDate:notTakenDates) {
                        //Integer[] predr = medOsmotryByDatesPredReis.get(takeDate);  //значения предрейса в эту дату
       /*тут Null*/     Integer[] posler = medOsmotryByDatesPosleReis.get(takeDate); //значения предрейса в эту дату
       /*!!!!*/         int predrSum = /*predr[0]+predr[1]*/ 0;       //всего предрейсовых в эту дату 0, т.к. считаем по листу послерейса
                        int predrOK = /*predr[0]*/ 0;                //допушено в эту дату
                        int vsegoNotOK = /*predr[1]+*/posler[1];    //не прошло предр. и послер. в эту дату
                        int poslerSum = posler[0]+posler[1];       //всего послерейсовых в эту дату т.к. это данные послерейса, а предрейс в эту дату не проводился
                        int vsegoSum = predrSum+poslerSum;        // всего осмотров в эту дату

                        int[] tempValues = new int[5]; //готовим таблицу с нулями
                        //заполняем её: общ.кол|предр|допущ|недоущ|послер| (в эту дату)
                        tempValues[0] = vsegoSum;
                        tempValues[1] = predrSum;
                        tempValues[2] = predrOK;
                        tempValues[3] = vsegoNotOK;
                        tempValues[4] = poslerSum;

                        //заносим в итоговую мапу
                        medOsmotryByDatesALL.put(takeDate, tempValues);
                    }
                }
            }

            //получаем массив дат
            for ( Integer keys:medOsmotryByDatesALL.keySet() ) {
                dates.add(keys);
            }
            // (Табл.2 Детализация, по фамилиям) предрейс+послерейс
            medOsmotryByFIO = prepareTable2(list, listPosleReis, dates);

            // gets absolute path of the web application
            String applicationPath = request.getServletContext().getRealPath("");
            // constructs path of the directory to save uploaded file
            String uploadFilePath = applicationPath + File.separator + REPORTS_DIR;

            //Создаем папку для формируемых отчетов Word если ее нет
            File uploadFolder = new File(uploadFilePath);
            if (!uploadFolder.exists()) {  //если папки не существует, то создаем
                uploadFolder.mkdirs();
            }

            try {   //заменить на суммарый с послерейсом +
                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table1FileName) в JSP
                table1FileName = makeWordDocumentTable1(medOsmotryByDatesALL, uploadFilePath);

                //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table2FileName)
                table2FileName = makeWordDocumentTable2XLS(dates, medOsmotryByFIO, uploadFilePath);

            } catch (XmlException e) {
                e.printStackTrace();
                //response.setContentType("text/html");
            }

            response.setContentType("text/html");
            response.setCharacterEncoding("UTF-8");
            request.setCharacterEncoding("UTF-8");
            request.setAttribute("docxName", table1FileName);
            request.setAttribute("docx2Name", table2FileName);
            request.setAttribute("reportsDir", REPORTS_DIR);
            request.setAttribute("message", "Отчёт по выгрузке medpoint24 сформирован успешно!");
            RequestDispatcher requestDispatcher = request.getRequestDispatcher("otchet.jsp");
            requestDispatcher.forward(request, response);
            return;
        }
        // только отчет поликлиники загружен
        if (size==0 & !(size_p == 0)){
            //получаем объект книги XLS из формы
            myExcelBookXLS = XLSFromPart(part_p);
            //разбираем лист "Реестр" файла поликлиники на объектную модель
            try {
                listP = getListFromSheetXLS(myExcelBookXLS);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (NullPointerException e) {
                e.printStackTrace();
                debug = e.toString();
                message = "Ошибка при формировании отчета - проверьте структуру таблицы: проблемная строка №"+String.valueOf(errorStringNumber);
                failed = true;

//                request.setAttribute("message", "Ошибка при формировании отчета - проверьте структуру таблицы: проблемная строка №"+String.valueOf(errorStringNumber));
//                request.setAttribute("debug", debug);
//                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
//                requestDispatcher.forward(request, response);
//                errorStringNumber = 0;
//                return;

            } catch (IllegalStateException e) {
                e.printStackTrace();
                debug = e.toString();
                message = "Ошибка при формировании отчета - проверьте корректность даты в строке №"+String.valueOf(errorStringNumber);
                failed = true;
//                String debug = e.toString();
//                request.setAttribute("message", "Ошибка при формировании отчета - проверьте корректность даты в строке №"+String.valueOf(errorStringNumber));
//                request.setAttribute("debug", debug);
//                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
//                requestDispatcher.forward(request, response);
//                errorStringNumber = 0;
//                return;
            }

            if (failed) { //если есть проблема - выводим сообщение об ошибке и сбрасываем маркер проблемы
                request.setAttribute("message", message);
                request.setAttribute("debug", debug);
                RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
                requestDispatcher.forward(request, response);
                failed = false;
                return;
            } else { // если проблем нет - выводим работу программы
                organization = getOrganizationNameFromXLS(listP.get(0)); //достаем строку, содержащую название компании.
                period = getMonthXLS(listP.get(1)); //достаем из заголовка отчетный месяц.
                god = getGodXLS(listP.get(1)); //достаем из заголовка отчетный год.

                //Причесываем список:
                // убираем заголовок таблицы, убираем шапку таблицы
                listP = listP.subList(3, listP.size());

                //производим подсчёт по предрейсовым и послерейсовым за раз
                // (Табл.1 Фактическая, по датам)
                medOsmotryByDatesXLS = prepareXLS(listP);

                //получаем массив дат
                for ( Integer keys:medOsmotryByDatesXLS.keySet() ) {
                      dates.add(keys);
                }

                //производим подсчёт по предрейсовым и послерейсовым за раз
                // (Табл.2 Детализация, по фамилиям)
                medOsmotryByFIOXLS = prepareTable2XLS(listP, dates);

                // gets absolute path of the web application
                String applicationPath = request.getServletContext().getRealPath("");
                // constructs path of the directory to save uploaded file
                String uploadFilePath = applicationPath + File.separator + REPORTS_DIR;

                //Создаем папку для формируемых отчетов Word если ее нет
                File uploadFolder = new File(uploadFilePath);
                if (!uploadFolder.exists()) {  //если папки не существует, то создаем
                    uploadFolder.mkdirs();
                }

                try {
                    //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table1FileName) в JSP
                    table1FileName = makeWordDocumentTable1XLS(medOsmotryByDatesXLS, uploadFilePath);

                    //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table2FileName)
                    table2FileName = makeWordDocumentTable2XLS(dates, medOsmotryByFIOXLS, uploadFilePath);
                } catch (XmlException e) {
                    e.printStackTrace();
                }

                response.setContentType("text/html");
                response.setCharacterEncoding("UTF-8");
                request.setCharacterEncoding("UTF-8");

                //System.out.println(organization);

                request.setAttribute("message", "Отчёт по реестру сформирован успешно!");
                request.setAttribute("title", "Результат");
                request.setAttribute("size", size);
                request.setAttribute("list", list);
                request.setAttribute("medOsmotryByDates", medOsmotryByDatesXLS); //заменить на суммарый с послерейсом
                request.setAttribute("docxName", table1FileName);
                request.setAttribute("docx2Name", table2FileName);
                request.setAttribute("reportsDir", REPORTS_DIR);
                RequestDispatcher requestDispatcher = request.getRequestDispatcher("otchet.jsp");
                requestDispatcher.forward(request, response);
                return;
            }


        }
        //Оба отчета загружено
        else {
        /*
        //получаем объект книги XLSX из формы
        workBookXLSX = XLSXFromPart(part);
        //получаем объект книги XLS из формы
        myExcelBookXLS = XLSFromPart(part_p);


        //разбираем первый лист файла medpoint24 на объектную модель
        list = getListFromSheet(workBookXLSX, 0);

        //разбираем лист "Лист1" файла поликлиники на объектную модель
        listP = getListFromSheetXLS(myExcelBookXLS);

        ArrayList<String> pervayaStroka = list.get(0); //первая строка (заголовок)
        organization = getOrganizationName(pervayaStroka); //достаем из первой строки (заголовка) название компании.
        period = getMonth(pervayaStroka); //достаем из первой строки (заголовка) отчетный месяц.
        god = getGod(pervayaStroka); //достаем из первой строки (заголовка) отчетный год.

        //Причесываем список:
        // убираем заголовок таблицы, убираем шапку таблицы, убирем последние 7 ненужных строк
        list = list.subList(2, list.size()-7);

        //производим подсчёт по предрейсовым
        medOsmotryByDatesPredReis = prepare(list);

        // TODO: 09.09.2020
        //производим подсчёт по послерейсовым
        //medOsmotryByDatesPosleReis = prepare(list);

        // TODO: 09.09.2020 Суммарная таблица предрейса и послерейса для формирования Word отчета

        // gets absolute path of the web application
        String applicationPath = request.getServletContext().getRealPath("");
        // constructs path of the directory to save uploaded file
        String uploadFilePath = applicationPath + File.separator + REPORTS_DIR;

        //Создаем папку для формируемых отчетов Word если ее нет
        File uploadFolder = new File(uploadFilePath);
        if (!uploadFolder.exists()) {  //если папки не существует, то создаем
            uploadFolder.mkdirs();
        }

        try {   //// TODO: 26.09.2020  заменить на суммарый: medpoint24+поликлиника
            //готовим отчет в ворде и сохраняем в папке отчетов, выдаем название файла для его скачивания (table1FileName) в JSP
            table1FileName = makeWordDocumentTable1(medOsmotryByDatesALL, uploadFilePath);
        } catch (XmlException e) {
            e.printStackTrace();
            //response.setContentType("text/html");
        }

        response.setContentType("text/html");
        response.setCharacterEncoding("UTF-8");
        request.setCharacterEncoding("UTF-8");

        request.setAttribute("message", "Отчёт сформирован успешно!");
        request.setAttribute("title", "Результат");
        request.setAttribute("size", size);
        request.setAttribute("list", list);
        request.setAttribute("medOsmotryByDates", medOsmotryByDatesPredReis); //заменить на суммарый с послерейсом
        request.setAttribute("docxName", table1FileName);
        request.setAttribute("reportsDir", REPORTS_DIR);
        RequestDispatcher requestDispatcher = request.getRequestDispatcher("otchet.jsp"); */
        RequestDispatcher requestDispatcher = request.getRequestDispatcher("pusto.jsp");
            request.setAttribute("message", "Функционал объединения двух отчетов в разработке.");
            request.setAttribute("debug", "-");
        requestDispatcher.forward(request, response);}
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
                        default:
                            break;
                    }
                }
                res.add(tempStringArray);
            }
            workBook.close();
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
            switch (stroka.get(16)){ // было 15
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

    private TreeMap<String, int[]> prepareTable2 (List<ArrayList<String>> spisokVesPred,
                                                  List<ArrayList<String>> spisokVesPosl,
                                                  ArrayList<Integer> alldates) {
        //заготовка для результата
        TreeMap<String, int[]> result = new TreeMap<>();
        int vsegoDat = alldates.size();

        // foreach для предрейса, потом для послерейса
        for (ArrayList<String> stroka : spisokVesPred) { //пробегаемся по строкам
            int[] calendDates = new int[vsegoDat]; //готовим таблицу дат осмотров для каждой фамилии
            String fio = stroka.get(5); // получаем ФИО из 5й ячейки строки
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
            String fio = stroka.get(5); // получаем ФИО из 5й ячейки строки
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
    private String makeWordDocumentTable1(TreeMap<Integer, int[]> preparedList, String uploadFilePath) throws IOException, XmlException {
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
        run.setText("за "+period.toLowerCase()+" месяц "+god+" года"); //todo: год тоже надо вытаскивать из эксель +
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
        tableRowOne.getCell(4).setParagraph(fillParagraph(document, "Количество водителей, допущенных к работе"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(5).setParagraph(fillParagraph(document, "Количество водителей, не допущенных к работе"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(6).setParagraph(fillParagraph(document, "Количество послерейсовых мед.осмотров"));


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

            //create next rows
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(fillParagraph(document, Integer.toString(count)));
            tableRowNext.getCell(1).setParagraph(fillParagraph(document, Integer.toString(key)));   //день месяца
            tableRowNext.getCell(2).setParagraph(fillParagraph(document, Integer.toString(value[0]))); //всего мед.осм.
            tableRowNext.getCell(3).setParagraph(fillParagraph(document, Integer.toString(value[1]))); //предрейсовых.
            tableRowNext.getCell(4).setParagraph(fillParagraph(document, Integer.toString(value[2]))); //допущ.
            tableRowNext.getCell(5).setParagraph(fillParagraph(document, Integer.toString(value[3]))); //не допущ.
            tableRowNext.getCell(6).setParagraph(fillParagraph(document, Integer.toString(value[4]))); //послерейсовых.
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

    private String makeWordDocumentTable1XLS (TreeMap<Integer, Integer[]> preparedList, String uploadFilePath) throws IOException, XmlException {
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
        run.setText("за "+period.toLowerCase()+" месяц "+god+" года"); //todo: год тоже надо вытаскивать из эксель +
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
        tableRowOne.getCell(4).setParagraph(fillParagraph(document, "Количество водителей, допущенных к работе"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(5).setParagraph(fillParagraph(document, "Количество водителей, не допущенных к работе"));

        tableRowOne.addNewTableCell();
        tableRowOne.getCell(6).setParagraph(fillParagraph(document, "Количество послерейсовых мед.осмотров"));


        //table.getRow(0).getCell(0).addParagraph();

        Iterator iterator = preparedList.keySet().iterator();
        int count = 0;          //счетчик строк таблицы
        int countDopusk = 0;    //счетчик допусков
        int countNoDopusk = 0;  //счетчик не допусков
        int countMedOsm = 0;    //счетчик мед.осмотров
        int countPoslereis = 0; //счетчик послерейс.мед.осмотров
        while(iterator.hasNext()) {
            count++;
            Integer key   =(Integer) iterator.next();
            Integer[] value = preparedList.get(key);
            int vsego = value[0]+value[1];  //всего осмотров за этот день (предрейс + послерейс)
            countDopusk = countDopusk + value[0]; //все из предрейса допущены
            countPoslereis = countPoslereis + value[1];
            countNoDopusk = countNoDopusk + value[2];
            countMedOsm = countMedOsm + vsego;

            //create next rows
            XWPFTableRow tableRowNext = table.createRow();
            tableRowNext.getCell(0).setParagraph(fillParagraph(document, Integer.toString(count)));
            tableRowNext.getCell(1).setParagraph(fillParagraph(document, Integer.toString(key)));      //день месяца
            tableRowNext.getCell(2).setParagraph(fillParagraph(document, Integer.toString(vsego)));    //всего мед.осм.
            tableRowNext.getCell(3).setParagraph(fillParagraph(document, Integer.toString(value[0]))); //предрейс.
            tableRowNext.getCell(4).setParagraph(fillParagraph(document, Integer.toString(value[0]))); //допущ.
            tableRowNext.getCell(5).setParagraph(fillParagraph(document, Integer.toString(value[2]))); //не допущ.
            tableRowNext.getCell(6).setParagraph(fillParagraph(document, Integer.toString(value[1]))); //послерейс.
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
        tableRowLast.getCell(3).setParagraph(fillParagraph(document, Integer.toString(countDopusk)));   //предрейс.
        //tableRowLast.getCell(4).setParagraph(paragraph);
        tableRowLast.getCell(4).setParagraph(fillParagraph(document, Integer.toString(countDopusk)));   //допущ.
        tableRowLast.getCell(5).setParagraph(fillParagraph(document, Integer.toString(countNoDopusk))); //не допущ.
        tableRowLast.getCell(6).setParagraph(fillParagraph(document, Integer.toString(countPoslereis))); //послерейс.

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

    private String makeWordDocumentTable2XLS(ArrayList<Integer> alldates,
                                             TreeMap<String, int[]> medOsmotryByFIOXLS,
                                             String uploadFilePath) throws IOException, XmlException {
        String copyright = "\u00a9";
        String res = File.separator+organization+" (детализ.) ["+period.toLowerCase()+"] "
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
        run.setText("Детализация прохождения предрейсовых/послерейсовых"); run.addCarriageReturn();
        run.setText("медицинских осмотров водителей"); run.addCarriageReturn();
        run.setText(organization);run.addCarriageReturn();
        run.setText("за "+period.toLowerCase()+" месяц "+god+" года");run.addCarriageReturn();
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
        tableRowOne.getCell(1).setParagraph(fillParagraphBold(document, "ФИО \\ День месяца"));

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
        out.close();
        document.close();
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
