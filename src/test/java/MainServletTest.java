
//import junit.framework.TestCase;
//import org.junit.*;
import junit.framework.TestCase;

import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;

import static junit.framework.Assert.assertEquals;


public class MainServletTest extends TestCase {

    //@Test
    public void testCountVozrast() {
        //создаем тестовые данные
        LocalDate firstDate = LocalDate.now(); ///of(2010, 5, 17); //
        LocalDate secondDate = LocalDate.of(2000, 5, 20); //
        //считаем
        long res = ChronoUnit.YEARS.between(secondDate, firstDate);
        //проверяем
        /*Assert.*/assertEquals(23, res);  //assertEquals();
    }

    //@Test
    public  void testSetSrednSAD() {
        //создаем тестовые данные
        ArrayList<Integer> srednSAD = new ArrayList<>();
        srednSAD.add(153);
        srednSAD.add(137);
        srednSAD.add(149);
        //считаем
        int sum = 0;
        for (int davlenie : srednSAD) {
            sum = sum + davlenie;
        }
        int res = Math.round(sum/srednSAD.size());
        //проверяем
        /*Assert.*/assertEquals(146, res);
    }

}