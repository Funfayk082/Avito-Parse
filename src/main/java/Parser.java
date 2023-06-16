import java.io.File;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Parser {
    public static String url = "https://www.avito.ru/kostroma/velosipedy?p=";
    public static Document getPage(int index, String link) throws IOException {
        Document page = Jsoup.connect(link+"&p="+index).userAgent("Opera").get();
        return page;
    }

    private static Pattern pattern = Pattern.compile("\\d{6} ");

    private static String getDateFromString(String string) throws Exception {
        Matcher matcher = pattern.matcher(string);
        if (matcher.find()) {
            return matcher.group();
        }
        throw new Exception("Невозможно преобразовать дату из строки!");
    }

    private static void printValues(Elements values, int index) {
        for (int i = 0; i < 4; i++) {
            Element valueLine = values.get(index);
            for (Element td: valueLine.select("td")) {
                System.out.print(td.text()+"    ");
            }
        }
    }

    public static void main(String[] args) throws Exception {
        System.out.print("Введите ссылку на страницу каталога Авито для сбора информации: ");
        Scanner sc = new Scanner(System.in);
        String link = sc.nextLine();
        Document page = getPage(1, link);
        String category = page.select("h1[data-marker=page-title/text]").first().text();
        Element maxIndex = page.select("span[class=styles-module-text-InivV]").last();
        int max = Integer.parseInt(maxIndex.text());
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "Данные.xls";

        WritableWorkbook workbook = Workbook.createWorkbook(new File(fileLocation));
        WritableSheet sheet = workbook.createSheet(category, 0);

        WritableCellFormat headerFormat = new WritableCellFormat();
        WritableFont font = new WritableFont(WritableFont.ARIAL, 16, WritableFont.BOLD);
        headerFormat.setFont(font);
        headerFormat.setBackground(Colour.SKY_BLUE);
        headerFormat.setWrap(true);

        Label headerLabel = new Label(0, 0, "Название", headerFormat);
        sheet.setColumnView(0, 60);
        sheet.addCell(headerLabel);

        headerLabel = new Label(1, 0, "Цена", headerFormat);
        sheet.setColumnView(0, 60);
        sheet.addCell(headerLabel);

        headerLabel = new Label(2, 0, "Ссылка", headerFormat);
        sheet.setColumnView(0, 60);
        sheet.addCell(headerLabel);
        /*Element mainTable = page.select("table[class=wt]").first();
        Elements names = mainTable.select("tr[class=wth]");
        Elements values = mainTable.select("tr[valign=top]");
        System.out.println( "Дата    Явления   Температура    Давление    Влажность   Ветер");
        int index = 0;
        for (Element name: names) {
            String data = name.select("th[id=dt]").text();
            String dataString = getDateFromString(data);
            System.out.print(dataString + "    ");
            printValues(values, index);
            index++;
            System.out.println("");
        }*/
        int index = 2;
        for (int i = 1; i < max; i++) {
            page = getPage(i, link);
            Element mainTable = page.select("div[class=items-items-kAJAg]").first();
            Elements elements = mainTable.select("div[data-marker=item]");
            for (Element e: elements) {
                String title = e.select("a[data-marker=item-title]").text();
                Element ee = e.select("a[data-marker=item-title]").get(0);
                String price = e.select("p[data-marker=item-price]").text();
                String linkString = ee.attr("href");

                WritableCellFormat cellFormat = new WritableCellFormat();
                cellFormat.setWrap(true);

                Label cellTitle = new Label(0, index, title, cellFormat);
                sheet.addCell(cellTitle);
                Label cellPrice = new Label(1, index, price, cellFormat);
                sheet.addCell(cellPrice);
                Label cellLink = new Label(2, index, (url+linkString), cellFormat);
                sheet.addCell(cellLink);
                index++;
            }
            TimeUnit.SECONDS.sleep(2);
        }
        workbook.write();
        workbook.close();
        System.out.println("Сбор информации завершён!!!");
    }
}
