import com.aspose.cells.Workbook;
import com.aspose.pdf.Document;

public class App {
    public static void main(String[] args) throws Exception {
        Workbook book = new Workbook("/Users/sitamsh.rijal/IdeaProjects/AsposeTinkering/src/main/java/input.xlsx");

        var pageSetup = book.getWorksheets().get(0).getPageSetup();
        var date = "Sep 27, 2022";
        pageSetup.setFooter(0, "&\"Arial\"&8&K02-074&K444444&B" + date + "&B");

        // save EXCEL as PDF
        book.save("pdfOutput.pdf", com.aspose.cells.SaveFormat.AUTO);

        // load the PDF file using Document class
        try (Document document = new Document("pdfOutput.pdf")) {
            // save document in DOC format
            document.save("output.doc", com.aspose.pdf.SaveFormat.Doc);
        }
    }
}