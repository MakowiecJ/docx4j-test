package poc;

import java.io.File;
import java.util.List;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import jakarta.xml.bind.JAXBException;

public class DynamicTableExample {

    public static void main(String[] args) {
        try {
            // Load the template document
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\TEST.docx"));
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            // Find and replace bookmarks with dynamic tables
            List<Object> bookmarks = mainDocumentPart.getJAXBNodesViaXPath("//w:bookmarkStart", false);
            for (Object bookmark : bookmarks) {
                CTBookmark bookmarkStart = (CTBookmark) bookmark;
                if (bookmarkStart.getName().equals("TABLE_BOOKMARK")) {
                    // Found the bookmark for the table
                    replaceBookmarkWithTable(mainDocumentPart, bookmarkStart);
                }
            }

            // Save the modified document
            wordMLPackage.save(new File("table_out.docx"));
        } catch (Docx4JException | JAXBException e) {
            e.printStackTrace();
        }
    }

    private static void replaceBookmarkWithTable(MainDocumentPart mainDocumentPart, CTBookmark bookmarkStart) throws JAXBException {
        // Retrieve client data (replace this with your own logic to fetch client data)
        List<Client> clients = getClientData();

        // Create a new table
        Tbl table = createTable(clients);

        // Find the parent of the bookmark and replace it with the table
        P paragraph = (P) bookmarkStart.getParent();
        paragraph.getContent().add(table);
        // Remove the bookmark from the document
        mainDocumentPart.getContent().remove(bookmarkStart);
    }

    private static Tbl createTable(List<Client> clients) {
        // Create a new table
        Tbl table = Context.getWmlObjectFactory().createTbl();
        // Create table header row
        Tr headerRow = createHeaderRow();
        table.getContent().add(headerRow);

        // Add rows for each client
        for (Client client : clients) {
            Tr row = createDataRow(client);
            table.getContent().add(row);
        }

        return table;
    }

    private static Tr createHeaderRow() {
        Tr headerRow = Context.getWmlObjectFactory().createTr();
        // Create header cells and add them to the row
        // Adjust as per your requirements
        addHeaderCell(headerRow, "ID");
        addHeaderCell(headerRow, "Name");
        addHeaderCell(headerRow, "Email");
        return headerRow;
    }

    private static void addHeaderCell(Tr row, String content) {
        Tc cell = Context.getWmlObjectFactory().createTc();
        var text = new Text();
        text.setValue(content);
        cell.getContent().add(text);
        row.getContent().add(cell);
    }

    private static Tr createDataRow(Client client) {
        Tr row = Context.getWmlObjectFactory().createTr();
        // Create data cells and add them to the row
        // Adjust as per your requirements
        addDataCell(row, client.getId());
        addDataCell(row, client.getName());
        addDataCell(row, client.getEmail());
        return row;
    }

    private static void addDataCell(Tr row, String content) {
        Tc cell = Context.getWmlObjectFactory().createTc();
        var text = new Text();
        text.setValue(content);
        cell.getContent().add(text);
        row.getContent().add(cell);
    }

    private static List<Client> getClientData() {
        // Replace this method with your own logic to fetch client data from a database or any other source
        // This is just a mock implementation
        return List.of(
                new Client("1", "John Doe", "john@example.com"),
                new Client("2", "Jane Smith", "jane@example.com")
        );
    }

    private static class Client {
        private String id;
        private String name;
        private String email;

        public Client(String id, String name, String email) {
            this.id = id;
            this.name = name;
            this.email = email;
        }

        public String getId() {
            return id;
        }

        public String getName() {
            return name;
        }

        public String getEmail() {
            return email;
        }
    }
}