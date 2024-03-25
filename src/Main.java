import java.util.zip.ZipFile;
import java.util.zip.ZipEntry;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Main {
    public static void main(String[] args) {
        System.out.println("=======================DOCX========================");
        String docxFile = "sample1.docx";
        parseOOXML(docxFile);

        System.out.println("=======================PPTX========================");
        String pptxFile = "sample2.pptx";
        parseOOXML(pptxFile);

        System.out.println("=======================XLSX========================");
        String xlsxfile = "sample3.xlsx";
        parseOOXML(xlsxfile);
    }

    static void parseOOXML(String filePath) {
        try (ZipFile zipFile = new ZipFile(filePath)) {
            ZipEntry entry = zipFile.getEntry("_rels/.rels");
            if (entry != null) {
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(zipFile.getInputStream(entry));
                System.out.println("[Relationships (_rels/.rels)]");

                NodeList relations = doc.getElementsByTagName("Relationship");
                for (int i = 0; i < relations.getLength(); i++) {
                    Element relationship = (Element) relations.item(i);
                    System.out.println("Type: " + relationship.getAttribute("Type"));
                    System.out.println("Target: " + relationship.getAttribute("Target"));
                    System.out.println("ID: " + relationship.getAttribute("Id"));
                }
                System.out.println("---------------------------------------------------");
            }

            entry = zipFile.getEntry("docProps/app.xml");
            if (entry != null) {
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(zipFile.getInputStream(entry));
                System.out.println("[App Properties (docProps/app.xml)]");

                Element root = doc.getDocumentElement();
                NodeList nodeList = root.getChildNodes();
                for (int i = 0; i < nodeList.getLength(); i++) {
                    Node node = nodeList.item(i);
                    if (node.getNodeType() == Node.ELEMENT_NODE) {
                        Element element = (Element) node;
                        String tag = element.getTagName();
                        String content = element.getTextContent();
                        System.out.println(tag + ": " + content);
                    }
                }
                System.out.println("---------------------------------------------------");
            }

            entry = zipFile.getEntry("docProps/core.xml");
            if (entry != null) {
                DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
                Document doc = dBuilder.parse(zipFile.getInputStream(entry));
                System.out.println("[Document Properties (docProps/core.xml)]");

                Element root = doc.getDocumentElement();
                NodeList nodeList = root.getChildNodes();
                for (int i = 0; i < nodeList.getLength(); i++) {
                    Node node = nodeList.item(i);
                    if (node.getNodeType() == Node.ELEMENT_NODE) {
                        Element element = (Element) node;
                        String tagName = element.getTagName();
                        String textContent = element.getTextContent();
                        System.out.println(tagName + ": " + textContent);
                    }
                }
                System.out.println("---------------------------------------------------");
            }

            String DocumentByType;
            switch (filePath.substring(filePath.lastIndexOf(".") + 1)) {
                case "docx" -> DocumentByType = "word/document.xml";
                case "pptx" -> DocumentByType = "ppt/presentation.xml";
                case "xlsx" -> DocumentByType = "xl/workbook.xml";
                default -> {
                    System.out.println("Not supporting File type");
                    return;
                }
            }

            entry = zipFile.getEntry(DocumentByType);
            if (entry != null) {
                DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
                Document doc = dBuilder.parse(zipFile.getInputStream(entry));

                switch (DocumentByType) {
                    case "word/document.xml" -> System.out.println("[Word Document (word/document.xml)]");
                    case "ppt/presentation.xml" -> System.out.println("[PPT Presentation (ppt/presentation.xml)]");
                    case "xl/workbook.xml" -> System.out.println("[Xl Workbook (xl/workbook.xml)]");
                    default -> {
                        System.out.println("Not supporting File type");
                        return;
                    }
                }

                Element root = doc.getDocumentElement();
                NodeList nodeList = root.getChildNodes();
                for (int i = 0; i < nodeList.getLength(); i++) {
                    Node node = nodeList.item(i);
                    if (node.getNodeType() == Node.ELEMENT_NODE) {
                        Element element = (Element) node;
                        String tagName = element.getTagName();
                        String textContent = element.getTextContent();
                        System.out.println(tagName + ": " + textContent);
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Error : parsingOOXML");
        }
    }
}
