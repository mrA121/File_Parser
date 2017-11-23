package text_parser;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.rtf.RTFEditorKit;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.Text;
import org.jdom.input.SAXBuilder;
import org.jsoup.Jsoup;
@WebServlet("/Parser")
public class Parser extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    private class ReadFileFormat{
    	//StringBuffer sb = new StringBuffer(8192);
        StringBuffer TextBuffer = new StringBuffer();
    	String DocxToText(String FilePath) {
    		String parsedText = null;
        	try {
        		FileInputStream fis;
                fis = new FileInputStream(new File(FilePath));
                XWPFDocument doc = new XWPFDocument(fis);
                XWPFWordExtractor extract = new XWPFWordExtractor(doc);
                parsedText = extract.getText();
            } catch (IOException e) {

                e.printStackTrace();
            }
        	return parsedText;
        }
    	String DocToText(String FilePath) {
    		String parsedText = null;
    		try {
    			FileInputStream fis;
                fis = new FileInputStream(new File(FilePath));
                HWPFDocument doc = new HWPFDocument(fis);
                WordExtractor extractor = new WordExtractor(doc);
                parsedText = extractor.getText();
            } catch (IOException e) {
                e.printStackTrace();
            }return parsedText;
    	}
    	String PdfToText(String FilePath) {
    		String parsedText = null;;
    		PDFTextStripper pdfStripper = null;
    		PDDocument pdDoc = null;
    		File file = new File(FilePath);  
    		if (!file.isFile()) {
    			System.err.println("File " + FilePath + " does not exist.");
    			return null;
    		}
    		try {
    			pdDoc = PDDocument.load(file);
    			pdfStripper = new PDFTextStripper();
    			pdfStripper.setStartPage(1);
    			pdfStripper.setEndPage(3);
    			parsedText = pdfStripper.getText(pdDoc);
    		} catch (Exception e) {
    			System.err
    					.println("An exception occured in parsing the PDF Document."
    							+ e.getMessage());
    		} finally {
    			try {
    				if (pdDoc != null)
    					pdDoc.close();
    			} catch (Exception e) {
    				e.printStackTrace();
    			}
    		}
    		return parsedText;
    	}
    	String RtfToText(String FilePath) throws Exception {
            DefaultStyledDocument styledDoc = new DefaultStyledDocument();
            new RTFEditorKit().read(new FileInputStream(new File(FilePath)), styledDoc, 0);
            return styledDoc.getText(0, styledDoc.getLength());
        }
    	String HtmlToText(String FilePath) throws IOException {
            StringBuilder sb = new StringBuilder();
            BufferedReader br = new BufferedReader(new FileReader(FilePath));
            String line;
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
            String textOnly = Jsoup.parse(sb.toString()).text();
            br.close();
            return textOnly;
        }
    	String TxtToText(String Filepath) throws IOException{
    		File file = new File(Filepath);
    		Scanner in = new Scanner(file);
    		StringBuilder sb = new StringBuilder();
    		while(in.hasNextLine()) {
    		    sb.append(in.nextLine());
    		}
    		in.close();
    		String outString = sb.toString();
    		return outString;
    	}
    	public void processElement(Object o) {
            if (o instanceof Element) {
                Element e = (Element) o;
                String elementName = e.getQualifiedName();
                if (elementName.startsWith("text")) {
                    if (elementName.equals("text:tab")) // add tab for text:tab
                    {
                        TextBuffer.append("\t");
                    } else if (elementName.equals("text:s")) // add space for text:s
                    {
                        TextBuffer.append(" ");
                    } else {
                        List<?> children = e.getContent();
                        Iterator<?> iterator = children.iterator();
                        while (iterator.hasNext()) {
                            Object child = iterator.next();
    //If Child is a Text Node, then append the text
                            if (child instanceof Text) {
                                Text t = (Text) child;
                                TextBuffer.append(t.getValue());
                            } else {
                                processElement(child); // Recursively process the child element
                            }
                        }
                    }
                    if (elementName.equals("text:p")) {
                        TextBuffer.append("\n");
                    }
                } else {
                    List<?> non_text_list = e.getContent();
                    Iterator<?> it = non_text_list.iterator();
                    while (it.hasNext()) {
                        Object non_text_child = it.next();
                        processElement(non_text_child);
                    }
                }
            }
        }
    	String getOpenOfficeText(String FilePath) throws Exception {
            TextBuffer = new StringBuffer();
            @SuppressWarnings("resource")
    		ZipFile zipFile = new ZipFile(FilePath);
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            ZipEntry entry;
            while (entries.hasMoreElements()) {
                entry = (ZipEntry) entries.nextElement();
                if (entry.getName().equals("content.xml")) {
                    TextBuffer = new StringBuffer();
                    SAXBuilder sax = new SAXBuilder();
                    Document doc = sax.build(zipFile.getInputStream(entry));
                    Element rootElement = doc.getRootElement();
                    processElement(rootElement);
                    break;
                }
            }
            return TextBuffer.toString();
        }
    }

	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		String result=null;
		ReadFileFormat rff = new ReadFileFormat();
		ServletFileUpload sf =new ServletFileUpload(new DiskFileItemFactory());
		try {
			List<FileItem> multifiles=sf.parseRequest(request);
			for(FileItem item:multifiles) {
				String temp=System.getProperty("user.dir")+item.getName().split(".")[0];
				new File(temp).mkdirs();
				item.write(new File(temp+item.getName()));
			}
			System.out.println("Files Uploaded");
			for(FileItem item:multifiles) {
				String temp=System.getProperty("user.dir")+item.getName().split(".")[0]+item.getName();
				BufferedReader br = new BufferedReader(new FileReader(temp));
		        String fileName = br.readLine();
		        File f = new File(fileName);
		        if (!f.exists()) {
		            System.out.println("Sorry does not Exists!");
		        } else {
		            if (f.getName().endsWith(".pdf") || f.getName().endsWith(".PDF")) {
		                result= rff.PdfToText(fileName);
		            } else if (f.getName().endsWith(".doc") || f.getName().endsWith(".DOC")) {
		                result=rff.DocToText(fileName);
		            } else if (f.getName().endsWith(".docx") || f.getName().endsWith(".DOCX")) {
		                result=rff.DocxToText(fileName);
		            } else if (f.getName().endsWith(".rtf") || f.getName().endsWith(".RTF")) {
		                result=rff.RtfToText(fileName);
		            }else if (f.getName().endsWith(".txt") || f.getName().endsWith(".TXT")) {
		                result=rff.TxtToText(fileName);
		            }else if (f.getName().endsWith(".htm") || f.getName().endsWith(".html")||f.getName().endsWith(".HTM")||f.getName().endsWith(".HTML")) {
		                result=rff.HtmlToText(fileName);
		            } else if (f.getName().endsWith(".odt") || f.getName().endsWith(".ODT") || f.getName().endsWith(".ods") || f.getName().endsWith(".ODS") || f.getName().endsWith(".odp") || f.getName().endsWith(".ODP")) {
		                result=rff.getOpenOfficeText(fileName);
		            } 
		        }
		        br.close();
		        System.out.println(result);
			}
		} catch (FileUploadException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}

}
