import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.mongodb.*;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.model.Filters;
import static com.mongodb.client.model.Filters.*;
import static com.mongodb.client.model.Updates.*;
import com.mongodb.client.model.UpdateOptions;
import com.mongodb.client.result.*;
import org.bson.Document;
import org.bson.types.ObjectId;

public class BiblioParser {
	
	final static int ISBNLength = 10;
	final static int ISSNLength = 8;
	final static int AuthorCell = 0;
	final static int YearCell = 1;
	final static int TitleCell = 2;
	final static int PublicationCell = 3;
	//final static int CityOfPublicationCell = 4;
	final static int PublisherCell = 4;
	final static int BibleoCell = 5;
	final static int LPCell = 6;
	final static int LRCell = 7;
	final static int CRCell = 8;
	final static int KeywordCell = 9;
	final static int ISBNCell = 10;
	final static int ISSNCell = 11;
	final static int LibraryCell = 12;
	final static int URLCell = 13;
	final static int DataOfEntryCell = 14;
	final static int SourceCell = 15;
	
	/******* create Excel sheet *******/
	static Workbook workbook = new HSSFWorkbook();
	static Sheet sheet = workbook.createSheet("test1");
	static Row heading = sheet.createRow(0);
	
	public static void main(String[] args) throws Exception {

		/******* insert cell names *******/		
		heading.createCell(0).setCellValue("Author");
		heading.createCell(1).setCellValue("Year");
		heading.createCell(2).setCellValue("Title");
		heading.createCell(3).setCellValue("Publication");
		//heading.createCell(4).setCellValue("City of publication");
		heading.createCell(4).setCellValue("Publisher");
		heading.createCell(5).setCellValue("Bibleo Name");
		heading.createCell(6).setCellValue("Language Published");
		heading.createCell(7).setCellValue("Language Researched");
		heading.createCell(8).setCellValue("Country of Research");
		heading.createCell(9).setCellValue("Keywords");
		heading.createCell(10).setCellValue("ISBN");
		heading.createCell(11).setCellValue("ISSN");
		heading.createCell(12).setCellValue("Library");
		heading.createCell(13).setCellValue("URL");
		heading.createCell(14).setCellValue("Date of Entry");
		heading.createCell(15).setCellValue("Source");
		
		try{
			
			// Connect to a single MongoDB instance
			MongoClientURI uri = new MongoClientURI(
					"mongodb://jiataow:jesse_0626X@cluster0-shard-00-00-f0faw.mongodb.net:27017,cluster0-shard-00-01-f0faw.mongodb.net:27017,cluster0-shard-00-02-f0faw.mongodb.net:27017/test?ssl=true&replicaSet=Cluster0-shard-0&authSource=admin");
			MongoClient mongoClient = new MongoClient(uri);
			
			// Access a database
			MongoDatabase database = mongoClient.getDatabase("test");
			
			// Access a collection
			//MongoCollection<Document> collection = database.getCollection("testCollection");
			MongoCollection<Document> collection = database.getCollection("biblio");
			
			List<Document> documents = new ArrayList<Document>();
			
			
			/******* localhost ******/
			//MongoClient mongoClient = new MongoClient("localhost", 27017);
			//MongoDatabase database = mongoClient.getDatabase("test");
			//MongoCollection<Document> collection = database.getCollection("testCollection");
			//List<Document> documents = new ArrayList<Document>();
			
			/******* Read Word document *******/
			String fileName = "C:/Users/jiatao/Downloads/MasterBibLeo_170726.docx";
			//String fileName = "C:/Users/jiatao/Downloads/test1-5.docx";
			//String fileName = "C:/Users/jiatao/Downloads/test6-10.docx";
			XWPFDocument docx = new XWPFDocument(new FileInputStream(fileName));
			
			/******* Extract text *******/
			XWPFWordExtractor we = new XWPFWordExtractor(docx);
			String test = we.getText();
			//System.out.println(test);
			
			/******* Split text into separate paragraphs *******/
			ArrayList<String> paras = new ArrayList<String>();
			StringBuilder temp = new StringBuilder();
			Scanner scanner = new Scanner(test);
			while(scanner.hasNextLine()){
				String ch = scanner.nextLine();
				if(ch.equals("")){
					// if current line is empty
					paras.add(temp.toString());	// add paragraph into list
					temp.delete(0, temp.length());	// clean the string
				}
				else
					temp.append(ch+"\r");	// expend the string
			}
			paras.add(temp.toString());	// add the last paragraph
			temp.delete(0, temp.length());
			
			System.out.println("****************** " + paras.size()+" ******************\r");
			
			/******* Parse paragraphs *******/
			for(int i = 0; i < paras.size(); i++){
				System.out.println("---------------- " + (i+1) + "th para --------------");
				System.out.println(paras.get(i));
				System.out.println();
				Bibliography biblio = new Bibliography();
				biblio = parsePara(paras.get(i),i);
				if(checkFormat(biblio)){
					/******* Output into Excel *******/
					Row row = sheet.createRow(i+1);
					row.createCell(AuthorCell).setCellValue(biblio.author);
					row.createCell(YearCell).setCellValue(biblio.year);
					row.createCell(TitleCell).setCellValue(biblio.title);
					row.createCell(PublicationCell).setCellValue(biblio.publication);
					row.createCell(PublisherCell).setCellValue(biblio.publisher);
					row.createCell(LPCell).setCellValue(biblio.LP);
					row.createCell(LRCell).setCellValue(biblio.LR);
					row.createCell(CRCell).setCellValue(biblio.CR);
					row.createCell(KeywordCell).setCellValue(biblio.KW);
					row.createCell(ISBNCell).setCellValue(biblio.ISBN);
					row.createCell(ISSNCell).setCellValue(biblio.ISSN);
					row.createCell(BibleoCell).setCellValue(biblio.bibleo);
					row.createCell(URLCell).setCellValue(biblio.URL);
					row.createCell(DataOfEntryCell).setCellValue(biblio.dateEntry);
					row.createCell(SourceCell).setCellValue(biblio.source);
					
					/******* Create&Add the document into the document list *******/
					Document doc = new Document();
					doc = createDoc(biblio);
					documents.add(doc);
					
				}
				else{
					Row row = sheet.createRow(i+1);
					row.createCell(AuthorCell).setCellValue("ERROR");
				}
				
				//System.out.println("\r-------------------------------------------------------------------------------------------\r");
			
			}
			
			/******* Insert the document list into database *******/
			collection.insertMany(documents);
			System.out.println("collection count: " + collection.count());
			
			/******* Output Excel file *******/
			FileOutputStream output = new FileOutputStream("test2.xls");
			workbook.write(output);
			output.close();
			
		}catch(Exception e){
			System.out.println(e);
		}
		
}

	private static Document createDoc(Bibliography biblio) {
		
		Bibliography bib = biblio;
		/*
		Document doc = new Document("Author", bib.author)
						.append("Year", bib.year)
						.append("Title", bib.title)
						.append("Publication", bib.publication)
						.append("Publisher", bib.publisher)
						.append("Bibleo Name", bib.bibleo)
						.append("Language Published", bib.LP)
						.append("Language Reaserched", bib.LR)
						.append("Country of Research", bib.CR)
						.append("Keywords", bib.KW)
						.append("ISBN", bib.ISBN)
						.append("ISSN", bib.ISSN)
						.append("URL", bib.URL)
						.append("Date of Entry", bib.dateEntry);*/
		
		Document doc = new Document("author", bib.author)
				.append("year", bib.year)
				.append("title", bib.title)
				.append("publication", bib.publication)
				.append("publisher", bib.publisher)
				.append("biblio_name", bib.bibleo)
				.append("language_published", bib.LP)
				.append("language_researched", bib.LR)
				.append("country_of_research", bib.CR)
				.append("keywords", bib.KW)
				.append("isbn", bib.ISBN)
				.append("issn", bib.ISSN)
				.append("url", bib.URL)
				.append("date_of_entry", bib.dateEntry)
				.append("source", bib.source);
		
		return doc;
	}

	private static boolean checkFormat(Bibliography biblio) {
		
		boolean validBib = true;
		
		if(biblio.year.equals(""))
			validBib = false;
		
		if(biblio.title.equals(""))
			validBib = false;
		
		char first_char = ' ';
		if (biblio.publication.equals(""))
			validBib = false;
		else
			first_char = biblio.publication.charAt(0);

		if(first_char == ' '
			|| first_char == '['
			|| first_char == '('
			|| first_char == '-'
			|| first_char == ';'
			|| first_char == ','
			|| first_char == ':'
			|| first_char == '/'
			|| first_char == ')'
			|| first_char == '\\'
			|| first_char == '"'
			|| first_char == '\''
			|| first_char == '0'
			|| first_char == '1'
			|| first_char == '2'
			|| first_char == '3'
			|| first_char == '4'
			|| first_char == '5'
			|| first_char == '6'
			|| first_char == '7'
			|| first_char == '8'
			|| first_char == '9'
			|| first_char == '.'){
			validBib = false;
		}
		
		return validBib;
	}

	private static Bibliography parsePara(String string, int sequence) throws ParseException {
		
		String para = string;
		
		Bibliography biblio = new Bibliography();
		
		String author = "";
		String year = "";
		String title = "";
		String publication = "";
		String publisher = "";
		List<String> LP = new ArrayList<String>();
		List<String> LR = new ArrayList<String>();
		List<String> CR = new ArrayList<String>();
		List<String> KW = new ArrayList<String>();
		String ISBN = "";
		String ISSN = "";
		String bibleo = "";
		String URL = "";
		String dateEntry = "";
		String source = para;
		
		
		/******* Extract Bibleo *******/
		if(para.toLowerCase().contains("bibleo")){
			int start = para.toLowerCase().indexOf("bibleo:");
			int end = start + 1;
			if(para.contains("[LP") && para.indexOf("[LP") > end)
				end = para.indexOf("[LP");
			else{
				while(para.charAt(end) != '\r')
					end++;
			}
				
			bibleo = para.substring(start+8, end);
			System.out.println("BibLeo: " + bibleo);
		}
		
		String[] lines = para.split("\\r");
		int lineNum = 0;
		
		while(lineNum < lines.length) {
			
			String current_line = lines[lineNum];
			
			/******* Extract date *******/
			Pattern pDate = Pattern.compile("([0-9]{2})[\\/]([0-9]{2})[\\/]([0-9]{2,4})");
			Matcher mDate = pDate.matcher(current_line);
			if(mDate.find()) {
				for(int k = 1; k <= mDate.groupCount(); k++)
				    dateEntry = String.join("/", mDate.group());
				System.out.println("Date of entry: " + dateEntry);
			}
			
			/******* Extract author, year, publication *******/
			if (lineNum==0) {
				
				System.out.println("1 line\r" + current_line);
				String[] line_tokens = current_line.replace(". /", "./").split("\\. |\\? ");
				
				if(line_tokens.length > 2){
					author = extractAuthor(line_tokens[0]);
					System.out.println("Author: " + author);
					year = line_tokens[1].trim();
					System.out.println("Year: " + year);
					title = line_tokens[2].trim();
					System.out.println("Title: " + title);
					
					//line_tokens[0] = "";
					//line_tokens[1] = "";
					//line_tokens[2] = "";
					String[] pub_tokens = Arrays.copyOfRange(line_tokens, 3, line_tokens.length);
					
					publication = String.join(".", pub_tokens);
					System.out.println("Publication: " + publication);
				}
				
				else if(line_tokens.length == 1){
					
					author = extractAuthor(line_tokens[0]);
					System.out.println("Author: " + author);
					line_tokens[0] = "";
				}
				else{
					author = extractAuthor(line_tokens[0]);
					System.out.println("Author: " + author);
					year = line_tokens[1].trim();
					System.out.println("Year: " + year);
					line_tokens[0] = "";
					line_tokens[1] = "";
				}
			
			}
			
//			/******* Extract publisher *******/
//			else if(lineNum == 1) {
//				
//				System.out.println("2 line\r" + current_line);
//				
//			}
			
			
			/******* Extract tags *******/
			else if(lineNum == lines.length-1) {
				System.out.println("last line\r" + current_line);
				LP = extractNamedTag(current_line, "LP");
				System.out.println("LP: " + LP);
				LR = extractNamedTag(current_line, "LR");
				System.out.println("LR: " + LR);
				CR = extractNamedTag(current_line, "CR");
				System.out.println("CR: " + CR);
				KW = extractUnknowTag(current_line);
				System.out.println("KW: " + KW);
			}
			
			else{
				System.out.println((lineNum+1) + " line\r" + current_line);
				if(current_line.contains("ISBN")){
					int start = current_line.toLowerCase().indexOf("isbn");
					int end = current_line.length();
					ISBN = current_line.substring(start, end).replaceAll("[ISBN]", "").replace("-", "").replace(":", "").replace(" ", "");
					System.out.println("ISBN: " + ISBN);
				}
				else if(current_line.contains("ISSN")){
					int start = current_line.toLowerCase().indexOf("issn");
					int end = current_line.length();
					ISSN = current_line.substring(start, end).replaceAll("[ISSN]", "").replace("-", "").replace(":", "").replace(" ", "");
					System.out.println("ISSN: " + ISSN);
				}
				else if(current_line.contains("http://") ||
						current_line.contains("https://")){
					int start = current_line.toLowerCase().indexOf("http");
					int end = start + 1;
					while(end < current_line.length() && !(current_line.charAt(end) == ' ' || current_line.charAt(end) == '\r')){
						end++;
					}
					URL = current_line.substring(start, end);
					System.out.println("URL: " + URL);
				}	
				else if(!current_line.toLowerCase().contains("bibleo")){
					publisher = current_line;
					System.out.println("Publisher: " + publisher);
				}
			}
			
			lineNum++;
		}
		
		biblio.author = author;
		biblio.year = year;
		biblio.title = title;
		biblio.publication = publication;
		biblio.publisher = publisher;
		biblio.LP = LP.toString().replace("[", "").replace("]", "");
		biblio.LR = LR.toString().replace("[", "").replace("]", "");
		biblio.CR = CR.toString().replace("[", "").replace("]", "");
		biblio.KW = KW.toString().replace("[", "").replace("]", "");
		biblio.ISBN = ISBN;
		biblio.ISSN = ISSN;
		biblio.bibleo = bibleo;
		biblio.URL = URL;
		biblio.dateEntry = dateEntry;
		biblio.source = source;
		
		return biblio;
	
	}

	private static List<String> extractUnknowTag(String current_line) {
		
		int i = 0;
		String tagName = "";
		String tagInfo = "";
		List<String> tagList = new ArrayList<String>();
		boolean tagFound = false;
		
		while(i < current_line.length()){
			
			char line_char = current_line.charAt(i);
			
			if(tagFound){
				if(line_char == ':'){
					tagInfo = tagInfo.trim();
					
					if(tagInfo.equals("LR")
						|| tagInfo.equals("LP")
						|| tagInfo.equals("CR")){
						tagName = "";
						tagInfo = "";
						tagFound = false;
					}
					else tagInfo += line_char;
				}
				
				else if(line_char == ']'){
					tagInfo = tagInfo.trim();
					tagList.add(tagInfo);
					tagInfo = "";
					tagFound = false;
				}
				else
					tagInfo += line_char;	// expend the information string
			}
			
			else if(line_char == '['){
				tagName = ""; // begin of tag, initialize to store the tag name
				tagFound = true;
			}
			
			else
				tagName += line_char;	// expend the information string
			
			i++;
		}
	
		return tagList;
		
	}

	private static List<String> extractNamedTag(String current_line, String str) {
		
		
		int i = 0;
		String tagName = "";
		String tagInfo = "";
		List<String> tagList = new ArrayList<String>();
		boolean tagFound = false;
		
		while(i < current_line.length()){
			
			char line_char = current_line.charAt(i);
			
			if(tagFound){
				if(line_char == ':')
					tagInfo = "";	// initialize to store tag information
				else if(line_char == ']'){
					tagInfo = tagInfo.trim();
					tagList.add(tagInfo);
					tagInfo = "";	// clear the storage
					tagFound = false;
				}
				else
					tagInfo += line_char;	// expend the information string
				
			}
			
			else if(line_char == '['){
				tagName = ""; // begin of tag, initialize tp store the tag name
			}
			
			else if(line_char == ':' ||
					line_char == ' '){
				tagName = tagName.trim();
				if(tagName.equals(str)){	// compare with the read in tag name
					tagName = "";
					tagFound = true;
				}
			}
			
			else
				tagName += line_char;	// expend the information string
			
			i++;
		}
	
		return tagList;
	}

	private static String extractAuthor(String str) {

		String[] authorList = str.split("\\/");
		String authors = String.join(",", authorList);
		
		return authors;
	}
		

}
