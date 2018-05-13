package main;
/**
 * 
 */



import java.io.IOException;
import java.net.UnknownHostException;
import java.util.ArrayList;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;


/**
 * @author jakem
 *
 */
public class main {

	/**
	 * @param args
	 */
	public static final String SRC = "src";
	public static final String HREF= "href";
	public static final String TEXT = "text";
	public static final String TEXTNO = "textno";
	public static final String TITLE= "title";
	public static WriteToExcelFile writeToExcelFile;
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ReadExcelFile file = new ReadExcelFile();
		
		try {
			List<Competitor> competitors = file.getCompetitors();
			List<Product> products = file.read();
			List<Links> links = new ArrayList<Links>();
			List<Linkname> linknames = new ArrayList<Linkname>();
			int counter = 1;
			for(Competitor competitor : competitors) {
				Linkname linkname = new Linkname();
				linkname.setName(competitor.getName());
				linknames.add(linkname);
			}
			//open workbook for reading and writing
			
			
			WriteToExcelFile.writeHeader(linknames);
			
			for(Product product : products) {
				System.out.println("Product Line " + counter + " Product SKU " + product.getSku() + " Started Scraping");
				for(Competitor competitor : competitors) {
					Links link = new Links();
					link.setUrl(connection_type(product.getSku(), competitor));
					links.add(link);
				}
				if(links.size() > 0) {
				product.setComp1(links.get(0).getUrl());
				}
				if(links.size() > 1) {
				product.setComp2(links.get(1).getUrl());
				}
				if(links.size() > 2) {
				product.setComp3(links.get(2).getUrl());
				}
				if(links.size() > 3) {
				product.setComp4(links.get(3).getUrl());
				}
				if(links.size() > 4) {
				product.setComp5(links.get(4).getUrl());
				}
				if(links.size() > 5) {
				product.setComp6(links.get(5).getUrl());
				}
				if(links.size() > 6) {
				product.setComp7(links.get(6).getUrl());
				}
				if(links.size() > 7) {
				product.setComp8(links.get(7).getUrl());
				}
				if(links.size() > 8) {
				product.setComp9(links.get(8).getUrl());
				}
				if(links.size() > 9) {
				product.setComp10(links.get(9).getUrl());
				}
				if(links.size() > 10) {
				product.setComp11(links.get(10).getUrl());
				}
				if(links.size() > 11) {
				product.setComp12(links.get(11).getUrl());
				}
				if(links.size() > 12) {
				product.setComp13(links.get(12).getUrl());
				}
				if(links.size() > 13) {
				product.setComp14(links.get(13).getUrl());
				}
				if(links.size() > 14) {
				product.setComp15(links.get(14).getUrl());
				}
				if(links.size() > 15) {
				product.setComp16(links.get(15).getUrl());
				}
				if(links.size() > 16) {
				product.setComp17(links.get(16).getUrl());
				}
				links.clear();
				System.out.println("Product Line " + counter + " Product SKU " + product.getSku() + " Finished Successfully!");
				counter++;
				WriteToExcelFile.writeProduct(product);
			}
			
			System.out.println("We have finished scraping successfully!");
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	public static String getWebContentGetLoop(String sku, Competitor competitor) {
		String href = "";
		try {
			Document doc = Jsoup.connect(competitor.getUrl()+sku)
					.ignoreContentType(true)
					.userAgent("Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:25.0) Gecko/20100101 Firefox/25.0")
					.referrer("http://www.google.com")
					.followRedirects(true)
					.get();
			Elements newsHeadlines = doc.select(competitor.getSelect());
			if(!newsHeadlines.isEmpty()) {
				if(competitor.getValidation_select() != null) {
	
					Elements newsHeadlines1 = doc.select(competitor.getValidation_select());
					int count = 0;
					if(!newsHeadlines1.isEmpty()) {
						for(Element line : newsHeadlines1) {
							String id = line.text();
							if(id.toLowerCase().contains(sku.toLowerCase())) {
								href = newsHeadlines.get(count).absUrl("href");
								break;
							}
							count++;
						}
					}
				}
			}
				
		} 
		catch (UnknownHostException e) {
			e.printStackTrace();
		}
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}
		catch (NullPointerException e) {
			e.printStackTrace();
		}
		
		
		return href;
	}
	public static String getWebContentGet(String sku, Competitor competitor) {
		String href = "";
		try {
			Document doc = Jsoup.connect(competitor.getUrl()+sku)
					.ignoreContentType(true)
					.userAgent("Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:25.0) Gecko/20100101 Firefox/25.0")
					.referrer("http://www.google.com")
					.followRedirects(true)
					.get();
			Elements newsHeadlines = doc.select(competitor.getSelect());
			
			if(!newsHeadlines.isEmpty()) {
				System.out.println("link: " + newsHeadlines.get(0).absUrl("href"));
				if(validation(doc,competitor,sku,newsHeadlines)) {
					href = newsHeadlines.get(0).absUrl("href");
				}
			}
				
		} 
		catch (UnknownHostException e) {
			e.printStackTrace();
		}
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}
		catch (NullPointerException e) {
			e.printStackTrace();
		}
		
		
		return href;
	}
	public static String getWebContentPost(String sku, Competitor competitor) {
		String href = "";
		try {
			Document doc = Jsoup.connect(competitor.getUrl())
					.ignoreContentType(true)
					.userAgent("Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:25.0) Gecko/20100101 Firefox/25.0")
					.referrer("http://www.google.com")
					.followRedirects(true)
					.data("CAT_ProductSearch",sku)
					.data("submit","")
					.post();
			Elements newsHeadlines = doc.select(competitor.getSelect());
			if(!newsHeadlines.isEmpty()) {
				if(validation(doc,competitor,sku,newsHeadlines)) {
					href = newsHeadlines.get(0).absUrl("href");
				}
			}
			
		} catch (UnknownHostException e) {
			e.printStackTrace();
		}
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (IndexOutOfBoundsException e) {
			e.printStackTrace();
		}
		catch (NullPointerException e) {
			e.printStackTrace();
		}
		
		return href;
	}
	
	//method used to vaidate urls for example searching for product code where no product exists but gives similar
	
	public static Boolean validation(Document doc,Competitor competitor, String sku, Elements newsHeadlines) {
		String temp = "";
		String code = "";
		boolean check = false;
		
		switch(competitor.getValidation_type()) {
			case SRC:
				if(competitor.getValidation_select() != null) {
					Elements newsHeadlines1 = doc.select(competitor.getValidation_select());
					temp = newsHeadlines1.get(0).attr("src");
					code = sku.toLowerCase();
					if(code.contains("-a")) {
					code = code.substring(0, code.length()-2);
					}
					if (temp.toLowerCase().contains(code.toLowerCase())) {
						check = true;
					}
				}
				break;
			case HREF:
				temp = newsHeadlines.get(0).attr("href");
				code = sku.toLowerCase();
				if(code.contains("-a")) {
				code = code.substring(0, code.length()-2);
				}
				if (temp.toLowerCase().contains(code)) { 
					check = true;
					}
				break;
			case TITLE:
				temp = newsHeadlines.get(0).attr("title");
				code = sku.toLowerCase();
				if(code.contains("-a")) {
				code = code.substring(0, code.length()-2);
				}
				if (temp.toLowerCase().contains(code)) { 
					check = true;
					}
				break;
			case TEXTNO:
				temp = newsHeadlines.get(0).text();
				code = sku.toLowerCase();
				if(code.contains("-a")) {
				code = code.substring(0, code.length()-2);
				}
				if (temp.toLowerCase().contains(code)) { 
					check = true;
					}
				break;
			case TEXT:
				if(competitor.getValidation_select() != null) {
					Elements newsHeadlines2 = doc.select(competitor.getValidation_select());
					temp = newsHeadlines2.get(0).text();
					code = sku.toLowerCase();
					if(code.contains("-a")) {
					code = code.substring(0, code.length()-2);
					}
					if (temp.toLowerCase().contains(code)) {
						check = true;
					}
				}
				break;
			case "no":
				check = true;
				break;
		}
		return check;
	}
	public static String connection_type(String sku, Competitor competitor) {
		String temp = "";
		switch(competitor.getConnection_type()) {
			case "post":
				temp = getWebContentPost(sku, competitor);
				break;
			case "get":
				temp = getWebContentGet(sku, competitor);
				break;
			case "getloop":
				temp = getWebContentGetLoop(sku, competitor);
				break;
		}
		System.out.println(competitor.getName() +": " + temp);
		
		return temp;
		
	}
	
}
