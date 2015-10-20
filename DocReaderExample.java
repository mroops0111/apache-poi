import java.io.*;
import java.util.Arrays;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hwpf.usermodel.*;

public class DocReaderExample{
	private File file;
	
	public static void readDocFile(String filename){
		try{
			//create document file
			InputStream fis = new FileInputStream(filename);
			POIFSFileSystem fs = new POIFSFileSystem(fis);
			HWPFDocument doc = new HWPFDocument(fs);
			
			//get the content in the document
			Range range = doc.getRange();
			System.out.println("paragraph numbers: " +range.numParagraphs());
			
			for (int i = 0; i < range.numParagraphs(); i++) {
				Paragraph par = range.getParagraph(i);
				if(par.isInTable()){
					i = readTable(range, i);
				}else{
					System.out.println(par.text());
				}
			}

		}catch(Exception e){
			System.out.println("Exception: " + e);
		}		
		
	}
	
	public static int readTable(Range range, int parIdx){
		Paragraph par = range.getParagraph(parIdx);
		System.out.println("====================Table====================");
		
		int numRow = 0;
		boolean inRow = true;
		do{
			//check the position of the paragraph in the table
			if(inRow){
				System.out.println("Row " + (++numRow));
				inRow = false;
			}
			if(par.isTableRowEnd()){
				inRow = true;
			}
			if(par instanceof ListEntry)	//check the list in the table
				parIdx = readList(range, parIdx);
			else
				System.out.println(par.text()); 
			
			//check and get the next paragraph
			parIdx++;
			if(parIdx < range.numParagraphs())
				par = range.getParagraph(parIdx);
			else
				break;
		}while(par.isInTable());
		System.out.println("==================Table End==================");
		
		return parIdx-1;
	}
	
	public static int readList(Range range, int parIdx){
		Paragraph par = range.getParagraph(parIdx);
		int curLvl = 0;
		int[] lvlIdx = {0}; //record the number of the level in the list
		System.out.println("--------------------List---------------------");
		
		do{
			int lvl = par.getIlvl();
			if(lvl < curLvl) //check the sub-level is finished or not, and reset
				for(int i=lvl+1; i<=curLvl; i++)
					lvlIdx[i] = 0;
			else if(lvl < lvlIdx.length) //adjust 
				lvlIdx = Arrays.copyOf(lvlIdx, lvlIdx.length+1);
			lvlIdx[lvl]++;
			curLvl = lvl;
			
			//print
			for(int i=0; i<=lvl; i++)
				System.out.print((i==0?"":".") + lvlIdx[i]);					
			System.out.println("\t" + par.text());
			
			//check and get the next paragraph
			parIdx++;
			if(parIdx < range.numParagraphs())
				par = range.getParagraph(parIdx);
			else
				break;
			
		}while(par instanceof ListEntry);
		System.out.println("------------------List End-------------------");
		
		return parIdx-1;
	}
	
	public static void main(String args[]){
		String filename = "DocExample.doc";
		readDocFile(filename);
	}	
	
}