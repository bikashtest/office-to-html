package cc.mask.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.UnsupportedEncodingException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.AbstractWordConverter;
import org.apache.poi.hwpf.converter.HtmlDocumentFacade;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Style extends WordToHtmlConverter {
	
	
	public static void main(String[] args) {
		/*String docPath = "c:\\poi\\test.doc";
		String output = "c:\\aaa.html";
		Style.convertToHtml(docPath, output);*/
	}
	
	public static void convertToHtml(String docPath, String output){
		 System.out.println( "Converting " + docPath );
	        System.out.println( "Saving output to " + output );
	        try
	        {
	            Document doc = Style.process( new File( docPath ) );
	            
	            
	            FileWriter out = new FileWriter( output );
	            DOMSource domSource = new DOMSource( doc );
	            
	            StreamResult streamResult = new StreamResult( out );

	            TransformerFactory tf = TransformerFactory.newInstance();
	            Transformer serializer = tf.newTransformer();
	            // TODO set encoding from a command argument
	            serializer.setOutputProperty( OutputKeys.ENCODING, "UTF-8" );
	            serializer.setOutputProperty( OutputKeys.INDENT, "yes" );
	            serializer.setOutputProperty( OutputKeys.METHOD, "html" );
	            serializer.transform( domSource, streamResult );
	            out.close();
	        }
	        catch ( Exception e )
	        {
	            e.printStackTrace();
	        }
	}
	
	
	private static Document process(File file) {
		 HWPFDocument wordDocument = null;
//		 ori type is WordToHtmlConverter @mc
		 Style wordConverter = null;
		 
		try {
//			wordDocument = WordToHtmlUtils.loadDoc( file );
			wordDocument = new HWPFDocument(new FileInputStream(file));
			wordConverter = new Style(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument() );
			wordConverter.processDocument( wordDocument );
	        return wordConverter.getDocument();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
    public void processDocument( HWPFDocument wordDocument )
    {
        try
        {
            final SummaryInformation summaryInformation = wordDocument.getSummaryInformation();
            if ( summaryInformation != null )
            {
                processDocumentInformation( summaryInformation );
            }
        }
        catch ( Exception exc )
        {
            logger.log( POILogger.WARN, "Unable to process document summary information: ", exc, exc );
        }

        final Range docRange = wordDocument.getRange();

        if ( docRange.numSections() == 1 )
        {
            processSingleSection( wordDocument, docRange.getSection( 0 ) );
            afterProcess();
            return;
        }

        processDocumentPart( wordDocument, docRange );
        afterProcess();
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
/////---------------------------- constractor s  && logger-------------------------------
    
    private static final POILogger logger = POILogFactory.getLogger( Style.class );
    
	public Style(HtmlDocumentFacade htmlDocumentFacade) {
		super(htmlDocumentFacade);
	}
	
	public Style(Document document) {
		super(document);
	}
	
}
