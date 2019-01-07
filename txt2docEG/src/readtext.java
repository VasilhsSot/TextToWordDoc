import java.io.*;
import java.io.IOException;
import java.io.PrintWriter;
import word.api.interfaces.*;
import word.utils.*;
import word.w2004.Document2004;
import word.w2004.Document2004.Encoding;
import word.w2004.Footer2004;
import word.w2004.Header2004;
import word.w2004.Body2004;
import word.w2004.style.*;
import word.w2004.elements.*;
import com.thoughtworks.xstream.*;
import com.thoughtworks.xstream.alias.*;
import com.thoughtworks.xstream.annotations.*;
import com.thoughtworks.xstream.converters.*;
import com.thoughtworks.xstream.converters.basic.*;
import com.thoughtworks.xstream.converters.collections.*;
import com.thoughtworks.xstream.converters.enums.*;
import com.thoughtworks.xstream.converters.extended.*;
import com.thoughtworks.xstream.converters.javabean.*;
import com.thoughtworks.xstream.converters.reflection.*;
import com.thoughtworks.xstream.core.*;
import com.thoughtworks.xstream.core.util.*;
import com.thoughtworks.xstream.io.*;
import com.thoughtworks.xstream.io.binary.*;
import com.thoughtworks.xstream.io.copy.*;
import com.thoughtworks.xstream.io.json.*;
import com.thoughtworks.xstream.io.path.*;
import com.thoughtworks.xstream.io.xml.*;
import com.thoughtworks.xstream.io.xml.xppdom.*;
import com.thoughtworks.xstream.mapper.*;
import com.thoughtworks.xstream.persistence.*;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import word.w2004.style.HeadingStyle.Align;



public class readtext {
   static boolean bold=false,justify=false,leftjustify=false;
   static boolean center=false,fsize=false,ftype=false;
   static String font,size;
   static String s="adeio";
   static final String nullline="null line";
   static IDocument myDoc = new Document2004();
   
public static String whatfont (String str) {
    return str.substring(6,str.length()-1);
}
   
public static int extractInt(String str) {
        Matcher matcher = Pattern.compile("\\d+").matcher(str);

        if (!matcher.find())
            throw new NumberFormatException("For input string [" + str + "]");

        return Integer.parseInt(matcher.group());
    }

public static String findsize (String str) {    
return str.substring(4,str.length()-1);
}
   
   
public static void whatisit (String str) { 
       if (str.contains("[b]")) {
       bold=true;
       s=nullline;
   }  else if (str.contains("[/b]")) {
       bold=false;
       s=nullline;
   }  else if (str.contains("[j]")) {
       justify=true;
       s=nullline;
   }  else if (str.contains("[/j]")) {
       justify=false; 
       s=nullline;
   }  else if (str.contains("[l]")) {
       leftjustify=true;
       s=nullline;
   }  else if (str.contains("[/l]")) {
       leftjustify=false;
       s=nullline;
   }  else if (str.contains("[c]")) {
       center=true;
       s=nullline;
   }  else if (str.contains("[/c]")) {
       center=false;
       s=nullline;
   }  else if (str.contains("[p]")) {myDoc.addEle(BreakLine.times(1).create()); s=nullline;
   }  else if (str.contains("[s=")) {fsize=true; s=nullline;
   }  else if (str.contains("[font=")) {ftype=true; s=nullline;}
   else  s=str;
}
   
public static Font definefont (String str, boolean b) {
    Font fonttype;
if (!b) {    
    switch(str) {
        case "courier": fonttype = Font.COURIER; break;
        case "agency fb": fonttype = Font.AGENCY_FB; break;
        case "algerian": fonttype=Font.ALGERIAN; break;
        case "arial narrow": fonttype=Font.ARIAL_NARROW; break;
        case "baskerville old face": fonttype=Font.BASKERVILLE_OLD_FACE; break;
        case "bauhaus 93": fonttype=Font.BAUHAUS_93; break;
        case "bell mt": fonttype=Font.BELL_MT; break;
        case "berlin sans fb": fonttype=Font.BERLIN_SANS_FB; break;
        case "blackadder itc": fonttype=Font.BLACKADDER_ITC; break;
        case "bodoni mt": fonttype=Font.BODONI_MT; break;  
        case "bodoni mt condensed": fonttype=Font.BODONI_MT_CONDENSED; break;
        case "book antiqua":fonttype=Font.BOOK_ANTIQUA; break;
        case "bookman old style":fonttype=Font.BOOKMAN_OLD_STYLE; break;
        case "bradley hand itc":fonttype=Font.BRADLEY_HAND_ITC; break;
        case "broadway":fonttype=Font.BROADWAY; break;    
        case "californian fb":fonttype=Font.CALIFORNIAN_FB; break;    
        case "caastellar":fonttype=Font.CASTELLAR; break;   
        case "centaur":fonttype=Font.CENTAUR; break;
        case "century gothic":fonttype=Font.CENTURY_GOTHIC; break;  
        case "century schoolbook":fonttype=Font.CENTURY_SCHOOLBOOK; break;   
        case "chiller":fonttype=Font.CHILLER; break;
        case "colonna mt":fonttype=Font.COLONNA_MT; break;
        case "cooper black":fonttype=Font.COOPER_BLACK; break;
        case "copperplate gothic":fonttype=Font.COPPERPLATE_GOTHIC_LIGHT; break;
        case "curlz mt":fonttype=Font.CURLZ_MT; break;
        case "edwardian script itc":fonttype=Font.EDWARDIAN_SCRIPT_ITC; break;
        case "elephant":fonttype=Font.ELEPHANT; break;
        case "engravers mt":fonttype=Font.ENGRAVERS_MT; break;
        case "eras itc":fonttype=Font.ERAS_LIGHT_ITC; break;
        case "felix titling":fonttype=Font.FELIX_TITLING; break; 
        case "footlight mt":fonttype=Font.FOOTLIGHT_MT_LIGHT; break;
        case "forte":fonttype=Font.FORTE; break;
        case "franklin gothic book":fonttype=Font.FRANKLIN_GOTHIC_BOOK; break;
        case "franklin gothic demi":fonttype=Font.FRANKLIN_GOTHIC_DEMI; break;
        case "franklin gothic heavy":fonttype=Font.FRANKLIN_GOTHIC_HEAVY; break;
        case "franklin gothic medium":fonttype=Font.FRANKLIN_GOTHIC_MEDIUM; break;
        case "freestyle script":fonttype=Font.FREESTYLE_SCRIPT; break;
        case "garamond":fonttype=Font.GARAMOND; break;
        case "gigi":fonttype=Font.GIGI; break;
        case "gill sans mt":fonttype=Font.GILL_SANS_MT; break;
        case "goudy old style":fonttype=Font.GOUDY_OLD_STYLE; break;    
        case "goudy stout":fonttype=Font.GOUDY_STOUT; break;
        case "haettenschweiler":fonttype=Font.HAETTENSCHWEILER; break;
        case "harrington":fonttype=Font.HARRINGTON; break;
        case "high tower text":fonttype=Font.HIGH_TOWER_TEXT; break;
        case "imprint mt shadow":fonttype=Font.IMPRINT_MT_SHADOW; break;
        case "informal roman":fonttype=Font.INFORMAL_ROMAN; break;   
        case "jokerman":fonttype=Font.JOKERMAN; break; 
        case "juice itc":fonttype=Font.JUICE_ITC; break;
        case "kristen itc":fonttype=Font.KRISTEN_ITC; break;
        case "kunstler script":fonttype=Font.KUNSTLER_SCRIPT; break;
        case "lucida bright":fonttype=Font.LUCIDA_BRIGHT; break;
        case "lucida calligraphy":fonttype=Font.LUCIDA_CALLIGRAPHY_ITALIC; break;  
        case "lucida fax":fonttype=Font.LUCIDA_FAX_REGULAR; break;
        case "lucida handwriting":fonttype=Font.LUCIDA_HANDWRITING_ITALIC; break;
        case "lucida sans":fonttype=Font.LUCIDA_SANS_REGULAR; break; 
        case "lucida sans typewriter oblique":fonttype=Font.LUCIDA_SANS_TYPEWRITER_OBLIQUE; break;
        case "lucida sans typewriter regular":fonttype=Font.LUCIDA_SANS_TYPEWRITER_REGULAR; break;
        case "magneto":fonttype=Font.MAGNETO_BOLD; break;   
        case "maiandra gd":fonttype=Font.MAIANDRA_GD; break;  
        case "matura mt script capitals":fonttype=Font.MATURA_MT_SCRIPT_CAPITALS; break;
        case "mistral":fonttype=Font.MISTRAL; break;   
        case "modern no 20":fonttype=Font.MODERN_NO_20; break;    
        case "monotype corsiva":fonttype=Font.MONOTYPE_CORSIVA; break;    
        case "niagara solid":fonttype=Font.NIAGARA_SOLID; break; 
        case "niagara engraved":fonttype=Font.NIAGARA_ENGRAVED; break;
        case "ocr a extended":fonttype=Font.OCR_A_EXTENDED; break; 
        case "old english text mt":fonttype=Font.OLD_ENGLISH_TEXT_MT; break;     
        case "onyx":fonttype=Font.ONYX; break; 
        case "palace script mt":fonttype=Font.PALACE_SCRIPT_MT; break;     
        case "palatino linotype":fonttype=Font.PALATINO_LINOTYPE; break;     
        case "papyrus":fonttype=Font.PAPYRUS; break;    
        case "parchment":fonttype=Font.PARCHMENT; break;
        case "perpetua":fonttype=Font.PERPETUA; break;    
        case "perpetua titling mt":fonttype=Font.PERPETUA_TITLING_MT_LIGHT; break;   
        case "playbill":fonttype=Font.PLAYBILL; break;
        case "poor richard":fonttype=Font.POOR_RICHARD; break;
        case "pristina":fonttype=Font.PRISTINA; break;    
        case "rage":fonttype=Font.RAGE_ITALIC; break;    
        case "rage italic":fonttype=Font.RAGE_ITALIC; break;
        case "ravie":fonttype=Font.RAVIE; break;    
        case "rockwell":fonttype=Font.ROCKWELL; break;
        case "rockwell condensed":fonttype=Font.ROCKWELL_CONDENSED; break;    
        case "script mt bold":fonttype=Font.SCRIPT_MT_BOLD; break;   
        case "script mt":fonttype=Font.SCRIPT_MT_BOLD; break;
        case "showcard gothic":fonttype=Font.SHOWCARD_GOTHIC; break;   
        case "snap itc":fonttype=Font.SNAP_ITC; break;
        case "stencil":fonttype=Font.STENCIL; break;    
        case "tempus sans itc":fonttype=Font.TEMPUS_SANS_ITC; break;
        case "tw cen mt":fonttype=Font.TW_CEN_MT; break;    
        case "tw cen mt condensed":fonttype=Font.TW_CEN_MT_CONDENSED; break;     
        case "viner hand itc":fonttype=Font.VINER_HAND_ITC; break;
        case "vivaldi italic":fonttype=Font.VIVALDI_ITALIC; break;    
        case "vivaldi":fonttype=Font.VIVALDI_ITALIC; break;    
        case "vladimir script":fonttype=Font.VLADIMIR_SCRIPT; break;
        case "wide latin":fonttype=Font.WIDE_LATIN; break;    
        case "wingdings 2":fonttype=Font.WINGDINGS_2; break;
        case "wingdings 3":fonttype=Font.WINGDINGS_3; break;    
        case "cambria":fonttype=Font.CAMBRIA; break;    
        case "times new roman":fonttype=Font.TIMES_NEW_ROMAN; break;
        case "times new roman italic":fonttype=Font.TIMES_NEW_ROMAN_ITALIC; break;    
        case "calibri":fonttype=Font.CALIBRI; break;    
        case "calibri italic":fonttype=Font.CALIBRI_ITALIC; break; 
            
        default: fonttype=Font.COURIER; System.out.println ("Can't recognize the font type you request. So it is put to default <courier> ");
    }            
} else { 
    switch (str) {
        case "courier": fonttype = Font.COURIER_BOLD; break;
        case "agency fb": fonttype=Font.AGENCY_FB_BOLD; break; 
        case "arial narrow": fonttype=Font.ARIAL_NARROW_BOLD; break;
        case "bell mt": fonttype=Font.BELL_MT_BOLD; break; 
        case "berlin sans fb": fonttype=Font.BERLIN_SANS_FB_BOLD; break;
        case "bodoni mt": fonttype=Font.BODONI_MT_BOLD; break;  
        case "bodoni mt condensed": fonttype=Font.BODONI_MT_CONDENSED_BOLD; break;   
        case "book antiqua":fonttype=Font.BOOK_ANTIQUA_BOLD; break; 
        case "bookman old style":fonttype=Font.BOOKMAN_OLD_STYLE_BOLD; break;    
        case "britannic":fonttype=Font.BRITANNIC_BOLD; break;    
        case "californian fb":fonttype=Font.CALIFORNIAN_FB_BOLD; break; 
        case "century gothic":fonttype=Font.CENTURY_GOTHIC_BOLD; break;    
        case "century schoolbook":fonttype=Font.CENTURY_SCHOOLBOOK_BOLD; break; 
        case "copperplate gothic":fonttype=Font.COPPERPLATE_GOTHIC_BOLD; break;    
        case "eras itc":fonttype=Font.ERAS_BOLD_ITC; break;
        case "garamond":fonttype=Font.GARAMOND_BOLD; break;
        case "gill sans mt":fonttype=Font.GILL_SANS_MT_BOLD; break;    
        case "goudy old style":fonttype=Font.GOUDY_OLD_STYLE_BOLD; break;
        case "lucida bright":fonttype=Font.LUCIDA_BRIGHT_DEMIBOLD; break;
        case "lucida fax":fonttype=Font.LUCIDA_FAX_DEMIBOLD; break;     
        case "lucida sans":fonttype=Font.LUCIDA_SANS_DEMIBOLD_ROMAN; break;
        case "lucida sans typewriter oblique":fonttype=Font.LUCIDA_SANS_TYPEWRITER_BOLD_OBLIQUE; break;    
        case "magneto":fonttype=Font.MAGNETO_BOLD; break; 
        case "palatino linotype":fonttype=Font.PALATINO_LINOTYPE_BOLD; break;    
        case "perpetua":fonttype=Font.PERPETUA_BOLD; break; 
        case "perpetua titling mt":fonttype=Font.PERPETUA_TITLING_MT_BOLD; break;     
        case "rockwell":fonttype=Font.ROCKWELL_BOLD; break;
        case "rockwell condensed":fonttype=Font.ROCKWELL_CONDENSED_BOLD; break;    
        case "script mt bold":fonttype=Font.SCRIPT_MT_BOLD; break;    
        case "script mt":fonttype=Font.SCRIPT_MT_BOLD; break; 
        case "tw cen mt":fonttype=Font.TW_CEN_MT_BOLD; break; 
        case "tw cen mt condensed":fonttype=Font.TW_CEN_MT_CONDENSED_BOLD; break;    
        case "cambria":fonttype=Font.CAMBRIA_BOLD; break;
        case "times new roman":fonttype=Font.TIMES_NEW_ROMAN_BOLD; break;    
        case "times new roman italic":fonttype=Font.TIMES_NEW_ROMAN_BOLD_ITALIC; break;    
        case "calibri":fonttype=Font.CALIBRI_BOLD; break;
        case "calibri italic":fonttype=Font.CALIBRI_BOLD_ITALIC; break;    
                   
        default: fonttype=Font.COURIER_BOLD; System.out.println ("Can't recognize the font type you request or don't support this type on bold. So it is put to default <courier bold>");     
    }
    
}
return fonttype;
}          

public static void main (String args[]) throws IOException {
 
  
//doc creation

myDoc.encoding(Encoding.UTF_8);

BufferedReader br = new BufferedReader(new FileReader ("C:/Users/Public/filein.txt"));
Font type;
String line;

while ((line = br.readLine()) != null) {
      
     whatisit(line); 
     
       if (fsize) { size=findsize(line); fsize=false;} 
       else if (ftype) { font=whatfont(line); ftype=false;}
       else if (s!=nullline) {
            type=definefont(font,bold);
                if (center){       
                myDoc.addEle(Heading3.with(s).withStyle().align(Align.CENTER).create());}
                else if (leftjustify || justify){ 
                myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with(s).withStyle().font(type).fontSize(size).create()).create());
                
                }
        }
    
    }
br.close();

String myWord = myDoc.getContent();
//file extract
File fileObj = new File("C:/Users/Public/fileout.doc");
PrintWriter writer = null;
try {
    writer = new PrintWriter(fileObj);
} catch (FileNotFoundException e) {
    e.printStackTrace();
}


writer.println(myWord);
writer.close();
}
}
