/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package libreofficesandbox;

/**
 *
 * @author yehster
 */
public class DocumentAnalyzer {
    protected static java.util.regex.Pattern pHeader = java.util.regex.Pattern.compile("^[A-Z]+[/#-]?\\s??[A-Z]*:");
    protected java.util.ArrayList alHeaders = new java.util.ArrayList();
    public boolean ParseParagraph(com.sun.star.text.XTextContent xText)
    {
        String sParagraph=xText.getAnchor().getString().trim();
        //System.out.println(sParagraph);
        java.util.regex.Matcher mHeader = pHeader.matcher(sParagraph);
        if(mHeader.find())
        {
            String mHeaderText=sParagraph.substring(mHeader.start(), mHeader.end()-1);
            System.out.println(mHeaderText);
            alHeaders.add(mHeaderText);
        }
        else
        {
        
        }
        return true;
    }
    
}
