/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package libreofficesandbox;

import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author yehster
 */


public class LibreOfficeSandbox {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        com.sun.star.frame.XDesktop xDesktop = null;
        xDesktop = getDesktop();
        com.sun.star.text.XTextDocument xTextDocument =OpenTextdocument(xDesktop,"e:/MedVocab/MSA01092012002/MSA01092012004.DOC");
        com.sun.star.text.XText xText = (com.sun.star.text.XText)xTextDocument.getText();
        

            com.sun.star.container.XEnumeration xParagraphEnumeration = null;
            com.sun.star.container.XEnumerationAccess xParaEnumerationAccess = null;
            com.sun.star.container.XEnumeration xPortionEnumeration = null;
            com.sun.star.container.XEnumeration xTextPortionEnum;
            com.sun.star.text.XTextContent xTextElement = null;
            
            System.out.println("create an enumeration of all paragraphs");
            // create an enumeration access of all paragraphs of a document
            com.sun.star.container.XEnumerationAccess xEnumerationAccess =
                (com.sun.star.container.XEnumerationAccess)
                    UnoRuntime.queryInterface(
                        com.sun.star.container.XEnumerationAccess.class, xText);
            xParagraphEnumeration = xEnumerationAccess.createEnumeration();

            java.util.regex.Pattern p = java.util.regex.Pattern.compile("^[A-Z]+[/#-]?\\s??[A-Z]*:");
            try
            {
            DocumentAnalyzer docAnalyzer = new libreofficesandbox.DocumentAnalyzer();
            while ( xParagraphEnumeration.hasMoreElements() ) {
                xTextElement = (com.sun.star.text.XTextContent)
                    UnoRuntime.queryInterface(
                        com.sun.star.text.XTextContent.class,
                        xParagraphEnumeration.nextElement());
                    
                docAnalyzer.ParseParagraph(xTextElement);

                }
            }
                
            // Loop through all paragraphs of the document
            catch( Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }
//        xTextDocument.dispose();
//        xDesktop.terminate();
    }
    
    public static com.sun.star.frame.XDesktop getDesktop() {
        com.sun.star.frame.XDesktop xDesktop = null;
        com.sun.star.lang.XMultiComponentFactory xMCF = null;

        try {
            com.sun.star.uno.XComponentContext xContext = null;

            // get the remote office component context
            xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();

            // get the remote office service manager
            xMCF = xContext.getServiceManager();
            if( xMCF != null ) {
                System.out.println("Connected to a running office ...");

                Object oDesktop = xMCF.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", xContext);
                xDesktop = (com.sun.star.frame.XDesktop) UnoRuntime.queryInterface(
                    com.sun.star.frame.XDesktop.class, oDesktop);
            }
            else
                System.out.println( "Can't create a desktop. No connection, no remote office servicemanager available!" );
        }
        catch( Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }


        return xDesktop;
    }
    
    protected static com.sun.star.lang.XComponent OpenDocument(
        com.sun.star.frame.XDesktop xDesktop,
        String sFilename )
    {
        String sURL="file:///"+sFilename;
        com.sun.star.lang.XComponent xComponent = null;
        com.sun.star.frame.XComponentLoader xComponentLoader = null;
        com.sun.star.beans.PropertyValue xValues[] =
            new com.sun.star.beans.PropertyValue[1];
        com.sun.star.beans.PropertyValue xEmptyArgs[] =
            new com.sun.star.beans.PropertyValue[0];
        
        try {
            xComponentLoader = (com.sun.star.frame.XComponentLoader)
                UnoRuntime.queryInterface(
                    com.sun.star.frame.XComponentLoader.class, xDesktop);
        
            xComponent  = xComponentLoader.loadComponentFromURL(
                sURL, "_blank", 0, xEmptyArgs);
        }
        catch( Exception e) {
            e.printStackTrace(System.err);
        }
        
        return xComponent ;
    }
    public static com.sun.star.text.XTextDocument OpenTextdocument(
        com.sun.star.frame.XDesktop xDesktop, String sFilename )
    {
        com.sun.star.text.XTextDocument aTextDocument = null;

        try {
            com.sun.star.lang.XComponent xComponent = OpenDocument(xDesktop,
                                                                        sFilename);
            aTextDocument = (com.sun.star.text.XTextDocument)
                UnoRuntime.queryInterface(
                    com.sun.star.text.XTextDocument.class, xComponent);
        }
        catch( Exception e) {
            e.printStackTrace(System.err);
        }

        return aTextDocument;
    }    
}
