/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */


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
        
        System.out.println(xTextDocument.getText().getString());
        
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
