// @Author Delvison Castillo

package gov.nasa.cassini;

import javax.swing.SwingUtilities;
import javax.swing.UIManager.*;
import javax.swing.UIManager;
import org.docx4j.openpackaging.exceptions.Docx4JException;

public class Main{
	public static void main(final String[] args)
	{
		// TURN OFF LOG4j
  	org.docx4j.Docx4jProperties.getProperties().setProperty(
  	    "docx4j.Log4j.Configurator.disabled", "true");
  	org.docx4j.utils.Log4jConfigurator.configure();
  	// INITIALIZE DOCx4j
		javax.xml.bind.JAXBContext c = org.docx4j.jaxb.Context.jc;
	
		SwingUtilities.invokeLater(new Runnable()
		{
			public void run()
			{
				try {
          for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels())
          {
            if ("Nimbus".equals(info.getName()))
            {
              UIManager.setLookAndFeel(info.getClassName());
              break;
            }
          }
        } catch (Exception e) {
        	try {
	          UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
	        } catch (Exception ex){

	        }
        }
				final String textfile;
				final String templatefile;
				final String savefile;
				boolean shouldOpen = false;
				boolean fourArgs = false;

        if (args.length == 4)
				{
					fourArgs = true;
				  if (args[3].equals("true")) shouldOpen = true;
				  if (args[3].equals("false")) shouldOpen = false;
				}

        if (args.length >= 3)
				{
					//TERMINAL MODE
					System.out.println("\nWelcome to Autodoc...");
					textfile = args[0];
					templatefile = args[1];
					savefile = args[2];
          final ProcessThread pt;

				  // IF AUTOOPEN OPTION IS INPUT
				  if (fourArgs){
            pt = new ProcessThread(textfile,templatefile,savefile,true,
                                 shouldOpen);
					}else{
						pt = new ProcessThread(textfile,templatefile, savefile,true,false);
					}
				  pt.run();
				}else{
						// GUI MODE
            GUI.getInstance();
			    }
			}
	  });
	}
}
