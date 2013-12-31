// @Author Delvison Castillo

package gov.nasa.cassini;

import java.util.ArrayList;
import java.awt.Desktop;

public class ProcessThread extends Thread
{
	protected String textfilePath;
	protected String templatefilePath;
	protected String targetPath;
	protected String extension;
	protected boolean terminalMode;
	protected boolean shouldOpen;
	protected GUI gui;

  ProcessThread(String textfilePath, String templatefilePath, String
  targetPath, boolean terminalMode, boolean shouldOpen)
  {
    this.textfilePath = textfilePath;
    this.templatefilePath = templatefilePath;
    this.targetPath = targetPath;
    this.terminalMode = terminalMode;
    this.shouldOpen = shouldOpen;
    extension = 
		extension = templatefilePath.substring( templatefilePath.lastIndexOf('.')
		                                         ,templatefilePath.length() );
  }

  /**
  * Initiates all of the backend processes.
  */
  public void run()
  {
	  if (!terminalMode) { gui = GUI.getInstance(); }
	  if (!terminalMode) gui.disableButton();
    if (!terminalMode) gui.progressBar.setVisible(true);
    if (!terminalMode) gui.progress(); //increase progressbar
    if (terminalMode) System.out.println("\nGenerating your doc...");
	  Functions f = new Functions(terminalMode);
	  boolean is_success = f.generateDocument(textfilePath, templatefilePath, 
	    targetPath);
    if (is_success)
    {
      if (!terminalMode) gui.success();
      if (!terminalMode)  gui.enableButton();
      if (!terminalMode) gui.progressBar.setValue(100);
      if (!terminalMode) gui.progressBar.setString("Complete");
      if (terminalMode) System.out.println("\nDocument successfully generated"+
      	" at "+targetPath);
      this.openDocument(targetPath, shouldOpen);
    }
    else {
      if (!terminalMode) gui.fail();
			if (!terminalMode) System.out.println("\nDocument generation failed."); 
    }
      return;
	}

  protected void openDocument(String target, boolean shouldOpen)
  {
		if (shouldOpen)
		{
			try
			{
				if (target.contains(".")){ target = 
					target.substring( 0,target.lastIndexOf('.') ); 
				}
				System.out.println("Opening document....");
				Desktop.getDesktop().open(new java.io.File(target+extension));
				System.out.println("Done.");
			} catch (java.io.IOException e)
			{
				 e.printStackTrace();
			}
		}
	}

}

