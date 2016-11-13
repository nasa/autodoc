// @Author Delvison Castillo

package gov.nasa.cassini;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.util.*;
import java.io.File;
import javax.swing.filechooser.FileFilter;
import java.io.BufferedReader;
import javax.swing.border.EmptyBorder;
import java.io.InputStreamReader;
import java.io.Reader;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;

class GUI
{
  //FRAMES
  JFrame frm;
  JFrame errorLog;
  //TEXTFIELDS
  final JTextField textField;
  final JTextField templateField;
  JComboBox fileType;
  //BUTTONS
  JButton gen;
  JButton templateBrowse;
  JButton textBrowse;
  JButton help;
  JProgressBar progressBar;
  //CHECKBOXES
  JCheckBox openCheck;
  //LABELS
  JLabel status;
  //LAYOUT
  GridBagLayout gbag = new GridBagLayout();
  GridBagConstraints gbc = new GridBagConstraints();
  File fileChosen;
  static final boolean debug = false; //FOR DEBUGGING PURPOSES
  private static GUI instance; //SINGLETON

  GUI()
  {
  	//JFrame
    frm = new JFrame("AutoDoc (Microsoft Office Document Automator) v 1.6");
    frm.setLayout(gbag);
    frm.setSize(640, 220); //give the frame a size
    frm.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    frm.setLocationRelativeTo( null ); //centers window on screen
    frm.setResizable(false); //restricts resizing

    // TEXTFILE INPUTS
    JPanel textFile = new JPanel();
    textFile.setLayout(new FlowLayout());
    textFile.setBorder(BorderFactory.createTitledBorder("Text file"));
    textField = new JTextField(15);
    textFile.add(textField);
    textBrowse = new JButton("Browse");
    textFile.add(textBrowse);
    gbc.gridx = 0;
    gbc.gridy = 0;
    gbag.setConstraints(textFile, gbc);
    frm.add(textFile,gbc);

    // TEMPLATE FILE INPUTS
    JPanel templateFile = new JPanel();
    templateFile.setLayout(new FlowLayout());
    templateFile.setBorder(BorderFactory.createTitledBorder("Template file"));
    templateField = new JTextField(15);
    templateFile.add(templateField);
    templateBrowse = new JButton("Browse");
    templateFile.add(templateBrowse);
    gbc.gridx = 1;
    gbc.gridy = 0;
    gbag.setConstraints(templateFile, gbc);
    frm.add(templateFile,gbc);

    // BUTTON PANEL
    JPanel btns = new JPanel();
    btns.setLayout(new FlowLayout());

    // HELP BUTTON
    help = new JButton("?");

    // GENERATE BUTTON
    gen = new JButton("Generate");

    // ADD HELP & GENERATE BTNS TO PANEL
    btns.add(help);
    help.setPreferredSize(new Dimension(35,20));
    btns.add(gen);
    gbc.gridx = 1;
    gbc.gridy = 1;
    frm.add(new JLabel("  "),gbc);
    gbc.anchor = GridBagConstraints.EAST;
    gbc.gridx = 1;
    gbc.gridy = 3;
    frm.add(btns,gbc);

    // PROGRESS BAR
		progressBar = new JProgressBar(0, 100);
		progressBar.setValue(0);
		progressBar.setStringPainted(true);
		gbc.anchor = GridBagConstraints.CENTER;
		gbc.gridx = 0;
		gbc.gridy = 3;
		frm.add(progressBar,gbc);
		progressBar.setPreferredSize(new Dimension(200,30));
		progressBar.setVisible(false);
		progressBar.setString("0%");

    // STATUS LABEL
		status = new JLabel(" ");
		status.setMaximumSize(new Dimension(6,6));
    status.setPreferredSize(new Dimension(200,30));
    gbc.anchor = GridBagConstraints.CENTER;
    gbc.gridx = 0;
    gbc.gridy = 4;
    frm.add(status,gbc);
    if (debug) { status.setText("DEBUG MODE");}

    // AUTO OPEN CHECKBOX
    openCheck = new JCheckBox("Auto-open", true);
    gbc.anchor = GridBagConstraints.EAST;
    gbc.gridx = 1;
    gbc.gridy = 4;
    frm.add(openCheck,gbc);

    initDragDrop();
    gbc.gridx = 0;
    gbc.gridy = 1;
    frm.setVisible(true);

    // ActionListener for Text File Field
    textBrowse.addActionListener(new ActionListener()
    {
      public void actionPerformed(ActionEvent e)
      {
        JFileChooser chooser= new JFileChooser();
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.addChoosableFileFilter(new FileFilter()
        {
          public boolean accept(File f)
          {
            // Allow directories to be seen.
            if ( f.isDirectory() )
            {
              return true;
            }
            if (f.getName().toLowerCase().endsWith(".txt"))
              return true;
            return false;
          }
          public String getDescription()
          {
            return "*.txt";
          }
        });
        int choice = chooser.showOpenDialog(frm);
        if (choice != JFileChooser.APPROVE_OPTION) return;
        textField.setText(chooser.getSelectedFile().getPath());
      }
    });

    // ACTIONLISTENER FOR TEMPLATE FILE FIELD
    templateBrowse.addActionListener(new ActionListener()
    {
      public void actionPerformed(ActionEvent e)
      {
        JFileChooser chooser= new JFileChooser();
        chooser.addChoosableFileFilter(new FileFilter()
        {
          public boolean accept(File f)
          {
            if (f.isDirectory()) {
              return true;
            }
            if (f.getName().toLowerCase().endsWith(".docx"))
              return true;
            if (f.getName().toLowerCase().endsWith(".pptx"))
              return true;
            if (f.getName().toLowerCase().endsWith(".xlsx"))
              return true;
            return false;
          }
          public String getDescription()
          {
            return "*.docx, *.pptx, *.xlsx";
          }
        });
        int choice = chooser.showOpenDialog(frm);
        if (choice != JFileChooser.APPROVE_OPTION) return;
        //chooser.setVisible(true);
        templateField.setText(chooser.getSelectedFile().getPath());
      }
    });

    // ACTIONLISTENER FOR GENERATE BUTTON
    gen.addActionListener(new ActionListener()
    {
      public void actionPerformed(ActionEvent thing)
      {
        status.setForeground(Color.BLACK);
        String templatePath="";
        String textFilePath="";
        progressBar.setValue(0);
        templatePath = templateField.getText();
        textFilePath = textField.getText();
        if (!templatePath.equals("") && !textFilePath.equals(""))
        {
          //SAVE FILECHOOSER
          JFileChooser chooser = new JFileChooser();
          chooser.setDialogTitle("Specify where to save file");
          chooser.addChoosableFileFilter(new FileFilter()
          {
            public boolean accept(File f)
            {
              if (f.isDirectory())
              {
                return true;
              }
                return false;
            }

            public String getDescription()
            {
              String templatePath = templateField.getText();
              return templatePath.substring(templatePath.lastIndexOf('.'),
                                            templatePath.length());
            }
          });

          int userSelection = chooser.showSaveDialog(frm);
          String targetPath = "";
          if (userSelection == JFileChooser.APPROVE_OPTION)
          {
            progressBar.setVisible(true);
            status.setText("Generating file...");
            File fileToSave = chooser.getSelectedFile();
            targetPath = fileToSave.getAbsolutePath();
            ProcessThread pt = new ProcessThread(textFilePath, templatePath,
                                        targetPath, false, checkOpen() );
            pt.start();
          }
        } else { setStatus("Error: Choose files."); }
      }
    });

    // HELP ACTION LISTENER
    help.addActionListener(new ActionListener()
    {
      public void actionPerformed(ActionEvent thing)
      {
        JFrame hf = new JFrame("Instructions");
        hf.setSize(300,480);
        hf.setLocationRelativeTo( frm ); //centers window on screen
        hf.setResizable(false); //restricts resizing
        JLabel ja = new JLabel( "<html>"+renderInstructions()+"</html>" );
        ja.setPreferredSize(new Dimension(90,48));
        JScrollPane sc = new JScrollPane(ja);
        sc.setPreferredSize(new Dimension(950,500));

        java.net.URL p = getClass().getResource("/jpl_logo.png" );
        JLabel about = new JLabel( "<html><center<<img src='"+p+"'></center>"+
                                  renderAbout()+"</html>" );

        //TABBED PANE
        JTabbedPane tab = new JTabbedPane();
        tab.addTab("Instructions",sc);
        tab.addTab("About",about);

        hf.add(tab);
        hf.setVisible(true);
      }
    });
  }
  // SINGLETON DESIGN
  protected static GUI getInstance()
  {
    if(instance == null) {
    	instance = new GUI();
    }
    return instance;
  }

  // DISPLAYS SUCCESS MESSAGE
  protected void success()
  {
    status.setForeground(new Color(7,85,10));
    status.setText("File successfully generated.");
  }

  // DISPLAYS FAILURE MESSAGE
  protected void fail()
  {
  	status.setForeground(new Color(7,85,10));
    status.setText("Generation was unsuccesful.");
  }

  // DISABLES BUTTONS
  protected void disableButton()
  {
  	gen.setText("Generating..");
    gen.setEnabled(false);
    textField.setEnabled(false);
    templateField.setEnabled(false);
    textBrowse.setEnabled(false);
    templateBrowse.setEnabled(false);
  }

  // ENABLES BUTTONS
  protected void enableButton()
  {
  	gen.setText("Generate");
    gen.setEnabled(true);
    textField.setEnabled(true);
    templateField.setEnabled(true);
    textBrowse.setEnabled(true);
    templateBrowse.setEnabled(true);
  }

  // RETURNS DEBUG VALUE
  protected boolean getDebug()
  {
  	return debug;
  }

  // SET STATUS
  protected void setStatus(String stat)
  {
  	this.status.setText(stat);
  }

  // INCREASE PROGRESS BAR
  protected void progress()
  {
  	int i = progressBar.getValue();
  	if (i<90){
      int j = i+2;
  	  progressBar.setValue(j);
      progressBar.setString(j+"%");
    }
  }

  // SHOW ERRORS PANE
  protected void showErrors(String[] errors){
  	errorLog = new JFrame("Error log");
    errorLog.setLayout(gbag);
    errorLog.setSize(310, 360); //give the frame a size
    errorLog.setLocationRelativeTo( null ); //centers window on screen
    gbc.anchor = GridBagConstraints.CENTER;
    gbc.gridx = 0;
    gbc.gridy = 0;

    //ERROR MESSAGE
    JPanel jp = new JPanel();
    jp.setPreferredSize(new Dimension(200,80));
    JLabel n = new JLabel("<html>These keywords were "+
      "not found in document:<html>");
    n.setMaximumSize(new Dimension(190, 50));
    n.setFont(new Font( n.getFont().getFontName(), Font.BOLD,
      n.getFont().getSize() ));
    n.setForeground(Color.RED);
    gbag.setConstraints(n, gbc);
    jp.add(n);
    errorLog.add(n);

    //Add Scrollpane
    JList lis = new JList(errors);
    JScrollPane scroll = new JScrollPane(lis);
    scroll.setPreferredSize(new Dimension(230,250));
    //lis.setPreferredSize(new Dimension(230,250));
    gbag.setConstraints(scroll, gbc);
    gbc.anchor = GridBagConstraints.CENTER;
    gbc.gridx = 0;
    gbc.gridy = 1;
    errorLog.add(scroll,gbc);

    //OK button
    JButton ok = new JButton("Ok");
    ok.registerKeyboardAction(ok.getActionForKeyStroke(
      KeyStroke.getKeyStroke(KeyEvent.VK_SPACE, 0, false)),
      KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, 0, false),
      JComponent.WHEN_FOCUSED);
    ok.registerKeyboardAction(ok.getActionForKeyStroke(
      KeyStroke.getKeyStroke(KeyEvent.VK_SPACE, 0, true)),
      KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, 0, true),
      JComponent.WHEN_FOCUSED);
    gbc.anchor = GridBagConstraints.CENTER;
    gbc.gridx = 0;
    gbc.gridy = 2;
    errorLog.add(ok,gbc);
    errorLog.getRootPane().setDefaultButton(ok);
    ok.requestFocus();
    //ok.requestFocusInWindow();
    //ok.setSelected(true);

    //OK button listener
    ok.addActionListener(new ActionListener()
    {
      public void actionPerformed(ActionEvent thing)
      {
        errorLog.dispose();
      }
    });

    //Display frame
    //errorLog.setResizable(false); //restricts resizing
    errorLog.setVisible(true);
  }

  // CHECK THE AUTO OPEN
  protected boolean checkOpen()
  {
  	return openCheck.isSelected();
  }

  // SHOW INSTRUCTIONS WINDOW
  protected String renderInstructions()
  {
  	StringBuilder contentBuilder = new StringBuilder();
    try
    {
    	Reader rd = new InputStreamReader(getClass().getResourceAsStream("/instr"+
    	  "uctions.html"));
      BufferedReader in = new BufferedReader(rd);
      String str;
      while ((str = in.readLine()) != null)
      {
        contentBuilder.append(str);
      }
      in.close();
    }catch (java.io.IOException e) {
    }
    return contentBuilder.toString();
  }

  protected String renderAbout()
  {
    StringBuilder contentBuilder = new StringBuilder();
    try
    {
      Reader rd = new InputStreamReader(getClass().getResourceAsStream(
                                        "/about.html"));
      BufferedReader in = new BufferedReader(rd);
      String str;
      while ((str = in.readLine()) != null)
      {
        contentBuilder.append(str);
      }
      in.close();
    }catch (java.io.IOException e) {
    }
    return contentBuilder.toString();
  }

  // INITIALIZE DRAG N DROP
  protected void initDragDrop(){
    //DRAG & DROP FOR TEXTFILE
    textField.setDropTarget(new DropTarget()
    {
      public synchronized void drop(DropTargetDropEvent e)
      {
        try
        {
          e.acceptDrop(DnDConstants.ACTION_COPY);
          java.util.List<java.io.File> dropfiles =
            (java.util.List<java.io.File>)e.getTransferable()
             .getTransferData( DataFlavor.javaFileListFlavor );
          for ( File file : dropfiles )
          {
            textField.setText( file.getAbsolutePath() );
          }
        } catch (Exception ex) {
          ex.printStackTrace();
        }
      }
    });

    //DRAG & DROP FOR TEMPLATEFILE
    templateField.setDropTarget(new DropTarget()
    {
      public synchronized void drop(DropTargetDropEvent e)
      {
        try
        {
          e.acceptDrop(DnDConstants.ACTION_COPY);
          java.util.List<File> dropfiles =(java.util.List<File>)e
            .getTransferable().getTransferData( DataFlavor.javaFileListFlavor );
          for ( File file : dropfiles )
          {
            templateField.setText( file.getAbsolutePath() );
          }
        } catch (Exception ex) {
          ex.printStackTrace();
        }
      }
    });
  }
}
