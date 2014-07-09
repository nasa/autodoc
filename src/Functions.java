// @Author Delvison Castillo
// Modifications by Warren Kaye, 5/6/2014

package gov.nasa.cassini;

//JAVA IMPORTS
import java.util.HashMap;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import javax.xml.bind.JAXBElement;

//DOCX4J IMPORTS
import org.docx4j.openpackaging.packages.*;
import org.docx4j.openpackaging.exceptions.*;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorkbookPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.parts.PresentationML.SlideMasterPart;
import org.docx4j.dml.CTRegularTextRun;
import org.docx4j.dml.CTTextBody;
import org.docx4j.dml.CTTextParagraph;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.GraphicData;
import org.docx4j.dml.CTTable;
import org.docx4j.dml.CTTableRow;
import org.docx4j.dml.CTTableCell;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.pptx4j.pml.Shape;
import org.pptx4j.pml.CTGraphicalObjectFrame;

public class Functions {
   private static Functions instance;
   private String extension;
   protected String targetName;
   protected boolean term;
   // FOR DOCX
   private MainDocumentPart docxDocumentPart;
   private List<Object> docxBodyContent;
   // FOR XLSX
   private WorkbookPart workbookPart;
   private SpreadsheetMLPackage xlsMLPackage;
   private List<org.xlsx4j.sml.CTRst> siElements;
   // FOR PPTX
   MainPresentationPart presentationPart;
   PresentationMLPackage pptMLPackage;
   ArrayList<SlidePart> pptSlideParts;
   List<SlideMasterPart> pptSlideMaster = null;
   // FOR ALL
   private GUI gui;
   private boolean debug;
   // FOR ERRORS
   private boolean errorFound = false;
   private ArrayList<String> errors = new ArrayList<String>();

   Functions(boolean terminalMode) {
      this.term = terminalMode;
   }

   /**
    * Prints out debugging messages
    * 
    * @param msg
    *           The debugging message
    */
   protected void debugger(String msg) {
      if (gui != null) {
         this.debug = gui.debug;
         if (this.debug) {
            System.out.println("Debug------------------");
            System.out.println(msg);
            if (!term)
               gui.setStatus(msg);
         }
      }
   }

   /**
    * Main function that gets called. Determines what file is to be generated
    * and calls the appropriate function to do so.
    * 
    * @param textFilePath
    *           the path of the text file
    * @param templatePath
    *           the path of the template file
    * @param targetName
    *           the path of the intended output file
    */
   protected boolean generateDocument(String textFilePath, String templatePath,
         String targetName) {
      if (!term)
         gui = GUI.getInstance();
      if (!term)
         gui.getDebug();
      boolean is_success = false;
      // DETERMINE THE FILETYPE BY EXTENSION
      extension = templatePath.substring(templatePath.lastIndexOf('.'),
            templatePath.length());
      // STRIP EXTENSION
      this.extensionFix(targetName);

      if (!term)
         gui.progress(); // progress scrollbar

      // .DOCX
      if (extension.equals(".docx")) {
         is_success = this.generateDOCX(textFilePath, templatePath);
      }

      // .PPTX
      if (extension.equals(".pptx")) {
         is_success = this.generatePPTX(textFilePath, templatePath); // make new class; pass in Dictionary instead of textfile
      }

      // .XLSX
      if (extension.equals(".xlsx")) {
         debugger("Excel file found");
         is_success = this.generateXLSX(textFilePath, templatePath);
      }

      // CHECK FOR ERRORS
      if (errorFound) {
         String[] er = errors.toArray(new String[errors.size()]);
         if (!term)
            gui.showErrors(er);
         if (term) {
            System.out.println("=======================================");
            System.out.println("THE FOLLOWING KEYWORDS WERE NOT FOUND:");
            for (int i = 0; i < er.length; i++)
               System.out.println(er[i]);
            System.out.println("END OF KEYWORDS NOT FOUND==============");
         }
      }

      if (!term){
         gui.progress(); // progress scrollbar
      }
      // RETURN BOOLEAN
      return is_success;
   }

   /**
    * Strips the extension of a file
    * 
    * @param t
    *           string of file path
    */
   protected void extensionFix(String t) {
      if (t.endsWith(".docx") || t.endsWith(".pptx") || t.endsWith(".xlsx")) {
         this.targetName = t.substring(0, t.lastIndexOf('.'));
         if (targetName.endsWith(".docx") || targetName.endsWith(".pptx")
               || targetName.endsWith(".xlsx")) {
            this.extensionFix(this.targetName);
         }
      } else {
         this.targetName = t;
      }
   }

   /**
    * Function that is responsible for generating .docx files.
    * 
    * @param textFilePath
    *           path for text file
    * @param templatePath
    *           path for original .docx template file
    */
   protected boolean generateDOCX(String textFilePath, String templatePath) {
      boolean is_success = false;
      try {
         // OPEN TEMPLATE FILE
         WordprocessingMLPackage templateFile = WordprocessingMLPackage
               .load(new java.io.File(templatePath));
         if (!term)
            gui.progress(); // progress scrollbar

         // GET MAIN DOCUMENT PART
         docxDocumentPart = templateFile.getMainDocumentPart();
         if (!term)
            gui.progress(); // progress scrollbar

         // GET DOCUMENT BODY
         this.docxBodyContent = docxDocumentPart.getJaxbElement().getBody()
               .getContent();
         if (!term)
            gui.progress(); // progress scrollbar

         // PARSE TEXTFILE
         if (!term)
            gui.setStatus("Parsing textfile..");
         this.parseTextFile(textFilePath);

         // SAVE THE FILE
         if (!term)
            gui.setStatus("Saving file..");
         templateFile.save(new java.io.File(targetName + ".docx"));
         if (!term)
            gui.progressBar.setString("90%");
         if (!term)
            gui.progressBar.setValue(90);

         // PROCESSING WAS SUCCESSFUL
         is_success = true;
      } catch (Docx4JException e) {
         debugger("DOCX4JERROR");
         is_success = false;
      }
      return is_success;
   }

   /**
    * Function that is responsible for generating .pptx files.
    * 
    * @param textFilePath
    *           path for text file
    * @param templatePath
    *           path for original .pptx template file
    */
   protected boolean generatePPTX(String textFilePath, String templatePath) {
      boolean is_success = false;
      try {
         // Create the wordprocessingmlpackage
         pptMLPackage = PresentationMLPackage.load(new java.io.File(
               templatePath));
         if (!term)
            gui.progress(); // progress scrollbar

         // GET PART
         debugger("presentationPart");
         presentationPart = pptMLPackage.getMainPresentationPart();
         if (!term)
            gui.progress(); // progress scrollbar

         // GET SLIDE PARTS
         debugger("slidePart");
         pptSlideMaster = new ArrayList<SlideMasterPart>();
         pptSlideParts = this.getpptSlideParts();
         // pptSlideMaster is now populated.

         if (!term)
            gui.progress(); // progress scrollbar

         // PARSE TEXTFILE
         debugger("parse textfile");
         this.parseTextFile(textFilePath);

         // SAVE FILE
         pptMLPackage.save(new java.io.File(this.targetName + ".pptx"));
         is_success = true;
      } catch (Docx4JException e) {
         debugger("DOCX4JERROR");
      }
      return is_success;
   }

   /**
    * Function that is responsible for generating .xlsx files.
    * 
    * @param textFilePath
    *           path for text file
    * @param templatePath
    *           path for original .xlsx template file
    */
   protected boolean generateXLSX(String textFilePath, String templatePath) {
      boolean is_success = false;
      try {
         // Create the wordprocessingmlpackage
         debugger(templatePath);
         debugger("Opening template file");
         xlsMLPackage = SpreadsheetMLPackage
               .load(new java.io.File(templatePath));
         if (!term)
            gui.progress(); // progress scrollbar

         // GET MAIN WORKBOOK PART
         debugger("GET WorkbookPart");
         workbookPart = xlsMLPackage.getWorkbookPart();
         if (!term)
            gui.progressBar.setString("25%");
         if (!term)
            gui.progressBar.setValue(25);
         if (!term)
            gui.progress(); // progress scrollbar

         // GET SHARED STRINGS
         debugger("GET SharedStrings");
         SharedStrings ss = workbookPart.getSharedStrings();
         org.xlsx4j.sml.CTSst cts = ss.getJaxbElement();
         siElements = cts.getSi();
         if (!term)
            gui.progressBar.setString("30%");
         if (!term)
            gui.progressBar.setValue(30);
         if (!term)
            gui.progress(); // progress scrollbar

         // PARSE TEXTFILE
         debugger("parseTextFile");
         this.parseTextFile(textFilePath);
         if (!term)
            gui.progressBar.setString("50");
         if (!term)
            gui.progressBar.setValue(50);
         if (!term)
            gui.progress(); // progress scrollbar

         // RESET VALUES FOR FORMULAS
         removeXLSXFormulaValues();

         // SAVE IT
         xlsMLPackage.save(new java.io.File(this.targetName + ".xlsx"));
         is_success = true;
      } catch (Docx4JException e) {
         debugger("Docx4JException");
      }
      return is_success;
   }

   /**
    * Parses text file. Calls appropriate replace function depending on what
    * file is being used.
    * 
    * @param textFilePath
    *           path for the text file
    */
   protected void parseTextFile(String textFilePath) {
      try {
         // READ IN THE TEXT FILE
         java.util.Scanner sc = new java.util.Scanner(new java.io.File(
               textFilePath));
         String line = "";
         while (sc.hasNext()) {
            // READ NEXT LINE
            line = sc.nextLine();
            if (!term)
               this.gui.progress(); // increase progress bar

            // PARSE LINE INTO TWO PARTS -- BOOKMARK & VALUE
            String[] temp = line.split("\\s+", 2);
            // ASSURE THAT LINE HAS CONTENT
            if (temp.length == 2) {
               String bookmark = temp[0]; // "${...}"
               String value = temp[1]; // (the text to replace bookmark with)

               // IGNORE LINES STARTING WITH # (COMMENTS)
               if (!line.startsWith("#")) {
                  debugger("oye! el bookmark es " + bookmark);
                  // LOG WHEN BOOKMARKS ARE NOT FOUND IN DOCUMENT
                  if (!replaceInit(bookmark, value)) {
                     this.errorFound = true;
                     this.errors.add(bookmark);
                  }
               } else {
                  debugger("comment found");
               }
            }
         }
      } catch (java.io.FileNotFoundException e) {
         debugger("FAIL FILENOTFOUNDEXCEPTION");
      } catch (java.io.IOException e) {
         debugger("FAIL IOEXCEPTION");
      }
   }

   /**
    * Inititalizes replace function. Considers file extension and forwards to
    * the appropriate replace function.
    * 
    * @param bookmark
    *           the name of the bookmark
    * @param value
    *           the value to be inserted in place of the bookmark
    */
   protected boolean replaceInit(String bookmark, String value) {
      boolean success = true;
      if (extension.equals(".docx")) {
         success = this.replaceForDocx(bookmark, value);
      } else if (extension.equals(".pptx")) {
         success = this.replaceForPptx(bookmark, value);
      } else if (extension.equals(".xlsx")) {
         success = this.replaceForXlsx(bookmark, value);
      }
      return success;
   }

   /**
    * Responsible for replacing text for .docx files
    * 
    * @param bookmark
    *           the name of the bookmark
    * @param value
    *           the value to be inserted in place of the bookmark
    */
   protected boolean replaceForDocx(final String bookmark, final String value) {
      boolean found = false;
      boolean f;
      // REPLACE BOOKMARKS IN HEADER FOOTER
      List<JaxbXmlPart> headFoots = getDocxHeaderFooterParts();
      for (JaxbXmlPart<? extends ContentAccessor> jx : headFoots) {
         List<Object> contents = jx.getJaxbElement().getContent();
         f = changeDocxParagraph(contents, bookmark, value, false);
         if (!found) {
            found = f;
         }
         if (found) {
            debugger("WAS FOUND ON " + bookmark);
         }
      }

      // REPLACE BOOKMARKS IN MAIN DOCUMENT
      f = changeDocxParagraph(this.docxBodyContent, bookmark, value, false);
      if (!found) {
         found = f;
      }
      if (found) {
         debugger("WAS FOUND ON " + bookmark);
      }
      return found;
   }

   /**
    * Responsible for replacing text for .xlsx files
    * 
    * @param bookmark
    *           the name of the bookmark
    * @param value
    *           the value to be inserted in place of the bookmark
    */
   private boolean replaceForXlsx(String bookmark, String value) {
      boolean found = false;
      // REPLACE EXCEL TEXT
      for (org.xlsx4j.sml.CTRst si : siElements) {
         String siValue = si.getT();
         if (siValue.equals(bookmark)) {
            if (value.equals("null")) {
               si.setT(" ");
            } else {
               si.setT(value);
            }
            found = true;
         }
      }
      return found;
   }

   //------------------------------//
   
   /**
    * Responsible for replacing text for .pptx files
    * 
    * @param bookmark
    *           the name of the bookmark
    * @param value
    *           the value to be inserted in place of the bookmark
    */
   private boolean replaceForPptx(String bookmark, String value) {
      debugger("entering replaceForPptx");
      boolean found = false;

      for (SlidePart sld : pptSlideParts) {
         List<Object> lst = sld.getJaxbElement().getCSld().getSpTree()
               .getSpOrGrpSpOrGraphicFrame();
         for (Object o : lst) {
            if (o instanceof org.pptx4j.pml.Shape) {
               // Extract TextBody and replace bookmarks
               Shape shp = (Shape) o;
               CTTextBody ctText = shp.getTxBody();
               boolean justFound = processTxBody(ctText, bookmark, value);
               found = (found | justFound);
            } else if (o instanceof org.pptx4j.pml.CTGraphicalObjectFrame) {
               // Object is a table; parse each cell
               CTGraphicalObjectFrame grf = (CTGraphicalObjectFrame) o;
               boolean justFound = parseGraphicalObjectFrame(grf, bookmark,
                     value);
               found = (found | justFound);
            }
         }
      }

      for (SlideMasterPart smp : pptSlideMaster) {
         List<Object> lst = smp.getJaxbElement().getCSld().getSpTree()
               .getSpOrGrpSpOrGraphicFrame();
         for (Object o : lst) {
            if (o instanceof org.pptx4j.pml.Shape) {
               // Extract TextBody and replace bookmarks
               Shape shp = (Shape) o;
               CTTextBody ctText = shp.getTxBody();
               boolean justFound = processTxBody(ctText, bookmark, value);
               found = (found | justFound);
            }
         }
      }

      return found;
   }

   /**
    * Helper function that looks for tables on a slide.
    */
   private boolean parseGraphicalObjectFrame(CTGraphicalObjectFrame goFrame,
         String bookmark, String value) {
      boolean found = false;
      Graphic grp;
      GraphicData grd;
      List<Object> anyList;
      if ((grp = goFrame.getGraphic()) == null)   return false;
      if ((grd = grp.getGraphicData()) == null)   return false;
      if ((anyList = grd.getAny()) == null)       return false;
      for (Object o : anyList) {
         if (o instanceof JAXBElement
               && ((JAXBElement) o).getDeclaredType().getName()
                     .equals("org.docx4j.dml.CTTable")) {
            Object tmp = ((JAXBElement) o).getValue();
            if (tmp instanceof CTTable) {
               CTTable table = (CTTable) tmp;
               boolean justFound = processPPTXTable(table, bookmark, value);
               found = (found || justFound);
            }
         }
      }
      return found;
   }

   /**
    * Helper function which evaluates all textruns in all cells of a table.
    * Currently used on pptx only.
    */
   private boolean processPPTXTable(CTTable table, String bookmark, String value) {
      boolean found = false;
      List<CTTableRow> rows;
      List<CTTableCell> cells;
      if ((rows = table.getTr()) == null)  return false;
      for (CTTableRow r : rows) {
         if ((cells = r.getTc()) == null)  return false;
         for (CTTableCell c : cells) {
            CTTextBody ctText = c.getTxBody();
            boolean justFound = processTxBody(ctText, bookmark, value);
            found = (found || justFound);
         }
      }
      return found;
   }

   /**
    * Helper function which evaluates textruns. Currently used on pptx only.
    */
   private boolean processTxBody(CTTextBody ctText, String bookmark,
         String value) {

      boolean found = false;
      // Extract paragraphs from text.
      List<CTTextParagraph> CtParas = ctText.getP();
      // Iterate through paragraphs.
      for (CTTextParagraph ctPara : CtParas) {
         List<Object> txtRuns = ctPara.getEGTextRun();
         List<CTRegularTextRun> textRuns = new ArrayList<CTRegularTextRun>();
         for (Object obj : txtRuns) {
            if (obj instanceof CTRegularTextRun) {
               textRuns.add((CTRegularTextRun) obj);
            }
         }
         boolean justFound = textRunParser(textRuns, bookmark, value);
         found = (found || justFound);
      }
      return found;
   }

   /**
    * This helper function is meant to be a bulletproof way of substituting text
    * for pptx textruns. It has been tested, but needs some safety nets.
    */
   private boolean textRunParser(List<CTRegularTextRun> textRuns,
         String bookmark, String value) {
      boolean found = false;
      String runString = null;

      // Use loop to evaluate paragraphs with multiple textruns
      for (int startIndex = 0; startIndex < textRuns.size(); startIndex++) {
         try {
            CTRegularTextRun tRun = textRuns.get(startIndex);
            runString = tRun.getT();
            int markIndex = -1;
            // Use loop to evaluate textRuns with multiple bookmarks
            while ((markIndex = runString.indexOf("${", markIndex + 1)) != -1) {
               String foundTag = "${";
               int bracketIndex = runString.indexOf('}', markIndex);
               if (bracketIndex != -1) {
                  // Case 1: ${...} is fully included in this textRun.
                  foundTag = runString.substring(markIndex, bracketIndex + 1);
                  if (foundTag.equals(bookmark)) {
                     String regexTerm = Pattern.quote(bookmark);
                     runString = runString.replaceFirst(regexTerm, value);
                     tRun.setT(runString);
                     found = true;
                  }
               } else {
                  // Case 2: ${...} spans two or more textRuns.
                  boolean justFound = parseRunSpan(textRuns, bookmark, value,
                        startIndex, markIndex);
                  found = (found || justFound);
               }
            }
            if (runString.endsWith("$") && startIndex < textRuns.size() - 1) {
               // Case 3: textRun ends with '$'
               // look for "{" in the begining of next textrun.
               CTRegularTextRun nextRun = textRuns.get(startIndex + 1);
               String nextString = nextRun.getT();
               if (nextString.startsWith("{")) {
                  // look for "}" in some future textrun
                  boolean justFound = parseRunSpan(textRuns, bookmark, value,
                        startIndex, runString.length() - 1);
                  found = (found || justFound);
               }
            }
         } catch (NullPointerException npe) {
            // Tolerates null textRun or String and keeps looping.
         }
      }
      return found;
   }

   /**
    * Helper function; searches for a tag that spans multiple text runs, then
    * replaces with value if found. This is tested, however there are still
    * null-pointer vulnerabilities.
    */
   private boolean parseRunSpan(List<CTRegularTextRun> textRuns,
         String bookmark, String value, int startIndex, int markIndex)
         throws NullPointerException {
      if (textRuns == null)
         return false;
      boolean found = false;
      boolean foundBracket = false;
      int endIndex = startIndex + 1;
      int bracketIndex = 0;

      CTRegularTextRun tRun = textRuns.get(startIndex);
      CTRegularTextRun nextRun = null;

      String runString = tRun.getT();
      String nextString = "";
      String foundTag = runString.substring(markIndex);

      while (endIndex < textRuns.size() && !foundBracket) {
         try {
            nextRun = textRuns.get(endIndex);
            nextString = nextRun.getT();
            if (nextString.contains("}")) {
               foundBracket = true;
               bracketIndex = nextString.indexOf('}'); // exits loop
               foundTag = foundTag.concat(nextString.substring(0,
                     bracketIndex + 1));
            } else {
               foundTag = foundTag.concat(nextString);
               endIndex++;
            }
         } catch (NullPointerException npe) {
            // Tolerates null textRun or String and keeps looping.
            endIndex++;
         }
      }
      if (foundBracket && foundTag.equals(bookmark)) {
         found = true;
         // Substitute value in place of beginning of bookmark
         runString = runString.substring(0, markIndex);
         runString = runString.concat(value);
         tRun.setT(runString);
         // Remove '...}' from last run
         nextString = nextString.substring(bracketIndex + 1);
         nextRun.setT(nextString);
         // blank out any runs in between
         for (int j = endIndex - 1; j > startIndex; j--) {
            CTRegularTextRun midRun = textRuns.get(j);
            midRun.setT("");
            // TODO: Delete empty textRuns.
         }
      }
      return found;
   }

   /**
    * Returns an arraylist of SlideParts (for use with .PPT files)
    */
   private ArrayList<SlidePart> getpptSlideParts() {
      // TODO: Consider refactoring this function:
      // Assign directly to class variable instead of returning ArrayList.
      ArrayList<SlidePart> slides = new ArrayList<SlidePart>();
      org.docx4j.openpackaging.parts.Parts parts = pptMLPackage.getParts();
      HashMap<PartName, org.docx4j.openpackaging.parts.Part> hashmaps = parts.getParts();

      for (org.docx4j.openpackaging.parts.PartName o : hashmaps.keySet()) {
         if (o.getName().startsWith("/ppt/slides/")) {
            SlidePart temp = (SlidePart) hashmaps.get(o);
            slides.add(temp);
         }
         // Added for v1.5.3:
         // Also grab parts from Slide Masters and add them to
         // class-member list.
         // Unknown whether this list is properly disposed of after
         // first run.
         else if (o.getName().startsWith("/ppt/slideMasters/")) {
            pptSlideMaster.add((SlideMasterPart) hashmaps.get(o));
         }
      }
      return slides;
   }

   /**
    * Returns an arraylist of SlideParts (for use with .PPT files)
    */
   private ArrayList<WorksheetPart> getWorksheets() {
      ArrayList<WorksheetPart> sheets = new ArrayList<WorksheetPart>();
      org.docx4j.openpackaging.parts.Parts parts = xlsMLPackage.getParts();
      HashMap<org.docx4j.openpackaging.parts.PartName, org.docx4j.openpackaging.parts.Part> hashmaps = parts
            .getParts();

      for (org.docx4j.openpackaging.parts.PartName o : hashmaps.keySet()) {
         if (o.getName().startsWith("/xl/worksheets/")) {
            WorksheetPart temp = (WorksheetPart) hashmaps.get(o);
            sheets.add(temp);
         }
      }
      return sheets;
   }

   //-------------------------------//
   
   /**
    * Returns paragraph parent for bookmark found
    * 
    * @param parent
    *           the current parent node of the tree representation of the
    *           document
    * @param bookmark
    *           the name of the bookmark being considered
    */
   private boolean changeDocxParagraph(Object parent, String bookmark,
         String value, boolean found) {
      boolean f;
      debugger("entering changeDocxParagraph with " + bookmark);
      if (parent == null) {
         debugger("PARENT WAS NULL WITHIN getDocxBookmarkParent");
         throw new RuntimeException("Parent node was null");
      }
      // GET ALL CHILDREN OF PARENT
      List<Object> children = TraversalUtil.getChildrenImpl(parent);
      if (children != null) {
         // LOOP THROUGH EACH CHILD ELEMENT
         for (Object o : children) {
            o = XmlUtils.unwrap(o);
            if (parent instanceof org.docx4j.wml.P) {
               // CHECK IF BOOKMARK WAS FOUND
               f = processDocxRuns(children, ((P) parent), bookmark, value);
               if (!found) {
                  found = f;
               }
            }
            // RECURSE FARTHER DOWN THE TREE FOR A MATCH
            f = changeDocxParagraph(o, bookmark, value, found);
            if (!found) {
               found = f;
            }
         }
      }
      return found;
   }

   /**
    * Takes in a list of runs pertaining to a docx file. Checks if the bookmark
    * in question is located somewhere in the run. If so, it forwards to
    * collectDocxBookmarkedRuns. After the bookmark has been replaced it checks
    * again for more bookmarks in the parts list.
    * 
    * @param runs
    *           The list of runs pertaining to a paragraph in question
    * @param parentParagraph
    *           The parent paragraph of the list of runs in question
    * @param bookmark
    *           The bookmark in question
    * @param value
    *           The value of the bookmark in question to be replaced
    */
   private boolean processDocxRuns(List<Object> runs, P parentParagraph,
         String bookmark, String value) {
      debugger("entering processDocxRuns");
      boolean found = false;
      String collectiveContent = checkForDocxBookmark(runs, bookmark);
      if (collectiveContent.contains(bookmark)) {
         debugger("Match made with " + bookmark);
         found = true;

         // FIND INDICES OF BOOKMARK
         int startIndex = collectiveContent.indexOf(bookmark);
         int endIndex = startIndex + bookmark.length() - 1;
         collectDocxBookmarkedRuns(startIndex, endIndex, runs, bookmark, value,
               parentParagraph);

         // CHECK FOR MORE BOOKMARKS IN PARAGRAPH
         collectiveContent = checkForDocxBookmark(runs, bookmark);
         if (!collectiveContent.equals("false")) {
            processDocxRuns(runs, parentParagraph, bookmark, value);
         }
      }
      return found;
   }

   /**
    * Takes in a list of runs pertaining to a docx file. Checks if the bookmark
    * in question is contained within one of the runs. If so, it will return a
    * concatenated String of all of the runs. If not, it will return a String
    * that reads "false".
    * 
    * @param runs
    *           The list of runs pertaining to a paragraph in question
    * @param bookmark
    *           The bookmark in question
    */
   private String checkForDocxBookmark(List<Object> runs, String bookmark) {
      String collectiveContent = "";
      for (Object o : runs) {
         // CHILDREN OF RUNS
         List<Object> children = TraversalUtil.getChildrenImpl(o);
         if (children != null) {
            // LOOP THROUGH CHILDREN OF THE RUN
            for (Object o2 : children) {
               // UNWRAP OBJECT
               o2 = XmlUtils.unwrap(o2);
               // CHECK IF CURRENT OBJECT IS A TEXT OBJECT
               if (o2 instanceof org.docx4j.wml.Text) {
                  collectiveContent = collectiveContent
                        + ((org.docx4j.wml.Text) o2).getValue();
               }
            }
         }
      }
      debugger("collective string is " + collectiveContent);
      if (!collectiveContent.contains(bookmark)) {
         collectiveContent = "false";
      }
      return collectiveContent;
   }

   /**
    * Takes in a list of runs, iterates through them and determines which runs
    * contain any portion of the bookmark in question.
    * 
    * @param r
    *           The list of runs pertaining to a paragraph in question
    * @param bookmark
    *           The bookmark in question
    * @param value
    *           The value of the bookmark in question to be replaced
    * @param parentParagraph
    *           The parent paragraph of the list of runs in question
    */
   private void collectDocxBookmarkedRuns(int startIndex, int endIndex,
         List<Object> r, String bookmark, String value, P parentParagraph) {
      debugger("enter collectDocxBookmarkedRuns");
      boolean inBookmark = false;
      boolean endFound = false;
      boolean startFound = false;
      List<Object> runs = new ArrayList<Object>();
      // LOOP THROUGH CHILDREN OF THE RUN
      for (Object o : r) {
         debugger("entered new run -- in collectDocxBookmarkedRuns");
         List<Object> children = TraversalUtil.getChildrenImpl(o);
         if (children != null && !endFound) {
            for (Object o2 : children) {
               debugger("entered new text object -- in collectDocxBookmarkedRuns");
               // UNWRAP OBJECT
               o2 = XmlUtils.unwrap(o2);
               // CHECK IF CURRENT OBJECT IS A TEXT OBJECT
               if (o2 instanceof org.docx4j.wml.Text) {
                  String t = ((org.docx4j.wml.Text) o2).getValue();
                  debugger("looking at string " + t);
                  if (inBookmark) {
                     debugger("CURRENTLY IN A BOOKMARK");
                  }
                  int len = t.length();
                  debugger("startIndex is " + startIndex + " endIndex is "
                        + endIndex + " current length is " + len);
                  if (!inBookmark && startIndex <= len - 1) {
                     // runs.add(((org.docx4j.wml.R)o));
                     if (!inBookmark) {
                        inBookmark = true;
                        startFound = true;
                     }
                  }
                  if (inBookmark) {
                     debugger("run added with string " + t);
                     runs.add(o);
                  }
                  if (inBookmark && endIndex <= len - 1) {
                     endFound = true;
                     inBookmark = false;
                     debugger("end of bookmark found");
                  }
                  if (!endFound) {
                     endIndex -= len;
                     debugger("endIndex is now" + endIndex);
                  }
                  if (!startFound) {
                     startIndex -= len;
                     debugger("startIndex is now" + startIndex);
                  }
               }
            }
         }
      }
      if (!runs.isEmpty()) {
         debugger("startIndex is " + startIndex + " endIndex is " + endIndex);
         editDocxRuns(runs, startIndex, endIndex, bookmark, value,
               parentParagraph);
      } else {
         debugger("RUNS IS EMPTY!");
      }
   }

   /**
    * Takes in a list of runs, the index of where the bookmark starts, the index
    * of where the bookmark ends and edits the runs accordingly. This method is
    * responsible for replacing bookmarks with their corresponding values in
    * docx.
    * 
    * @param runs
    *           The list of runs pertaining to a paragraph in question
    * @param startIndex
    *           The index of where the bookmark starts in the concatenation of
    *           the runs
    * @param endIndex
    *           The index of where the bookmark ends in the concatenation of the
    *           runs
    * @param bookmark
    *           The bookmark in question
    * @param value
    *           The value of the bookmark in question to be replaced
    * @param parentParagraph
    *           The parent paragraph of the list of runs in question
    */
   private void editDocxRuns(List<Object> runs, int startIndex, int endIndex,
         String bookmark, String value, P parentParagraph) {
      List<Object> parentNodes = parentParagraph.getContent();
      debugger("enter editDocxRuns$$$$$$$$$$$$$$$$$$$");
      boolean fixed = false;
      // ALTER FIRST RUN IN LIST
      Object first = runs.get(0);
      Object last = runs.get(runs.size() - 1);
      org.docx4j.wml.R firstRun = (org.docx4j.wml.R) first;

      List<Object> children = TraversalUtil.getChildrenImpl(firstRun);
      if (children != null) {
         debugger("in first run");
         for (Object o : children) {
            // UNWRAP OBJECT
            o = XmlUtils.unwrap(o);
            if (o instanceof org.docx4j.wml.Text && !fixed) {
               debugger("in new text element");
               String content = ((org.docx4j.wml.Text) o).getValue();
               if (runs.size() == 1
                     && !value.toLowerCase().trim().equals("null")) {
                  ((org.docx4j.wml.Text) o).setValue(content.replace(bookmark,
                        value));
                  fixed = true;
                  debugger("RUN WAS FIXED. CONTENT WAS " + content);
                  debugger("Content is now "
                        + ((org.docx4j.wml.Text) o).getValue());
                  break;
               } else {
                  if (value.toLowerCase().trim().equals("null")) {
                     int spaceFiller = bookmark.length();
                     String filler = "";
                     for (int i = 0; i < spaceFiller; i++) {
                        char spc = ' ';
                        filler = filler + spc;
                        debugger("null found");
                     }
                     content = content.substring(0, startIndex);
                     ((org.docx4j.wml.Text) o).setValue(content);
                     debugger("RUN WAS FIXED but more remain. CONTENT WAS "
                           + content);
                     debugger("Content is now "
                           + ((org.docx4j.wml.Text) o).getValue());
                     break;
                  } else {
                     content = content.substring(0, startIndex);
                     ((org.docx4j.wml.Text) o).setValue(content + " " + value);
                     debugger("RUN WAS FIXED but more remain. CONTENT WAS "
                           + content);
                     debugger("Content is now "
                           + ((org.docx4j.wml.Text) o).getValue());
                     break;
                  }
               }
            }
         }
      }

      // UPDATE ENDINDEX
      endIndex = endIndex - startIndex;
      debugger("endIndex is now " + endIndex);

      // ALTER LAST RUN IN LIST
      if (!fixed) {
         org.docx4j.wml.R lastRun = (org.docx4j.wml.R) last;
         List<Object> children2 = TraversalUtil.getChildrenImpl(lastRun);
         if (children2 != null) {
            debugger("in last run");
            for (Object o : children2) {
               // UNWRAP OBJECT
               o = XmlUtils.unwrap(o);
               if (o instanceof org.docx4j.wml.Text) {
                  debugger("in last run. text element");
                  String content = ((org.docx4j.wml.Text) o).getValue();
                  if (content.length() > 0) {
                     content = content
                           .substring(endIndex + 1, content.length());
                     if (endIndex == content.length() - 1) {
                        // ((org.docx4j.wml.Text)o).setValue("");
                        parentNodes.remove(o);
                     } else {
                        ((org.docx4j.wml.Text) o).setValue(content);
                     }
                     fixed = true;
                     debugger("RUN WAS FIXED. CONTENT WAS " + content);
                     debugger("CONTENT IS "
                           + ((org.docx4j.wml.Text) o).getValue());
                  }
               }
            }
         }
      }

      // REMOVE EXTRA RUNS IN BETWEEN
      if (first != null) {
         runs.remove(first);
      }
      if (last != null) {
         runs.remove(last);
      }
      if (!runs.isEmpty()) {
         for (Object x : runs) {
            List<Object> childrenX = TraversalUtil.getChildrenImpl(x);
            if (childrenX != null) {
               for (Object x2 : childrenX) {
                  x2 = XmlUtils.unwrap(x2);
                  if (x2 instanceof org.docx4j.wml.Text) {
                     if (((org.docx4j.wml.Text) x2) != first
                           && ((org.docx4j.wml.Text) x2) != last) {
                        ((org.docx4j.wml.Text) x2).setValue("");
                        docxBodyContent.remove(x2);
                     }
                  }
               }
            }
         }
      }
   }

   /**
    * Removes all of the calculated values belonging to formulas so that they
    * are recalculated when the document is opened.
    */
   private void removeXLSXFormulaValues() {
      // PREPARE COLLECTION OF WORKSHEETS IN DOCUMENT
      ArrayList<WorksheetPart> sheets = getWorksheets();

      // ITERATE THROUGH EACH WORKSHEET
      for (WorksheetPart sheet : sheets) {
         org.xlsx4j.sml.SheetData sData = sheet.getJaxbElement().getSheetData();
         List<org.xlsx4j.sml.Row> rows = sData.getRow();
         // ITERATE THROUGH THE ROWS
         for (org.xlsx4j.sml.Row row : rows) {
            // GET CELLS
            List<org.xlsx4j.sml.Cell> cells = row.getC();
            // ITERATE THROUGH THE CELLS
            for (org.xlsx4j.sml.Cell cell : cells) {
               org.xlsx4j.sml.STCellType a = cell.getT();
               if (a == org.xlsx4j.sml.STCellType.E) {
                  String formula = cell.getF().getValue();
                  cell.setV(null);
               }
            }
         }
      }
   }

   /**
    * Returns the HeaderParts and FooterParts of a docx file
    */
   private List<JaxbXmlPart> getDocxHeaderFooterParts() {
      debugger("in getDocxHeaderFooterParts");
      // Add headers/footers
      RelationshipsPart rp = docxDocumentPart.getRelationshipsPart();
      List<JaxbXmlPart> parts = new ArrayList<JaxbXmlPart>();
      for (Relationship r : rp.getJaxbElement().getRelationship()) {
         if (r.getType().equals(Namespaces.HEADER)
               || r.getType().equals(Namespaces.FOOTER)) {
            JaxbXmlPart part = (JaxbXmlPart) rp.getPart(r);
            parts.add((part));
         }

         debugger("GOT RelationshipsPart");
      }
      return parts;
   }

   /**
    * Turn off log4j logging.
    */
   protected void turnOffLogging() {
      org.docx4j.Docx4jProperties.getProperties().setProperty(
            "docx4j.Log4j.Configurator.disabled", "true");
      org.docx4j.utils.Log4jConfigurator.configure();
   }
}