namespace QcGoldArchive
{
    class UniversalStrings
    {
        public static string QC_GOLD_ARCHIVE = "QC Gold Archive";
        
        public static string DateTimeFormatEU = "dd/MM/yy  HH:mm";
        public static string DateFormatEU = "dd/MM/yy";
        public static string DateTimeFormatUS = "MM/dd/yy  HH:mm";
        public static string DateFormatUS = "MM/dd/yy";

        public static string Font = "Ariel";

        public static string ENG_US = "en-US";
        public static string GER_GERMANY = "de-DE";
        public static string IT_ITALIA = "it-IT";
        public static string FR_FRENCH = "fr-FR";
        public static string CH_CHINESE = "zh-CN";


        public static string EMPTY_VALUE = "---";

        public static string[] REMOVE_LABELS = {"DEVICE_SN", "SW_VERSION", "TEST_DATE", "PATIENT_ID", "VOID", "BIRTH_DATE", "ABSTINENCE",
                               "ACCESSION", "VOID", "COLLECTED_DATE", "RECEIVED", "TYPE", "VOLUME", "WBC_CONC",
                               "PH", "VOID", "CONC", "VOID", "PR_NP_REM", "VOID", "PROG_REM", "NONPROG", "IMMOT",
                               "VOID", "WHO_5_REM", "MSC", "PMSC", "FSC", "VELOCITY", "SMI", "VOID", "NUM_SPERM",
                               "MOTILE_SPERM", "PROG_SPERM", "FUNC_SPERM", "VOID", "SPERM_REM"};

        public static string[] FROZEN_LABELS = {"MSC", "PMSC", "VELOCITY", "SMI", "MOTILE_SPERM", "PROG_SPERM" };

        public static string[] UNITS = {"DAYS", "ML", "M_PER_ML", "PERCENT", "M", "MIC_PER_SEC"};

        public static string[] NEW_FILE_LABELS = {"TEST_DATE", "PATIENT_ID","PAT_NAME" ,"BIRTH_DATE", "ABSTINENCE_WU", "ACCESSION", "COLLECTED_DATE",
                               "RECEIVED", "TYPE", "VOLUME_WU", "WBC_CONC_WU", "PH", "CONC_WU","PR_NP_WU", "PROG_WU", "NONPROG_WU", 
                               "IMMOT_WU", "WHO_5_WU", "MSC_WU", "PMSC_WU", "FSC_WU", "VELOCITY_WU", "SMI", "NUM_SPERM_WU", 
                               "MOTILE_SPERM_WU", "PROG_SPERM_WU", "FUNC_SPERM_WU", "SPERM_WU","TESTER_NAME","TESTER_DESIGNATION","FRUCTOSE","LIQUIFACTION","Referred_dr","REF_BY_DR2","REF_BY_DR3","Avg_str","CNT_Str","OD_Str","aw_Data","SW_VERSION", "DEVICE_SN_WU" };
        public static string[] NEW_FILE_INDIAN_LABELS = {"TEST DATE", "PATIENT ID","PATIENT NAME" ,"BIRTH DATE", "ABSTINENCE", "ACCESSION", "COLLECTED DATE",
                               "RECEIVED DATE", "TYPE", "VOLUME", "WBC_CONC", "PH", "CONC.","Total Motile PR_NP", "PROG_WU", "NONPROG_WU",
                               "IMMOTILE", "WHO_5", "MSC", "PMSC", "FSC", "VELOCITY", "SMI", "SPERM #",
                               "MOTILE_SPERM ", "PROG_SPERM ", "FUNC_SPERM ", "SPERM","TESTER NAME","TESTER DESIGNATION","FRUCTOSE","LIQUIFACTION","Referred_dr","REF_BY_DR2","REF_BY_DR3","Avg_str","CNT_Str","OD_Str","aw_Data","AGE","OPTIONAL1","MAN_FRUCTOSE","MAN_VITALITY","MAN_RBC","MAN_ROUNDEDCELLS","MAN_AGGREGATION","MAN_AGGLUTINATION","MAN_OPTIONAL","MAN_NORMALFORMS","MAN_HEADDEFECT","MAN_NECK","MAN_TAIL","MAN_CYTOPLASM","MAN_ACROSOME","MAN_PINHEADS","COMMENTS","SW_VERSION", "DEVICE_SN_WU" };

        public static string[] NEW_EXCEL_FILE_LABELS = {"TEST_DATE", "PATIENT_ID", "BIRTH_DATE", "ABSTINENCE_WU", "ACCESSION", "COLLECTED_DATE",
                               "RECEIVED", "TYPE", "VOLUME_WU", "WBC_CONC_WU", "PH", "CONC_WU","PR_NP_WU", "PROG_WU", "NONPROG_WU",
                               "IMMOT_WU", "WHO_5_WU", "MSC_WU", "PMSC_WU", "FSC_WU", "VELOCITY_WU", "SMI", "NUM_SPERM_WU",
                               "MOTILE_SPERM_WU", "PROG_SPERM_WU", "FUNC_SPERM_WU", "SPERM_WU","PAT_NAME","Tes_name","TESTER_DESIGNATION","FRUCTOSE","LIQUIFICTION","Ref_By_Dr","REF_BY_DR2","REF_BY_DR3", "SW_VERSION", "DEVICE_SN_WU" };

    }
   
    class MessagesLabelsEN
    {
        public static string Arc_same_test = "You can average two tests that were taken of the same sample within 10 min";
        public static string Average = "Average";
        public static string AVG_Val = "AVRG";
        public static string COMMENTS = "Comments : ";
        public static string Average_Report_subhead = "By the Automated Semen Quality Analyzer";
        public static string Filter_translation = "Ensure The Translation Type";
        public static string ARCHIVE = "Archive";
        public static string Average_Report_head = "SEMEN ANALISYS AVERAGE REPORT";
        public static string good = "guideline Good";
        public static string Medium = "guideline Medium";
        public static string Ref_Val = "REF.VALUE";
        public static string System_Language = "Select the CSV file according to your system language";
        public static string Enter = "Enter";
        public static string Title = "Title";
        public static string Sub_Title = "Sub Title";
        public static string ADD_SIGNATURE = "Add Signature";
        public static string Motility = "MOTILITY";
        public static string SIGNATURE = "Signature";
        public static string Analysing_Charts = "ANALYSING CHARTS";
        public static string Old_File_MESSAGE = "You have chosen an older archive file. Please choose a new file, otherwise all your old tests will be erased.Would you like to continue? ";
        public static string BASIC_REPORT = "Basic Report";
        public static string Warning = "Warning";
        public static string sys_info_label = "SYSTEM INFORMATION ";
        public static string Print_Warning = "Please Select Printer Name";
        public static string Print_Error_Warning = "Select printer";
        public static string Use_Default = "Use Default";
        public static string HeaderSpace_Warning = "Please Enter Header Space Value";
        public static string HeaderSpace_Error_Warning = "Enter value";
        public static string Title_Warning = "Please Enter The title of the report ";
        public static string Title_Error_Warning = "Enter Value";
        public static string Report_title = "Report Title ";
        public static string WRONG_DATA_ERROR = "Make sure the file contains valid data";
        public static string NAME = "Tester Name";
        public static string Max_Exceed = "Maximum value is 30 mm ";
        public static string Report_type = "Report Type ";
        public static string SAVE_DATA = "There is unsaved data. Do you want to save the changes?";
        public static string NO_MOTILITY_CHART = "MOTILITY IS CHART NOT AVAILABLE";
        public static string NO_SMI_CHART = "SMI CHART IS NOT AVAILABLE";
        public static string IMMOTILITY_TEXT = "IMMOTILITY ";
        public static string TOTALPROGRESSIVE_MOTILITY = "TOTAL PROGRESSIVE \nMOTILITY ";
        public static string NONPROGRESSIVE_MOTILITY = "NON PROGRESSIVE \nMOTILITY ";
        public static string Clear_Result = "Clear Result";
        public static string CLEAR_MESSAGE = "Do You Want To Clear The Result?";
        public static string PRINT_RESULTS_STRIP = "Print Result Strip";
        public static string DESIGNATION = "Designation";
        public static string SEMEN_EXAMINATION = "SEMEN EXAMINATION";
        public static string ADVANCED_REPORT = "Advanced Report";
        public static string PATIENT_NAME = "PATIENT NAME";
        public static string TESTER_NAME = "TESTER NAME";
        public static string SAMPLE_TYPE = "SAMPLE TYPE";
        public static string SPERMS = "SPERM (M)";
        public static string Fsc = "FSC (M/ml)";
        public static string MOT_SPERMS = "MOT. SPERM (M)";
        public static string PROG_SPERMS = "PROG. SPERM (M)";
        public static string FUNC_SPERMS = "FUNC. SPERM (M)";
        public static string PROG_MOT = "PROG.MOT.(%)";
        public static string NON_PROG_MOT = "NONPROG.MOT.(%)";
        public static string MORPH_NORM_FORMS = "MORPH.NORM.SPERM(M)";
        public static string IMOTILITY = "IMMOT. (%) ";
        public static string TESTER_DESG = "DESIGNATION:";
        public static string MORPH_NORM = "MORPH.NORM.FORM(%)";
        public static string TOT_MOT = "TOTAL MOT. (%)";
        public static string TESTER_INFO = "TESTER INFORMAION";
        public static string RESULT_INFO = "RESULT INFORMATION";
        public static string REF_BY_DR = "REFERRED BY DR.:";
        public static string PORT_ERROR = "Can't open port";
        public static string PORT_ERROR_CAPTION = "Port Error";
        public static string NO_REPORT_ERROR = "There Is No Report For This Data";
        public static string REPORT_Error = "Report error ";
        public static string Archive_Error = "Could Not Open Archive";
        public static string Archive_Error_Missing = "Archive is empty or doesn’t exist";
        public static string Report_head = "SEMEN ANALYSIS TEST REPORT";
        public static string Report_Subhead = "By Automated Semen Analisys";
        public static string EXPORT_ERROR = "Unexpected Error. Data was not exported to Excel";
        public static string END_OF_REPORT = "end of report";
        public static string REPORT_ERROR = "No data available. Please import the test results again ";
        public static string REPORT_ERROR_CAPTION = "Report error";
        public static string Numerical_Error = "Enter only numerical  Values";
        public static string LIQUIFICTION = "LIQUEFACTION";
        public static string REPORT_CREATION_ERROR = "Could not create report";
        public static string P_Name = "PATIENT NAME";
        public static string DEVICE_SN = "DEVICE SN";
        public static string FRUCTOSE = "FRUCTOSE";
        public static string MANUAL_IMPORT_FAILED = "Manual Data could not be saved";
        public static string SW_VERSION = "SW VER.";
        public static string TEST_DATE = "TEST DATE";
        public static string PATIENT_ID = "PATIENT ID";
        public static string VOID = " ";
        public static string AGE = "AGE";
        public static string BIRTH_DATE = "BIRTH DATE";
        public static string ABSTINENCE = "ABSTINENCE";
        public static string ACCESSION = "ACCESSION #";
        public static string COLLECTED_DATE = "COLLECTED";
        public static string RECEIVED = "RECEIVED";
        public static string TYPE = "TYPE";
        public static string VOLUME = "VOLUME";
        public static string WBC_CONC = "WBC CONC.";
        public static string PH = "PH";
        public static string CONC = "CONC.";
        public static string PR_NP = "TOTAL MOTILITY <PR+NP>";
        public static string PR_NP_REM = "<PR+NP>";
        public static string PROG = "MOTILITY GRADES: PROG. <PR>";
        public static string PROG_REM = "PROG. <PR>";
        public static string MANUAL_ENTRY = "Manual Entry";

        public static string NONPROG = "NONPROG. <NP>";
        public static string IMMOT = "IMMOT. <IM>";
        public static string WHO_5 = "MORPH.NORM.FORMS <WHO 5th>";
        public static string WHO_5_REM = "<WHO 5th>";
        public static string MSC = "MSC";
        public static string PMSC = "PMSC";
        public static string FSC = "FSC";
        public static string VELOCITY = "VELOCITY";
        public static string SMI = "SMI";
        public static string NUM_SPERM = "SPERM #";
        public static string MOTILE_SPERM = "MOTILE SPERM";
        public static string PROG_SPERM = "PROG. SPERM";
        public static string FUNC_SPERM = "FUNC. SPERM";
        public static string SPERM = "MORPH.NORM.SPERM";
        public static string SPERM_REM = "SPERM";
        public static string SYSTEM_DATA = "SYSTEM DATA";

        public static string BIRTH = "BIRTH";
        public static string COLLECTED = "COLLECTED";

        public static string MOTILITY_RESULTS = "motility results only";

        public static string DAYS = "DAYS";
        public static string ML = "ml";
        public static string M_PER_ML = "M/ml";
        public static string PERCENT = "%";
        public static string M = "M";
        public static string MIC_PER_SEC = "mic/sec";

        public static string FROZEN = "FROZEN";
        public static string FRESH = "FRESH";
        public static string WASHED = "WASHED";

        public static string RECEIVED_WU = "RECEIVED (Date/Time)";  /// With Unit
        public static string COLLECTED_DATE_WU = "COLLECTED (Date/Time)";
        public static string TEST_DATE_WU = "TEST DATE (Date/Time)";
        public static string ABSTINENCE_WU = "ABSTINENCE (Days)"; 
        public static string VOLUME_WU = "VOLUME (ml)";
        public static string WBC_CONC_WU = "WBC CONC. (M/ml)";
        public static string CONC_WU = "CONC. (M/ml)";
        public static string PR_NP_WU = "TOTAL MOTILITY <PR+NP> (%)";
        public static string PROG_WU = "PROG. MOTILITY <PR> (%)";
        public static string NONPROG_WU = "NONPROG MOTILITY <NP> (%)";
        public static string IMMOT_WU = "IMMOTILITY (%)";
        public static string WHO_5_WU = "MORPH.NORM.FORMS <WHO 5th> (%)";
        public static string MSC_WU = "MSC (M/ml)";
        public static string PMSC_WU = "PMSC (M/ml)";
        public static string FSC_WU = "FSC (M/ml)";
        public static string VELOCITY_WU = "VELOCITY (mic/sec)";
        public static string NUM_SPERM_WU = "SPERM # (M)";
        public static string MOTILE_SPERM_WU = "MOTILE SPERM (M)";
        public static string PROG_SPERM_WU = "PROG. SPERM (M)";
        public static string FUNC_SPERM_WU = "FUNC. SPERM (M)";
        public static string SPERM_WU = "MORPH.NORM.SPERM (M)";
        public static string DEVICE_SN_WU = "DEVICE SN #";

        public static string TITLE_TEST_RESULTS = "RESULTS";
        public static string PATIENT_DATA = "Patient Data";
        public static string SAMPLE_DATA = "Sample Data";
        public static string TEST_RESULTS = "Test Results";
        public static string TOTALS_PER_VOL = "Totals Per Volume";
        public static string OTHER = "Other";
        public static string EXCEL_FILE_ERROR_CAPTION = "Excel file error";
        public static string ACCESS_ERROR = "Cannot be accessed. If the file is open, please close it and press retry";
        public static string IMPORT_FAILED = "Import Failed";
        public static string EXPORT_FILE_ERROR = "Excel file error. Data was not exported to Excel";
        public static string PRINTER_ERROR = "Printer Error";
        public static string SETTINGS_WINDOW_ALREADY_OPENED = "Settings window is already opened";
        public static string SETTINGS = "Settings";
        public static string SETTINGS_ERROR = "Settings error ";
        public static string CLEAR = "Clear";
        public static string PRINT = "Export PDF";
        public static string YES = "Yes";
        public static string NO = "No";
        public static string EXPORT_FILE_NAME = "Export File Name: ";
        public static string VERSION = "Version: ";
        public static string SETTINGS_FILE_CORRUPTED_DEFAULT = "Settings file is corrupted. Default settings will be used";
        public static string SETTINGS_FILE_CORRUPTED_RELOAD = "Settings file is corrupted. Please reload the application";
        public static string NONE = "None";
        public static string SELECT_FILE_TO_EXPORT = "Select a file to export";
        public static string EXCEL_NOT_EXIST = "Excel does not exist";
        public static string SAVE_CHANGES = "Save Changes?";
        public static string CONFIRM_CHANGES = "Confirm Changes";
        public static string SPECIFY_FILE_NAME = "Please specify export file name";
        public static string EXPORT_FILENAME_MISSING = "Export filename missing";
        public static string INVALID_FILEPATH = "Invalid file path";
        public static string INVALID_FILENAME = "Invalid filename";
        public static string SETTINGS_FILE_ERROR = "Settings File Errror";
        public static string FILE_OPEN_ERROR = "Unable to export. Please close the file.";
        public static string EXPORT_ERROR_CAPTION = "Export Error";
        public static string CSV_OPEN_ERROR = "Make sure that the CSV file is not open ";
        public static string PORT = "Port";
        public static string PRINTER = "Printer";
        public static string EXPORT_FILE = "ExportFile";
        public static string AUTO_PRINT = "AutoPrint";
        public static string EXPORT = "Export";
        public static string DATE_FORMAT = "DateFormat";

        public static string SELECT_PORT = "Select Port: ";
        public static string PRINTER_NAME = "Printer Name: ";
        public static string LANGUAGE = "Language";
        public static string CHOOSE = "Choose";
        public static string EXPORT_TO_EXCEL = "Export to Excel";
        public static string EXPORT_TO_EXCEL_SETTINGS = "Export to CSV Settings";
        public static string FILE_NAME = "File Name: ";
        public static string DATE_FORMAT_SET = "Date Format: ";
        public static string OPEN = "Open";
        public static string SAVE = "Save";
        public static string CANCEL = "Cancel";
        public static string PRINTED_FROM = "Printed from";
        public static string REPORT_SETTINGS = "Report Settings";
        public static string MAKE_PDF_REPORT = "Create pdf report";
        public static string PRINT_TEST_RESULTS = "Print results strip";
        public static string ADD_HEADER_SPACE = "Add header space";
        public static string HEADER_SPACE_UNIT = "mm";
        public static string EXPORT_PDF = "Export Pdf";



        public static string REPORT_FILE_NAME = "Semen Analysis Report";
        public static string TEST_REPORT = "SEMEN ANALYSIS TEST REPORT";
        public static string REPORT_CONTINUED = "REPORT CONTINUED";
        public static string DEVICE_INFO = "DEVICE INFORMATION";
        public static string PATIENT_INFO = "PATIENT INFORMATION";
        public static string SAMPLE_INFO = "SAMPLE INFORMATION";
        public static string PARAMETER = "PARAMETER";
        public static string RESULT = "RESULT";

        public static string RESULT_INFORMATION = "RESULT INFORMATION";
        public static string ARCHIVE_HEADING = "Archive Test Results";
        public static string SEARCH_BY = "Search by";
        public static string SEARCH = "Search";
        public static string CLEAR_FILTER = "Clear Filter";
        public static string CONFIRM = "Confirm";
        

    }

    class MessagesLabelsDE
    {
        public static string PORT_ERROR = "Port kann nicht geöffnet werden";
        public static string PORT_ERROR_CAPTION = "Port Fehler";
        public static string EXPORT_ERROR = "Unerwarteter Fehler. Daten wurden nicht nach Excel exportiert";
        public static string END_OF_REPORT = "Ende des Reports";
        public static string REPORT_ERROR = "Keine Daten verfügbar. Bitte importieren Sie die Testergebnisse neu.";
        public static string REPORT_ERROR_CAPTION = "Report Fehler";
        public static string REPORT_CREATION_ERROR = "Report konnte nicht erstellt werden.";

        public static string DEVICE_SN = "GERÄTE SN #";
        public static string SW_VERSION = "SW VER.";
        public static string TEST_DATE = "TEST DATUM";
        public static string PATIENT_ID = "PATIENTEN ID";
        public static string VOID = " ";
        public static string BIRTH_DATE = "GEBURTSDATUM";
        public static string ABSTINENCE = "ABSTINENZ";
        public static string ACCESSION = "EINTRAG #:";
        public static string COLLECTED_DATE = "PROBENAHME";
        public static string RECEIVED = "PROBENEINGANG";
        public static string TYPE = "ART";
        public static string VOLUME = "VOLUMEN";
        public static string WBC_CONC = "LEUKOZYTENKONZ.";
        public static string PH = "PH";
        public static string CONC = "KONZ.";
        public static string PR_NP = "MOTILITÄT <PR+NP>";
        public static string PROG = "MOTILITÄTSGRAD: PROG. <PR>";
        public static string NONPROG = "NICHT-PROG. <NP>";
        public static string IMMOT = "IMMOT. <IM>";
        public static string WHO_5 = "MORPH.NORM.FORMEN <WHO 5th>";
        public static string MSC = "MSC";
        public static string PMSC = "PMSC";
        public static string FSC = "KONZ. FUNKTIONELLER SPERMIEN";
        public static string VELOCITY = "GESCHWINDIGKEIT";
        public static string SMI = "SMI";
        public static string NUM_SPERM = "SPERMIEN #";
        public static string MOTILE_SPERM = "MOTILE SPERMIEN";
        public static string PROG_SPERM = "PROG. SPERMIEN";
        public static string FUNC_SPERM = "FUNK. SPERMIEN";
        public static string SPERM = "MORPH.NORM.SPERMIEN";
        public static string SYSTEM_DATA = "SYSTEM DATEN";

        public static string BIRTH = "GEBURT";
        public static string COLLECTED = "PROBENAHME";

        public static string MOTILITY_RESULTS = "nur Motilitätsergebnisse";

        public static string DAYS = "TAGE";
        public static string ML = "ml";
        public static string M_PER_ML = "M/ml";
        public static string PERCENT = "%";
        public static string M = "M";
        public static string MIC_PER_SEC = "µ/s";

        public static string FROZEN = "GEFROREN";
        public static string FRESH = "FRISCH";
        public static string WASHED = "GEWASCHEN";

        public static string RECEIVED_WU = "PROBENEINGANG (Datum/Zeit)";  // With Unit
        public static string COLLECTED_DATE_WU = "PROBENAHME (Datum/Zeit)";
        public static string TEST_DATE_WU = "TEST DATUM (Datum/Zeit)";
        public static string ABSTINENCE_WU = "ABSTINENZ (Tage)";
        public static string VOLUME_WU = "VOLUMEN (ml)";
        public static string WBC_CONC_WU = "LEUKOZYTENKONZ. (M/ml)";
        public static string CONC_WU = "KONZ. (M/ml)";
        public static string PR_NP_WU = "GESAMT-MOTILITÄT <PR+NP> (%)";
        public static string PROG_WU = "PROG. MOTILITÄT <PR> (%)";
        public static string NONPROG_WU = "NICHT-PROG. MOTILITÄT <NP> (%)";
        public static string IMMOT_WU = "IMMOTILITÄT (%)";
        public static string WHO_5_WU = "MORPH.NORM.FORMEN <WHO 5th> (%)";
        public static string MSC_WU = "MSC (M/ml)";
        public static string PMSC_WU = "PMSC (M/ml)";
        public static string FSC_WU = "KONZ. FUNKTIONELLER SPERMIEN (M/ml)";
        public static string VELOCITY_WU = "GESCHWINDIGKEIT (µ/s)";
        public static string NUM_SPERM_WU = "SPERMIEN # (M)";
        public static string MOTILE_SPERM_WU = "MOTILE SPERMIEN (M)";
        public static string PROG_SPERM_WU = "PROG. SPERMIEN (M)";
        public static string FUNC_SPERM_WU = "FUNK. SPERMIEN (M)";
        public static string SPERM_WU = "MORPH.NORM.SPERMIEN (M)";
        public static string DEVICE_SN_WU = "GERÄTE SN #";

        public static string TITLE_TEST_RESULTS = "QwikCheck GOLD: SAMEN ANALYSE TEST ERGEBNISSE";
        public static string PATIENT_DATA = "Patientendaten";
        public static string SAMPLE_DATA = "Probedaten";
        public static string TEST_RESULTS = "Testergebnisse";
        public static string TOTALS_PER_VOL = "Gesamtanzahl Pro Volumen";
        public static string OTHER = "Andere";
        public static string EXCEL_FILE_ERROR_CAPTION = "Excel-Datei-Fehler";
        public static string ACCESS_ERROR = "Kein Zugriff. Falls die Datei geöffnet ist, bitte schließen Sie sie und klicken Sie Wiederholen";
        public static string IMPORT_FAILED = "Import fehlgeschlagen";
        public static string EXPORT_FILE_ERROR = "Excel-Datei-Fehler. Daten wurden nicht nach Excel exportiert";
        public static string PRINTER_ERROR = "Drucker-Fehler";
        public static string SETTINGS_WINDOW_ALREADY_OPENED = "Einstellungen-Fenster ist bereits geöffnet";
        public static string SETTINGS = "Einstellungen";
        public static string CLEAR = "Löschen";
        public static string PRINT = "Report drucken";
        public static string YES = "Ja";
        public static string NO = "Nein";
        public static string EXPORT_FILE_NAME = "Export-Dateiname: ";
        public static string VERSION = "Version: ";
        public static string SETTINGS_FILE_CORRUPTED_DEFAULT = "Einstellungen-Datei ist beschädigt. Default Einstellungen werden verwendet";
        public static string SETTINGS_FILE_CORRUPTED_RELOAD = "Einstellungen-Datei ist beschädigt. Bitte starten Sie die Anwendung neu";
        public static string NONE = "Keine";
        public static string SELECT_FILE_TO_EXPORT = "Wählen Sie eine Datei für den Export";
        public static string EXCEL_NOT_EXIST = "Excel existiert nicht";
        public static string SAVE_CHANGES = "Änderungen speichern?";
        public static string CONFIRM_CHANGES = "Änderungen bestätigen";
        public static string SPECIFY_FILE_NAME = "Bitte geben Sie den Export-Datei-Namen an";
        public static string EXPORT_FILENAME_MISSING = "Fehlender Export-Dateiname";
        public static string INVALID_FILEPATH = "Unzulässiger Dateipfad";
        public static string INVALID_FILENAME = "Unzulässiger Dateiname";
        public static string SETTINGS_FILE_ERROR = "Einstellungen-Datei-Fehler";
        public static string FILE_OPEN_ERROR = "Export konnte nicht durchgeführt werden. Bitte schließen Sie das Dokument.";
        public static string EXPORT_ERROR_CAPTION = "Export Fehler";

        public static string PORT = "Port";
        public static string PRINTER = "Drucker";
        public static string EXPORT_FILE = "ExportDatei";
        public static string AUTO_PRINT = "AutoDruck";
        public static string EXPORT = "Exportieren";
        public static string DATE_FORMAT = "DatumsFormat";

        public static string SELECT_PORT = "Port Auswählen: ";
        public static string PRINTER_NAME = "Drucker Name: ";
        public static string LANGUAGE = "Sprache";
        public static string CHOOSE = "Wählen";
        public static string EXPORT_TO_EXCEL = "Nach Excel exportieren";
        public static string EXPORT_TO_EXCEL_SETTINGS = "Nach Excel Einstellungen exportieren";
        public static string FILE_NAME = "Datei Name: ";
        public static string DATE_FORMAT_SET = "Datum Format: ";
        public static string OPEN = "Öffnen";
        public static string SAVE = "Speichern";
        public static string CANCEL = "Abbrechen";
        public static string PRINTED_FROM = "Gedruckt von";
        public static string REPORT_SETTINGS = "Report Einstellungen";
        public static string MAKE_PDF_REPORT = "Pdf Report erstellen";
        public static string PRINT_TEST_RESULTS = "Testergebnisse drucken";
        public static string ADD_HEADER_SPACE = "Kopfzeilenbereich hinzufügen";
        public static string HEADER_SPACE_UNIT = "mm";

        public static string REPORT_FILE_NAME = "Sperma Analyse Report";
        public static string TEST_REPORT = "SPERMA ANALYSE TEST REPORT"; // has to be same length as in English, so pdf generation can close document
        public static string REPORT_CONTINUED = "REPORT FORTSETZUNG";
        public static string DEVICE_INFO = "GERÄTEINFORMATION";
        public static string PATIENT_INFO = "PATIENTENINFORMATION";
        public static string SAMPLE_INFO = "PROBEINFORMATION";
        public static string PARAMETER = "PARAMETER";
        public static string RESULT = "ERGEBNIS";

    }
}
