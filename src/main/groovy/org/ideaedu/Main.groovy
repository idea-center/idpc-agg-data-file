package org.ideaedu

import groovyx.net.http.RESTClient
import groovyx.net.http.ContentType
import groovy.swing.SwingBuilder
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.codehaus.groovy.runtime.DateGroovyMethods

import java.awt.FlowLayout
import javax.swing.BoxLayout
import javax.swing.border.EmptyBorder





/**
 * The Main class provides a way to pull IDEA Feedback System data from the
 * IDEA Data Portal. In this case, it pulls data to create the aggregate data
 * file (a Microsoft Excel spreadsheet with 6 tabs). It has some required and
 * some optional command line arguments that control the behavior. The
 * arguments include:
 * <ul>
 * <li>i (institution) - the institution ID to get the data for.</li>
 * <li>s (start) - the start date to get data for.</li>
 * <li>e (end) - the end date to get data for.</li>
 * <li>h (host) - the hostname of the IDEA Data Portal</li>
 * <li>p (port) - the port to communicate to the IDEA Data Portal on</li>
 * <li>v (verbose) - provide more output on the command line</li>
 * <li>a (app) - the client application name</li>
 * <li>k (key) - the client application key</li>
 * <li>? (help) - show the usage of this</li>
 * </ul>
 *
 * @author Todd Wallentine todd AT IDEAedu org
 * @author Joonsuk Lee joon AT IDEAedu org
 */
public class Main {

    private static def verboseOutput = false

    //TODO Remove this hard-coded path
    private static final def template = "Data_Disk_Format.xlsx"
    private static final def out = "out.xlsx"

    private static final def DEFAULT_HOSTNAME = "localhost" //"rest.ideasystem.org"
    private static final def DEFAULT_PORT = 8091 //localhost port: 8091, production port: 443
    private static final def DEFAULT_BASE_PATH = "IDEA-REST-SERVER/v1"
    private static final def DEFAULT_AUTH_HEADERS = [ "X-IDEA-APPNAME": "", "X-IDEA-KEY": "" ]
    private static final def DEFAULT_PROTOCOL = "http" //local: http, production: https
    private static final def DEFAULT_IDENTITY_ID = 1029
    private static final def DEFAULT_START_DATE = getFormattedDate(new Date()-900 )
    private static final def DEFAULT_END_DATE = getFormattedDate(new Date()-1)
    private static final def DEFAULT_TYPE = "Diagnostic"

    private static final def FORM_NAMES = [
        9: "Short",
        10: "Diagnostic"
    ]

    private static final def OBJECTIVE_MAP = [
        571: "Obj1",
        572: "Obj2",
        573: "Obj3",
        574: "Obj4",
        575: "Obj5",
        576: "Obj6",
        577: "Obj7",
        578: "Obj8",
        579: "Obj9",
        580: "Obj10",
        581: "Obj11",
        582: "Obj12"
    ]

    private static final def TERM_CODE_MAP = [
        "Fall" : 1,
        "Winter" : 2,
        "Spring" : 3,
        "Summer" : 4,
        "Other" : 5
    ]
    private static final def DIAGNOSTIC_QUESTION_ID_LIST = [490..492,0,471..482].flatten() //0 is where Summary Evaluation gets output
    private static final def DIAGNOSTIC_QUESTION_TEACHING_METHOD_ID_LIST = [451..470].flatten()
    private static final def DIAGNOSTIC_ONLY_QUESTION_ID_LIST = [483..489,493].flatten()
    private static final def DIAGNOSTIC_ADDITIONAL_QUESTION_ID_LIST = [546..565].flatten()
    private static final def SHORT_QUESTION_ID_LIST = [716..718,0,701..712].flatten() ////0 is where Summary Evaluation gets output
    private static final def RESEARCH_QUESTION_ID_LIST = [695,494..497].flatten()
    private static final def SHORT_ADDITIONAL_QUESTION_ID_LIST = [546..565].flatten()

    /** The maximum number of surveys to get before quitting. */
    private static final def MAX_SURVEYS = 10 //TODO: set to 1 for test

    /** The number of surveys to get per page */
    private static final def PAGE_SIZE = 10 //TODO: set to 1 for test

    private static def identityID = DEFAULT_IDENTITY_ID
    private static def startDate = DEFAULT_START_DATE
    private static def endDate = DEFAULT_END_DATE
    private static def hostname = DEFAULT_HOSTNAME
    private static def protocol = DEFAULT_PROTOCOL
    private static def port = DEFAULT_PORT
    private static def basePath = DEFAULT_BASE_PATH
    private static def authHeaders = DEFAULT_AUTH_HEADERS

    private static RESTClient restClient

    private static def institutionList = []

    public static void main(String[] args) {

        def cli = new CliBuilder( usage: 'Main -v -h host -p port -a "TestClient" -k "ABCDEFG123456"' )
        cli.with {
            v longOpt: 'verbose', 'verbose output'
            i longOpt: 'institution', 'institution ID', args:1
            s longOpt: 'start', 'start date', args:1
            e longOpt: 'end', 'end date', args:1
            h longOpt: 'host', 'host name (default: rest.ideasystem.org)', args:1
            p longOpt: 'port', 'port number (default: 443)', args:1
            a longOpt: 'app', 'client application name', args:1
            k longOpt: 'key', 'client application key', args:1
            '?' longOpt: 'help', 'help'
        }
        def options = cli.parse(args)
        if(options.'?') {
            cli.usage()
            return
        }
        if(options.v) {
            verboseOutput = true
        }
        if(options.i) {
            identityID = options.i.toInteger()
        }
        if(options.s) {
            startDate = options.s // TODO Get the date
        }
        if(options.e) {
            endDate = options.e // TODO Get the date
        }
        if(options.h) {
            hostname = options.h
        }
        if(options.p) {
            port = options.p.toInteger()
        }
        if(options.a) {
            authHeaders['X-IDEA-APPNAME'] = options.a
        }
        if(options.k) {
            authHeaders['X-IDEA-KEY'] = options.k
        }

        /**
         * Creates a simple UI to select institution and start/end date
         */
        def institutionComboBox
        def startDateTextField
        def endDateTextField
        institutionList = getAllInstitutions()
        if (verboseOutput) println("institutionList= ${institutionList}")
        new SwingBuilder().edt {
            frame(title: 'Combo Aggregate Data File Generator', size: [600, 200], show: true,  defaultCloseOperation:javax.swing.WindowConstants.EXIT_ON_CLOSE) {
                panel(border:new EmptyBorder(2,2,2,2)) {
                    boxLayout(axis:BoxLayout.Y_AXIS)
                    panel(layout:new FlowLayout()){
                        label(text: 'Institution')
                        institutionComboBox = comboBox()
                        institutionList.each { institution ->
                            institutionComboBox.addItem("${institution.name} | ${institution.id}")
                        }
                    }
                    panel(layout:new FlowLayout()){
                        label(text: 'Start Date')
                        startDateTextField = textField(text:startDate) //TODO Format checking is required
                    }
                    panel(layout:new FlowLayout()){
                        label(text: 'End Date')
                        endDateTextField = textField(text:endDate) //TODO Format checking is required
                    }
                    panel(layout:new FlowLayout()){
                        button(label: 'Generate', actionPerformed: {
                            identityID = institutionComboBox?.getSelectedItem()?.split("\\|")[1]?.trim()
                            startDate = startDateTextField.text
                            endDate = endDateTextField.text
                            def successfullyGenerated = generate()
                            //optionPane().showMessageDialog(null, successfullyGenerated, "Result", JOptionPane.INFORMATION_MESSAGE) //Optional pop-up message
                        })
                    }
                }
            }
        }
    }

    /**
     * Generate and outputs aggregate data given institution id, start / end date
     * @return
     */
    private static def generate(pb){
        /*
        * The following will get all the surveys that are available of the
        * given type and print out the overall ratings for each survey subject.
        * This will print the raw and adjusted mean and t-score for each survey
        * subject.
        */
        if (verboseOutput) println("ID=${identityID}, startDate=${startDate}, endDate=${endDate}")
        def successfullyGenerated = false

        def types = [ 9, 10 ]
        def institution = getInstitution(institutionList, identityID)
        def surveys = getAllSurveys(identityID, types, startDate, endDate)
        def disciplines = getDisciplines()
        def wb = WorkbookFactory.create(new FileInputStream(new File(template))) //Get the Excel template
        if(surveys) {
            // Print the CSV header
            if (verboseOutput) "ID, FICE, Institution, Term, Year, Instructor, Dept_Code_Name, Course_Num, Dept_Name, Dept_Code, Local_Code, Time, Days, Enrolled, Responses, Form, Delivery, Batch, " +
                    "Obj1, Obj2, Obj3, Obj4, Obj5, Obj6, Obj7, Obj8, Obj9, Obj10, Obj11, Obj12, " +
                    "PRO_Raw_Mean,PRO_Adj_Mean,PRO_CRaw_IDEA,PRO_CAdj_IDEA,PRO_CRaw_Disc,PRO_CAdj_Disc,PRO_CRaw_Inst,PRO_CAdj_Inst,Impr_Stu_Att_Raw_Mean,Impr_Stu_Att_Adj_Mean,Impr_Stu_Att_CRaw_IDEA,Impr_Stu_Att_CAdj_IDEA,Exc_Tchr_Raw_Mean,Exc_Tchr_Adj_Mean,Exc_Tchr_CRaw_IDEA,Exc_Tchr_CAdj_IDEA,Exc_Tchr_CRaw_Disc,Exc_Tchr_CAdj_Disc,Exc_Tchr_CRaw_Inst,Exc_Tchr_CAdj_Inst,Exc_Crs_Raw_Mean,Exc_Crs_Adj_Mean,Exc_Crs_CRaw_IDEA,Exc_Crs_CAdj_IDEA,Exc_Crs_CRaw_Disc,Exc_Crs_CAdj_Disc,Exc_Crs_CRaw_Inst,Exc_Crs_CAdj_Inst,SumEval_Raw_Mean,SumEval_Adj_Mean,SumEval_CRaw_IDEA,SumEval_CAdj_IDEA,SumEval_CRaw_Disc,SumEval_CAdj_Disc,SumEval_CRaw_Inst,SumEval_CAdj_Inst," +
                    "Obj1,Obj1_Raw_Mean,Obj1_Adj_Mean,Obj1_CRaw_IDEA,Obj1_CAdj_IDEA,Obj1_CRaw_Disc,Obj1_CAdj_Disc,Obj1_CRaw_Inst,Obj1_CAdj_Inst,Obj2,Obj2_Raw_Mean,Obj2_Adj_Mean,Obj2_CRaw_IDEA,Obj2_CAdj_IDEA,Obj2_CRaw_Disc,Obj2_CAdj_Disc,Obj2_CRaw_Inst,Obj2_CAdj_Inst,Obj3,Obj3_Raw_Mean,Obj3_Adj_Mean,Obj3_CRaw_IDEA,Obj3_CAdj_IDEA,Obj3_CRaw_Disc,Obj3_CAdj_Disc,Obj3_CRaw_Inst,Obj3_CAdj_Inst,Obj4,Obj4_Raw_Mean,Obj4_Adj_Mean,Obj4_CRaw_IDEA,Obj4_CAdj_IDEA,Obj4_CRaw_Disc,Obj4_CAdj_Disc,Obj4_CRaw_Inst,Obj4_CAdj_Inst,Obj5,Obj5_Raw_Mean,Obj5_Adj_Mean,Obj5_CRaw_IDEA,Obj5_CAdj_IDEA,Obj5_CRaw_Disc,Obj5_CAdj_Disc,Obj5_CRaw_Inst,Obj5_CAdj_Inst,Obj6,Obj6_Raw_Mean,Obj6_Adj_Mean,Obj6_CRaw_IDEA,Obj6_CAdj_IDEA,Obj6_CRaw_Disc,Obj6_CAdj_Disc,Obj6_CRaw_Inst,Obj6_CAdj_Inst,Obj7,Obj7_Raw_Mean,Obj7_Adj_Mean,Obj7_CRaw_IDEA,Obj7_CAdj_IDEA,Obj7_CRaw_Disc,Obj7_CAdj_Disc,Obj7_CRaw_Inst,Obj7_CAdj_Inst,Obj8,Obj8_Raw_Mean,Obj8_Adj_Mean,Obj8_CRaw_IDEA,Obj8_CAdj_IDEA,Obj8_CRaw_Disc,Obj8_CAdj_Disc,Obj8_CRaw_Inst,Obj8_CAdj_Inst,Obj9,Obj9_Raw_Mean,Obj9_Adj_Mean,Obj9_CRaw_IDEA,Obj9_CAdj_IDEA,Obj9_CRaw_Disc,Obj9_CAdj_Disc,Obj9_CRaw_Inst,Obj9_CAdj_Inst,Obj10,Obj10_Raw_Mean,Obj10_Adj_Mean,Obj10_CRaw_IDEA,Obj10_CAdj_IDEA,Obj10_CRaw_Disc,Obj10_CAdj_Disc,Obj10_CRaw_Inst,Obj10_CAdj_Inst,Obj11,Obj11_Raw_Mean,Obj11_Adj_Mean,Obj11_CRaw_IDEA,Obj11_CAdj_IDEA,Obj11_CRaw_Disc,Obj11_CAdj_Disc,Obj11_CRaw_Inst,Obj11_CAdj_Inst,Obj12,Obj12_Raw_Mean,Obj12_Adj_Mean,Obj12_CRaw_IDEA,Obj12_CAdj_IDEA,Obj12_CRaw_Disc,Obj12_CAdj_Disc,Obj12_CRaw_Inst,Obj12_CAdj_Inst," +
                    "Method1_Mean,Method2_Mean,Method3_Mean,Method4_Mean,Method5_Mean,Method6_Mean,Method7_Mean,Method8_Mean,Method9_Mean,Method10_Mean,Method11_Mean,Method12_Mean,Method13_Mean,Method14_Mean,Method15_Mean,Method16_Mean,Method17_Mean,Method18_Mean,Method19_Mean,Method20_Mean,Read_Mean,Read_Cnv_IDEA," +
                    "Read_Cnv_Disc,Read_Cnv_Inst,NonRead_Mean,NonRead_Cnv_IDEA,NonRead_Cnv_Disc,NonRead_Cnv_Inst,Diff_Mean,Diff_Cnv_IDEA,Diff_Cnv_Disc,Diff_Cnv_Inst,Q36_Mean,Effort_Mean,Effort_Cnv_IDEA,Effort_Cnv_Disc,Effort_Cnv_Inst,Q38_Mean,Motivation_Mean,Motivation_Cnv_IDEA,Motivation_Cnv_Disc,Motivation_Cnv_Inst,WkHabit_Mean,WkHabit_Cnv_IDEA,WkHabit_Cnv_Disc,WkHabit_Cnv_Inst,Background_Mean,Q44_Mean,Q45_Mean,Q46_Mean,Q47_Mean,Add_Q1_Mean,Add_Q2_Mean,Add_Q3_Mean,Add_Q4_Mean,Add_Q5_Mean,Add_Q6_Mean,Add_Q7_Mean,Add_Q8_Mean,Add_Q9_Mean,Add_Q10_Mean,Add_Q11_Mean,Add_Q12_Mean,Add_Q13_Mean,Add_Q14_Mean,Add_Q15_Mean,Add_Q16_Mean,Add_Q17_Mean,Add_Q18_Mean,Add_Q19_Mean,Add_Q20_Mean," +
                    "Impr_Stu_Att_1,Impr_Stu_Att_2,Impr_Stu_Att_3,Impr_Stu_Att_4,Impr_Stu_Att_5,Impr_Stu_Att_Omit,Exc_Tchr_Inst_1,Exc_Tchr_Inst_2,Exc_Tchr_Inst_3,Exc_Tchr_Inst_4,Exc_Tchr_Inst_5,Exc_Tchr_Inst_Omit,Exc_Crs_1,Exc_Crs_2,Exc_Crs_3,Exc_Crs_4,Exc_Crs_5,Exc_Crs_Omit,Obj1_1,Obj1_2,Obj1_3,Obj1_4,Obj1_5,Obj1_Omit,Obj2_1,Obj2_2,Obj2_3,Obj2_4,Obj2_5,Obj2_Omit,Obj3_1,Obj3_2,Obj3_3,Obj3_4,Obj3_5,Obj3_Omit,Obj4_1,Obj4_2,Obj4_3,Obj4_4,Obj4_5,Obj4_Omit,Obj5_1,Obj5_2,Obj5_3,Obj5_4,Obj5_5,Obj5_Omit,Obj6_1,Obj6_2,Obj6_3,Obj6_4,Obj6_5,Obj6_Omit,Obj7_1,Obj7_2,Obj7_3,Obj7_4,Obj7_5,Obj7_Omit,Obj8_1,Obj8_2,Obj8_3,Obj8_4,Obj8_5,Obj8_Omit,Obj9_1,Obj9_2,Obj9_3,Obj9_4,Obj9_5,Obj9_Omit,Obj10_1,Obj10_2,Obj10_3,Obj10_4,Obj10_5,Obj10_Omit,Obj11_1,Obj11_2,Obj11_3,Obj11_4,Obj11_5,Obj11_Omit,Obj12_1,Obj12_2,Obj12_3,Obj12_4,Obj12_5,Obj12_Omit,Method1_1,Method1_2,Method1_3,Method1_4,Method1_5,Method1_Omit,Method2_1,Method2_2,Method2_3,Method2_4,Method2_5,Method2_Omit,Method3_1,Method3_2,Method3_3,Method3_4,Method3_5,Method3_Omit,Method4_1,Method4_2,Method4_3,Method4_4,Method4_5,Method4_Omit,Method5_1,Method5_2,Method5_3,Method5_4,Method5_5,Method5_Omit,Method6_1,Method6_2,Method6_3,Method6_4,Method6_5,Method6_Omit,Method7_1,Method7_2,Method7_3,Method7_4,Method7_5,Method7_Omit,Method8_1,Method8_2,Method8_3,Method8_4,Method8_5,Method8_Omit,Method9_1,Method9_2,Method9_3,Method9_4,Method9_5,Method9_Omit,Method10_1,Method10_2,Method10_3,Method10_4,Method10_5,Method10_Omit,Method11_1,Method11_2,Method11_3,Method11_4,Method11_5,Method11_Omit,Method12_1,Method12_2,Method12_3,Method12_4,Method12_5,Method12_Omit,Method13_1,Method13_2,Method13_3,Method13_4,Method13_5,Method13_Omit,Method14_1,Method14_2,Method14_3,Method14_4,Method14_5,Method14_Omit,Method15_1,Method15_2,Method15_3,Method15_4,Method15_5,Method15_Omit,Method16_1,Method16_2,Method16_3,Method16_4,Method16_5,Method16_Omit,Method17_1,Method17_2,Method17_3,Method17_4,Method17_5,Method17_Omit,Method18_1,Method18_2,Method18_3,Method18_4,Method18_5,Method18_Omit,Method19_1,Method19_2,Method19_3,Method19_4,Method19_5,Method19_Omit,Method20_1,Method20_2,Method20_3,Method20_4,Method20_5,Method20_Omit,Reading_1,Reading_2,Reading_3,Reading_4,Reading_5,Reading_Omit,NonRead_1,NonRead_2,NonRead_3,NonRead_4,NonRead_5,NonRead_Omit,Diff_1,Diff_2,Diff_3,Diff_4,Diff_5,Diff_Omit,Effort_1,Effort_2,Effort_3,Effort_4,Effort_5,Effort_Omit,Motivation_1,Motivation_2,Motivation_3,Motivation_4,Motivation_5,Motivation_Omit,WkHabit_1,WkHabit_2,WkHabit_3,WkHabit_4,WkHabit_5,WkHabit_Omit," +
                    "Q36_1,Q36_2,Q36_3,Q36_4,Q36_5,Q36_Omit,Q38_1,Q38_2,Q38_3,Q38_4,Q38_5,Q38_Omit,Background_1,Background_2,Background_3,Background_4,Background_5,Background_Omit,Q44_1,Q44_2,Q44_3,Q44_4,Q44_5,Q44_Omit,Q45_1,Q45_2,Q45_3,Q45_4,Q45_5,Q45_Omit,Q46_1,Q46_2,Q46_3,Q46_4,Q46_5,Q46_Omit,Q47_1,Q47_2,Q47_3,Q47_4,Q47_5,Q47_Omit,Add_Q1_1,Add_Q1_2,Add_Q1_3,Add_Q1_4,Add_Q1_5,Add_Q1_Omit,Add_Q2_1,Add_Q2_2,Add_Q2_3,Add_Q2_4,Add_Q2_5,Add_Q2_Omit,Add_Q3_1,Add_Q3_2,Add_Q3_3,Add_Q3_4,Add_Q3_5,Add_Q3_Omit,Add_Q4_1,Add_Q4_2,Add_Q4_3,Add_Q4_4,Add_Q4_5,Add_Q4_Omit,Add_Q5_1,Add_Q5_2,Add_Q5_3,Add_Q5_4,Add_Q5_5,Add_Q5_Omit,Add_Q6_1,Add_Q6_2,Add_Q6_3,Add_Q6_4,Add_Q6_5,Add_Q6_Omit,Add_Q7_1,Add_Q7_2,Add_Q7_3,Add_Q7_4,Add_Q7_5,Add_Q7_Omit,Add_Q8_1,Add_Q8_2,Add_Q8_3,Add_Q8_4,Add_Q8_5,Add_Q8_Omit,Add_Q9_1,Add_Q9_2,Add_Q9_3,Add_Q9_4,Add_Q9_5,Add_Q9_Omit,Add_Q10_1,Add_Q10_2,Add_Q10_3,Add_Q10_4,Add_Q10_5,Add_Q10_Omit,Add_Q11_1,Add_Q11_2,Add_Q11_3,Add_Q11_4,Add_Q11_5,Add_Q11_Omit,Add_Q12_1,Add_Q12_2,Add_Q12_3,Add_Q12_4,Add_Q12_5,Add_Q12_Omit,Add_Q13_1,Add_Q13_2,Add_Q13_3,Add_Q13_4,Add_Q13_5,Add_Q13_Omit,Add_Q14_1,Add_Q14_2,Add_Q14_3,Add_Q14_4,Add_Q14_5,Add_Q14_Omit,Add_Q15_1,Add_Q15_2,Add_Q15_3,Add_Q15_4,Add_Q15_5,Add_Q15_Omit,Add_Q16_1,Add_Q16_2,Add_Q16_3,Add_Q16_4,Add_Q16_5,Add_Q16_Omit,Add_Q17_1,Add_Q17_2,Add_Q17_3,Add_Q17_4,Add_Q17_5,Add_Q17_Omit,Add_Q18_1,Add_Q18_2,Add_Q18_3,Add_Q18_4,Add_Q18_5,Add_Q18_Omit,Add_Q19_1,Add_Q19_2,Add_Q19_3,Add_Q19_4,Add_Q19_5,Add_Q19_Omit,Add_Q20_1,Add_Q20_2,Add_Q20_3,Add_Q20_4,Add_Q20_5,Add_Q20_Omit"

            println("# of surveys found: ${surveys.size()}")
            surveys.eachWithIndex { survey, pbIndex ->
                println ("processing ${pbIndex+1} of ${surveys.size()} surveys")
                //Lists that store data used for excel output
                def fifDataList = [] //sheet 0
                def identificationFieldDataList = [] //identification fields used in sheet 1-3
                def meansDataList = []
                def frequenciesPrimaryItemsDataList = []
                def frequenciesResearchAdditonalQuestionDataList = []

                def surveySubject = survey.info_form.respondents[0]
                def formName = getFormName(survey.rater_form.id)
                def discipline = getDiscipline(disciplines, survey.info_form.discipline_code)
                def reports = getReports(survey.id, "Short")
                def questionList = SHORT_QUESTION_ID_LIST
                def additionalQuestionList = SHORT_ADDITIONAL_QUESTION_ID_LIST
                if(!reports) {
                    reports = getReports(survey.id, "Diagnostic")
                    questionList = DIAGNOSTIC_QUESTION_ID_LIST
                    additionalQuestionList = DIAGNOSTIC_ADDITIONAL_QUESTION_ID_LIST
                }
                if (verboseOutput) ("questionList ${questionList}")
                def reportID = reports[0]?.id
                if(reportID) {
                    def model = getReportModel(reportID)
                    def frequencies_map = [:] //Frequencies will be stored
                    def objectives = [] //objectives will be stored

                    def rowIndex = pbIndex + 3 //Start writing from row #4 in Excel template

                    if (verboseOutput) {
                        print "${survey.id},"
                        print "${institution?.fice},"
                        print "${institution?.name},"
                        print "${survey.term}," //e.g. Fall 2013
                        print "${surveySubject.first_name} ${surveySubject.last_name},"
                        print "${survey.info_form.discipline_code},"
                        print "${survey.course.number},"
                        print "${discipline?.name},"
                        print "${discipline?.abbrev},"
                        print "${survey.course.local_code}," // local code unused in IDEA-CL
                        print "${survey.course.time}," // course time is unused in IDEA-CL?
                        print "${survey.course.days}," // course days is unused in IDEA-CL?
                        print "${model.aggregate_data.asked},"
                        print "${model.aggregate_data.answered},"
                        print "${formName},"
                        print "," // Skip delivery
                        print "," // batch is unused in IDEA-CL

                        //TODO Number formatting (null, 1/10 for means and 1/1000 for t-scores)?
                        /* Progress on Relevant Objectives */
                        //Means
                        print "${model.aggregate_data.relevant_results.result.raw.mean},"                       //PRO_Raw_Mean
                        print "${model.aggregate_data.relevant_results.result.adjusted.mean},"                  //PRO_Adj_Mean
                        //Compared to IDEA
                        print "${model.aggregate_data.relevant_results.result.raw.tscore},"                     //PRO_CRaw_IDEA
                        print "${model.aggregate_data.relevant_results.result.adjusted.tscore},"                //PRO_CAdj_IDEA
                        //Compared to Discipline
                        print "${model.aggregate_data.relevant_results.discipline_result.raw.tscore},"          //PRO_CRaw_Disc
                        print "${model.aggregate_data.relevant_results.discipline_result.adjusted.tscore},"     //PRO_CAdj_Disc
                        //Compared to Institution
                        print "${model.aggregate_data.relevant_results.institution_result.raw.tscore},"         //PRO_CRaw_Inst
                        print "${model.aggregate_data.relevant_results.institution_result.adjusted.tscore},"    //PRO_CAdj_Inst
                    }
                    fifDataList.add(survey.id)
                    fifDataList.add(institution?.fice)
                    fifDataList.add(institution?.name)
                    def term
                    def year
                    if (survey.term){
                        if (survey.term.split(" ").size() == 2){
                            term = survey.term?.split(" ")[0]
                            year = survey.term?.split(" ")[1]
                        }
                        else term = survey.term
                    }
                    fifDataList.add(term) //term
                    fifDataList.add(year) //year
                    fifDataList.add(surveySubject.first_name+ " " +surveySubject.last_name)
                    fifDataList.add(survey.info_form.discipline_code)
                    fifDataList.add(survey.course.number)
                    fifDataList.add(discipline?.name)
                    fifDataList.add(discipline?.abbrev)
                    fifDataList.add(survey.course.local_code)
                    fifDataList.add(survey.course.time)
                    fifDataList.add(survey.course.days)
                    fifDataList.add(model?.aggregate_data?.asked)
                    fifDataList.add(model?.aggregate_data?.answered)
                    fifDataList.add(formName)
                    fifDataList.add("") //skip delivery
                    fifDataList.add("") //skip batch

                    //Objectives
                    OBJECTIVE_MAP.eachWithIndex { objective, index ->
                        def response = "Default-Imp"
                        if(model){
                            model.aggregate_data.relevant_results.questions.each { question ->
                                if(question.question_id == objective.key) {
                                    response = question.response
                                }
                            }
                        }
                        objectives[index] = response //This is ok as we always have 12 objectives
                        fifDataList.add(response)
                    }

                    /** Means Sheet*/
                    identificationFieldDataList.add(survey.id)
                    identificationFieldDataList.add(term) //term
                    identificationFieldDataList.add(year) //year
                    identificationFieldDataList.add(surveySubject.first_name+ " " +surveySubject.last_name)
                    identificationFieldDataList.add(survey.info_form.discipline_code)
                    identificationFieldDataList.add(survey.course.number)
                    identificationFieldDataList.add(survey.course.local_code)
                    identificationFieldDataList.add(model?.aggregate_data?.asked)
                    identificationFieldDataList.add(model?.aggregate_data?.answered)
                    identificationFieldDataList.add(formName)
                    identificationFieldDataList.add("") //skip delivery

                    /* Progress on Relevant Objectives */
                    //Means
                    meansDataList.add(model?.aggregate_data?.relevant_results?.result?.raw?.mean)
                    meansDataList.add(model?.aggregate_data?.relevant_results?.result?.adjusted?.mean)

                    //Compared to IDEA
                    meansDataList.add(model?.aggregate_data?.relevant_results?.result?.raw?.tscore)
                    meansDataList.add(model?.aggregate_data?.relevant_results?.result?.adjusted?.tscore)

                    //Compared to Discipline
                    meansDataList.add(model?.aggregate_data?.relevant_results?.discipline_result?.raw?.tscore)
                    meansDataList.add(model?.aggregate_data?.relevant_results?.discipline_result?.adjusted?.tscore)

                    //Compared to Institution
                    meansDataList.add(model?.aggregate_data?.relevant_results?.institution_result?.raw?.tscore)
                    meansDataList.add(model?.aggregate_data?.relevant_results?.institution_result?.adjusted?.tscore)

                    /* As a result of taking this course, I have more positive feelings toward this field of study --- Diagnostic - 40 and Short - 16 */
                    /* Excellent teacher: Overall, I rate this instructor an excellent teacher --- Diagnostic - 41 and Short - 17 */
                    /* Excellent course: Overall, I rate this course as excellent --- Diagnostic - 42 and Short - 18 */

                    /* Each index represents a section in Means
                       index = 0 : As a result of taking this course, I have more positive feelings toward this field of study --- Diagnostic - 40 and Short - 16
                       index = 1 : Overall, I rate this instructor an excellent teacher --- Diagnostic - 41 and Short - 17
                       index = 2 : Overall, I rate this course as excellent --- Diagnostic - 42 and Short - 18
                       index = 3 : Summary Evaluation (where questionID = 0)
                       index = 4 : Objective 1 - Gaining factual knowledge (terminology, classifications, methods, trends) --- Diagnostic - 21 and Short - 1
                       ....
                       index = 15 : Objective 12 - Acquiring an interest in learning more by asking my own questions and seeking answers --- Diagnostic - 32 and Short - 12
                    */
                    questionList.eachWithIndex{ questionID, index ->
                        def reportData = getReportDataByQuestion(reportID, questionID, frequencies_map)

                        //Add Summary before Objective 1 (Diagnostic 21 / Short 1)
                        if (index == 3){
                            if (verboseOutput){
                                print "${model?.aggregate_data?.summary_evaluation.result.raw.mean},"                     //SumEval_Raw_Mean
                                print "${model?.aggregate_data?.summary_evaluation.result.adjusted.mean},"                //SumEval_Adj_Mean
                                print "${model?.aggregate_data?.summary_evaluation.result.raw.tscore},"                   //SumEval_CRaw_IDEA
                                print "${model?.aggregate_data?.summary_evaluation.result.adjusted.tscore},"              //SumEval_CAdj_IDEA
                                print "${model?.aggregate_data?.summary_evaluation.discipline_result.raw.tscore},"        //SumEval_CRaw_Disc
                                print "${model?.aggregate_data?.summary_evaluation.discipline_result.adjusted.tscore},"   //SumEval_CAdj_Disc
                                print "${model?.aggregate_data?.summary_evaluation.institution_result.raw.tscore},"       //SumEval_CRaw_Inst
                                print "${model?.aggregate_data?.summary_evaluation.institution_result.adjusted.tscore},"  //SumEval_CAdj_Inst
                            }

                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.result?.raw?.mean)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.result?.adjusted?.mean)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.result?.raw?.tscore)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.result?.adjusted?.tscore)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.discipline_result?.raw?.tscore)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.discipline_result?.adjusted?.tscore)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.institution_result?.raw?.tscore)
                            meansDataList.add(model?.aggregate_data?.summary_evaluation?.institution_result?.adjusted?.tscore)

                        }else{
                            //Print Obj1-12
                            if (index >= 4 && index <= 15) {
                                if (verboseOutput) print "${objectives[index-4]}," //e.g., objectives[0] = obj1
                                meansDataList.add(objectives[index-4])
                            }
                            if (verboseOutput){
                                print "${reportData?.results?.result?.raw?.mean},"          //raw_mean
                                print "${reportData?.results?.result?.adjusted?.mean},"     //adj_mean
                                print "${reportData?.results?.result?.raw?.tscore},"        //raw t-score
                                print "${reportData?.results?.result?.adjusted?.tscore},"   //raw adjusted t-score
                            }
                            meansDataList.add(reportData?.results?.result?.raw?.mean)
                            meansDataList.add(reportData?.results?.result?.adjusted?.mean)
                            meansDataList.add(reportData?.results?.result?.raw?.tscore)
                            meansDataList.add(reportData?.results?.result?.adjusted?.tscore)

                            //Except Diagnostic 40 / Short 16 (index = 0), we need 'compared to discipline / institution' data
                            if (index != 0){
                                if (verboseOutput){
                                    print "${reportData?.results?.discipline_result?.raw?.tscore},"         //discipline raw t-score
                                    print "${reportData?.results?.discipline_result?.adjusted?.tscore},"    //discipline adjusted t-score
                                    print "${reportData?.results?.institution_result?.raw?.tscore},"        //institution raw t-score
                                    print "${reportData?.results?.institution_result?.adjusted?.tscore},"   //institution adjusted t-score
                                }
                                meansDataList.add(reportData?.results?.discipline_result?.raw?.tscore)
                                meansDataList.add(reportData?.results?.discipline_result?.adjusted?.tscore)
                                meansDataList.add(reportData?.results?.institution_result?.raw?.tscore)
                                meansDataList.add(reportData?.results?.institution_result?.adjusted?.tscore)
                            }
                        }
                    }

                    //Teaching Method 1-20 (Diagnostic 1-20)
                    DIAGNOSTIC_QUESTION_TEACHING_METHOD_ID_LIST.each { questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID, frequencies_map)
                        if (verboseOutput) "${reportData?.results?.result?.raw?.mean},"
                        meansDataList.add(reportData?.results?.result?.raw?.mean)
                    }

                    //Disgnostic 33-40
                    DIAGNOSTIC_ONLY_QUESTION_ID_LIST.each{ questionID ->
                        if (verboseOutput) ("questionID= ${questionID}")
                        def reportData = getReportDataByQuestion(reportID, questionID, frequencies_map)
                        if (verboseOutput) ("${reportData?.results?.result?.raw?.mean},")          //raw_mean
                        meansDataList.add(reportData?.results?.result?.raw?.mean)

                        //Diagnostic 36, 38 (QUESTION_ID = 486, 488) only requires mean
                        if (questionID != 486 && questionID != 488){
                            if (verboseOutput){
                                print "${reportData?.results?.result?.raw?.tscore},"                    //raw t-score
                                print "${reportData?.results?.discipline_result?.raw?.tscore},"        //discipline raw t-score
                                print "${reportData?.results?.institution_result?.raw?.tscore},"       //institution raw t-score
                            }
                            meansDataList.add(reportData?.results?.result?.raw?.tscore)
                            meansDataList.add(reportData?.results?.discipline_result?.raw?.tscore)
                            meansDataList.add(reportData?.results?.institution_result?.raw?.tscore)
                        }
                    }

                    /* Research Questions */
                    RESEARCH_QUESTION_ID_LIST.each{ questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID, frequencies_map)
                        if (verboseOutput) "${reportData?.results?.result?.raw?.mean},"          //raw_mean
                        meansDataList.add(reportData?.results?.result?.raw?.mean)
                    }

                    //Additional Questions
                    additionalQuestionList.each{ questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID, frequencies_map)
                        if (verboseOutput) print "${reportData?.results?.result?.raw?.mean},"          //raw_mean
                        meansDataList.add(reportData?.results?.result?.raw?.mean)
                    }

                    /* Frequencies-Primary Items */
                    (questionList + DIAGNOSTIC_QUESTION_TEACHING_METHOD_ID_LIST + (DIAGNOSTIC_ONLY_QUESTION_ID_LIST - [486,488])).each{ questionID ->
                        frequencies_map.get(questionID).each{ count ->
                            if (verboseOutput) print ("${count},")
                            frequenciesPrimaryItemsDataList.add(count)
                        }
                    }

                    /* Frequencies-Research, Add Q's */
                    ([486,488] + RESEARCH_QUESTION_ID_LIST+ additionalQuestionList).each{ questionID ->
                        def freq_map = frequencies_map.get(questionID)
                        if (freq_map) {
                            frequencies_map.get(questionID).each{ count ->
                                if (verboseOutput) ("${count},")
                                frequenciesResearchAdditonalQuestionDataList.add(count)
                            }
                        }else{
                            //if the map for the questionID doesn't return anything, fill it with blank
                            0.upto(5, {
                                if (verboseOutput) print (",")
                                frequenciesResearchAdditonalQuestionDataList.add("")
                            })
                        }
                        writeDataToExcel(wb.getSheetAt(0), fifDataList, rowIndex)
                        writeDataToExcel(wb.getSheetAt(1), identificationFieldDataList+meansDataList, rowIndex)
                        writeDataToExcel(wb.getSheetAt(2), identificationFieldDataList+frequenciesPrimaryItemsDataList, rowIndex)
                        writeDataToExcel(wb.getSheetAt(3), identificationFieldDataList+frequenciesResearchAdditonalQuestionDataList, rowIndex)
                    }
                    if (verboseOutput) println ""
                }
            }
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(new File(out))
            wb.write(fileOut)
            fileOut.close()
            successfullyGenerated = true
            println("done")
        } else {
            println "No surveys are available."
        }
    }

    /**
     * Get make a discipline list
     * @return disciplines all disciplines
     */
    static def getDisciplines() {
        def disciplines = []
        def client = getRESTClient()
        def response = client.get(
            path: "${basePath}/disciplines/",
            query: [max: 0], //Returns all disciplines
            requestContentType: ContentType.JSON,
            headers: authHeaders)
        if(response.status == 200) {
            if(verboseOutput) {
                println "Discipline data: ${response.data}"
            }
            if(response.data) {
                response.data.data.each { discipline ->
                    disciplines << discipline
                }
            }
        } else {
            println "An error occured while getting the disciplines: ${response.status}"
        }
        disciplines
    }

    /**
     * Get discipline data based on disc code from survey
     * @param disciplines Disciplines
     * @param code disc code
     * @return discipline data that contains name and abbrev
     */
    static def getDiscipline(disciplines, code){
        def discipline
        disciplines.each{ disc ->
            if (disc.code==code){
                discipline = disc
            }
        }
        discipline

    }

    /**
     * Get the institution information that has the given ID.
     *
     * @param institutionList institutions
     * @param identityID The ID of the institution.
     * @return The institution with the given ID.
     */
    static def getInstitution(institutionList, identityID) {
        def institution
        institutionList.each {inst ->
            if (inst.id == identityID as Integer) institution = inst
        }
        institution
    }

    static def getReports(surveyID, type=DEFAULT_TYPE) {
        def reports = []

        def client = getRESTClient()
        def response = client.get(
            path: "${basePath}/reports",
            requestContentType: ContentType.JSON,
            query: [survey_id: surveyID, type: type],
            headers: authHeaders)
        if(response.status == 200) {
            if(verboseOutput) {
                println "Reports data: ${response.data}"
            }
            reports = response.data.data
        } else {
            println "An error occured while getting the reports with survey ID ${surveyID} and type ${type}: ${response.status}"
        }

        reports
    }

    static def getFormName(id) {
        return FORM_NAMES[id]
    }

    /**
     * Get report model based on report id
     * @param reportID report id
     * @return Report model
     */
    static def getReportModel(reportID) {
        def reportModel

        def client = getRESTClient()
        def response = client.get(
            path: "${basePath}/report/${reportID}/model",
            requestContentType: ContentType.JSON,
            headers: authHeaders)
        if(response.status == 200) {
            if(verboseOutput) {
                println "Report model data: ${response.data}"
            }
            reportModel = response.data
        } else {
            println "Report model doesn't exist for Report ID ${reportID} | Status:${response.status}"
        }
        reportModel
    }

    /**
     * Get report data based on report id and question id
     * @param reportID Report id
     * @param questionID Question id
     * @param frequencies_map Map stores frequency counts
     * @return Report data
     */
    static def getReportDataByQuestion(reportID, questionID, frequencies_map) {
        if (verboseOutput) println ("reportID= ${reportID}, questionID=${questionID}")
        def reportData
        def client = getRESTClient()
        def response = client.get(
                path: "${basePath}/report/${reportID}/model/${questionID}",
                requestContentType: ContentType.JSON,
                headers: authHeaders)
        if(response.status == 200) {
            if(verboseOutput) {
                println "Report  data: ${response.data}"
            }
            reportData = response.data

            //Puts frequencies in the map
            frequencies_map.put(questionID, [reportData.response_option_data_map."0".count, //omit
                                             reportData.response_option_data_map."1".count,
                                             reportData.response_option_data_map."2".count,
                                             reportData.response_option_data_map."3".count,
                                             reportData.response_option_data_map."4".count,
                                             reportData.response_option_data_map."5".count,
                                            ])
        } else {
//            //Puts frequencies with blank in the map if the questionID doesn't have any data
//            frequencies_map.put(questionID, ["","","","","","",""])
            //println "An error occured while getting the report data with REPORT_ID = ${reportID}, QUESTION_ID ${questionID}: ${response.status}"
        }
        reportData
    }

    /**
     * Get all the surveys for the given type (chair, admin, diagnostic, short).
     *
     * @param institutionID The ID of the institution to get the data for.
     * @param types An array of survey types to get data for.
     * @return A list of surveys of the given type; might be empty but never null.
     */
    static def getAllSurveys(institutionID, types, startDate, endDate) {
        //println("Getting surveys for insittutionID: ${institutionID}")
        def surveys = []
        def client = getRESTClient()
        def resultsSeen = 0
        def totalResults = Integer.MAX_VALUE
        def currentResults = 0
        def page = 0
        while((totalResults > resultsSeen + currentResults) && (resultsSeen < MAX_SURVEYS)) {
            def response = client.get( //TODO change the path for start /enddate
                path: "${basePath}/surveys",
                query: [ max: PAGE_SIZE, page: page, institution_id: institutionID, start_date: startDate, end_date: endDate],
                requestContentType: ContentType.JSON,
                headers: authHeaders)
            if(response.status == 200) {
                if(verboseOutput) {
                    println "Surveys data: ${response.data}"
                }

                response.data.data.each { survey ->
                    surveys << survey
                }

                totalResults = response.data.total_results
                currentResults = response.data.data.size()
                resultsSeen += currentResults
                page++
            } else {
                println "An error occured while getting the surveys: ${response.status}"
                break
            }
        }
        surveys
    }

    /**
     * Get all sorted institutions
     * @return
     */
    static def getAllInstitutions(){
        def institutions = []
        def client = getRESTClient()
        def response = client.get(
                path: "${basePath}/institutions",
                query: [ max: 0],
                requestContentType: ContentType.JSON,
                headers: authHeaders)
        if(response.status == 200) {
            if(verboseOutput) {
                println "Institution Data: ${response.data}"
            }
            response.data.data.each { institution ->
                if (getAllSurveys(institution.id,[9,10],"","")){ //TODO This only adds institutions that have surveys: testing purpose only
                    institutions << institution
                }

            }
        } else {
            println "An error occured while getting the institution data."
        }
        institutions.sort{it.name}
    }

    /**
     * Get an instance of the RESTClient that can be used to access the REST API.
     *
     * @return RESTClient An instance that can be used to access the REST API.
     */
    private static RESTClient getRESTClient() {
        if(restClient == null) {
            if(verboseOutput) println "REST requests will be sent to ${hostname} on port ${port} with protocol ${protocol}"

            restClient = new RESTClient("${protocol}://${hostname}:${port}/")
            restClient.ignoreSSLIssues()
            restClient.handler.failure = { response ->
                if(verboseOutput) {
                    println "The REST call failed with status ${response.status}"
                }
                return response
            }
        }
        restClient
    }

    /**
     * Get yyyy/mm/dd as a date format
     * @param date Formatted date
     */
    private static getFormattedDate(date){
        DateGroovyMethods.format(date, 'yyyy-MM-dd')
    }

    /**
     * Writes cell values in Excel from list
     * @param sheet
     * @param cellDataList
     * @param rowIndex
     * @return
     */
    private static writeDataToExcel(sheet, cellDataList, rowIndex){
        Row row = sheet.createRow(rowIndex)
        cellDataList.eachWithIndex { cellValue, index ->
            //println "row=${rowIndex}, cell=${index}, value=${cellValue}"
            Cell cell = row.createCell(index)
            cell.setCellValue(cellValue)
        }
    }
}