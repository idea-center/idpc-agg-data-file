package org.ideaedu

import groovyx.net.http.RESTClient
import groovyx.net.http.ContentType

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

    private static final def DEFAULT_HOSTNAME = "rest.ideasystem.org"
    private static final def DEFAULT_PORT = 443
    private static final def DEFAULT_BASE_PATH = "IDEA-REST-SERVER/v1"
    private static final def DEFAULT_AUTH_HEADERS = [ "X-IDEA-APPNAME": "", "X-IDEA-KEY": "" ]
    private static final def DEFAULT_PROTOCOL = "https"
    private static final def DEFAULT_INSTITUTION_ID = 1029
    private static final def DEFAULT_START_DATE = new Date() - 10
    private static final def DEFAULT_END_DATE = new Date() - 1
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
        "Sumnmer" : 4,
        "Other" : 5
    ]
    private static final def DIAGNOSTIC_QUESTION_ID_LIST = [490..492,471..482].flatten()
    private static final def DIAGNOSTIC_QUESTION_TEACHING_METHOD_ID_LIST = [451..470].flatten()
    private static final def DIAGNOSTIC_ONLY_QUESTION_ID_LIST = [483..489,493].flatten()
    private static final def DIAGNOSTIC_ADDITIONAL_QUESTION_ID_LIST = [546..565].flatten()
    private static final def SHORT_QUESTION_ID_LIST = [716..718,701..712].flatten()
    private static final def RESEARCH_QUESTION_ID_LIST = [695,494..497].flatten()
    private static final def SHORT_ADDITIONAL_QUESTION_ID_LIST = [546..565].flatten()

    /** The maximum number of surveys to get before quitting. */
    private static final def MAX_SURVEYS = 1

    /** The number of surveys to get per page */
    private static final def PAGE_SIZE = 1

    private static def institutionID = DEFAULT_INSTITUTION_ID
    private static def startDate = DEFAULT_START_DATE
    private static def endDate = DEFAULT_END_DATE
    private static def hostname = DEFAULT_HOSTNAME
    private static def protocol = DEFAULT_PROTOCOL
    private static def port = DEFAULT_PORT
    private static def basePath = DEFAULT_BASE_PATH
    private static def authHeaders = DEFAULT_AUTH_HEADERS

    private static def verboseOutput = false

    private static RESTClient restClient

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
            institutionID = options.i.toInteger()
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



        /*
         * The following will get all the surveys that are available of the
         * given type and print out the overall ratings for each survey subject.
         * This will print the raw and adjusted mean and t-score for each survey
         * subject.
         */
        def types = [ 9, 10 ]
        def institution = getInstitution(institutionID)
        def surveys = getAllSurveys(institutionID, types)
        if(surveys) {
            // Print the CSV header
            println "ID, FICE, Institution, Term_Year, Instructor, Dept_Code_Name, Course_Num, Dept_Name, Dept_Code, Local_Code, Time, Days, Enrolled, Responses, Form, Year_Term_Code, Batch, " +
                    "Obj1, Obj2, Obj3, Obj4, Obj5, Obj6, Obj7, Obj8, Obj9, Obj10, Obj11, Obj12, " +
                    "PRO_Raw_Mean,PRO_Adj_Mean,PRO_CRaw_IDEA,PRO_CAdj_IDEA,PRO_CRaw_Disc,PRO_CAdj_Disc,PRO_CRaw_Inst,PRO_CAdj_Inst,Impr_Stu_Att_Raw_Mean,Impr_Stu_Att_Adj_Mean,Impr_Stu_Att_CRaw_IDEA,Impr_Stu_Att_CAdj_IDEA,Exc_Tchr_Raw_Mean,Exc_Tchr_Adj_Mean,Exc_Tchr_CRaw_IDEA,Exc_Tchr_CAdj_IDEA,Exc_Tchr_CRaw_Disc,Exc_Tchr_CAdj_Disc,Exc_Tchr_CRaw_Inst,Exc_Tchr_CAdj_Inst,Exc_Crs_Raw_Mean,Exc_Crs_Adj_Mean,Exc_Crs_CRaw_IDEA,Exc_Crs_CAdj_IDEA,Exc_Crs_CRaw_Disc,Exc_Crs_CAdj_Disc,Exc_Crs_CRaw_Inst,Exc_Crs_CAdj_Inst,SumEval_Raw_Mean,SumEval_Adj_Mean,SumEval_CRaw_IDEA,SumEval_CAdj_IDEA,SumEval_CRaw_Disc,SumEval_CAdj_Disc,SumEval_CRaw_Inst,SumEval_CAdj_Inst,Obj1,Obj1_Raw_Mean,Obj1_Adj_Mean,Obj1_CRaw_IDEA,Obj1_CAdj_IDEA,Obj1_CRaw_Disc,Obj1_CAdj_Disc,Obj1_CRaw_Inst,Obj1_CAdj_Inst,Obj2,Obj2_Raw_Mean,Obj2_Adj_Mean,Obj2_CRaw_IDEA,Obj2_CAdj_IDEA,Obj2_CRaw_Disc,Obj2_CAdj_Disc,Obj2_CRaw_Inst,Obj2_CAdj_Inst,Obj3,Obj3_Raw_Mean,Obj3_Adj_Mean,Obj3_CRaw_IDEA,Obj3_CAdj_IDEA,Obj3_CRaw_Disc,Obj3_CAdj_Disc,Obj3_CRaw_Inst,Obj3_CAdj_Inst,Obj4,Obj4_Raw_Mean,Obj4_Adj_Mean,Obj4_CRaw_IDEA,Obj4_CAdj_IDEA,Obj4_CRaw_Disc,Obj4_CAdj_Disc,Obj4_CRaw_Inst,Obj4_CAdj_Inst,Obj5,Obj5_Raw_Mean,Obj5_Adj_Mean,Obj5_CRaw_IDEA,Obj5_CAdj_IDEA,Obj5_CRaw_Disc,Obj5_CAdj_Disc,Obj5_CRaw_Inst,Obj5_CAdj_Inst,Obj6,Obj6_Raw_Mean,Obj6_Adj_Mean,Obj6_CRaw_IDEA,Obj6_CAdj_IDEA,Obj6_CRaw_Disc,Obj6_CAdj_Disc,Obj6_CRaw_Inst,Obj6_CAdj_Inst,Obj7,Obj7_Raw_Mean,Obj7_Adj_Mean,Obj7_CRaw_IDEA,Obj7_CAdj_IDEA,Obj7_CRaw_Disc,Obj7_CAdj_Disc,Obj7_CRaw_Inst,Obj7_CAdj_Inst,Obj8,Obj8_Raw_Mean,Obj8_Adj_Mean,Obj8_CRaw_IDEA,Obj8_CAdj_IDEA,Obj8_CRaw_Disc,Obj8_CAdj_Disc,Obj8_CRaw_Inst,Obj8_CAdj_Inst,Obj9,Obj9_Raw_Mean,Obj9_Adj_Mean,Obj9_CRaw_IDEA,Obj9_CAdj_IDEA,Obj9_CRaw_Disc,Obj9_CAdj_Disc,Obj9_CRaw_Inst,Obj9_CAdj_Inst,Obj10,Obj10_Raw_Mean,Obj10_Adj_Mean,Obj10_CRaw_IDEA,Obj10_CAdj_IDEA,Obj10_CRaw_Disc,Obj10_CAdj_Disc,Obj10_CRaw_Inst,Obj10_CAdj_Inst,Obj11,Obj11_Raw_Mean,Obj11_Adj_Mean,Obj11_CRaw_IDEA,Obj11_CAdj_IDEA,Obj11_CRaw_Disc,Obj11_CAdj_Disc,Obj11_CRaw_Inst,Obj11_CAdj_Inst,Obj12,Obj12_Raw_Mean,Obj12_Adj_Mean,Obj12_CRaw_IDEA,Obj12_CAdj_IDEA,Obj12_CRaw_Disc,Obj12_CAdj_Disc,Obj12_CRaw_Inst,Obj12_CAdj_Inst,Method1_Mean,Method2_Mean,Method3_Mean,Method4_Mean,Method5_Mean,Method6_Mean,Method7_Mean,Method8_Mean,Method9_Mean,Method10_Mean,Method11_Mean,Method12_Mean,Method13_Mean,Method14_Mean,Method15_Mean,Method16_Mean,Method17_Mean,Method18_Mean,Method19_Mean,Method20_Mean,Read_Mean,Read_Cnv_IDEA,Read_Cnv_Disc,Read_Cnv_Inst,NonRead_Mean,NonRead_Cnv_IDEA,NonRead_Cnv_Disc,NonRead_Cnv_Inst,Diff_Mean,Diff_Cnv_IDEA,Diff_Cnv_Disc,Diff_Cnv_Inst,Q36_Mean,Effort_Mean,Effort_Cnv_IDEA,Effort_Cnv_Disc,Effort_Cnv_Inst,Q38_Mean,Motivation_Mean,Motivation_Cnv_IDEA,Motivation_Cnv_Disc,Motivation_Cnv_Inst,WkHabit_Mean,WkHabit_Cnv_IDEA,WkHabit_Cnv_Disc,WkHabit_Cnv_Inst,Background_Mean,Q44_Mean,Q45_Mean,Q46_Mean,Q47_Mean,Add_Q1_Mean,Add_Q2_Mean,Add_Q3_Mean,Add_Q4_Mean,Add_Q5_Mean,Add_Q6_Mean,Add_Q7_Mean,Add_Q8_Mean,Add_Q9_Mean,Add_Q10_Mean,Add_Q11_Mean,Add_Q12_Mean,Add_Q13_Mean,Add_Q14_Mean,Add_Q15_Mean,Add_Q16_Mean,Add_Q17_Mean,Add_Q18_Mean,Add_Q19_Mean,Add_Q20_Mean"

            surveys.each { survey ->
                def surveySubject = survey.info_form.respondents[0]
                def formName = getFormName(survey.rater_form.id)
                def discipline = [ code:"", name: "", abbrev: "" ] // TODO getDiscipline(survey.info_form.discipline_code)
                def reports = getReports(survey.id, "Short")
                def questionList = SHORT_QUESTION_ID_LIST
                def additionalQuestionList = SHORT_ADDITIONAL_QUESTION_ID_LIST
                if(!reports) {
                    reports = getReports(survey.id, "Diagnostic")
                    questionList = DIAGNOSTIC_QUESTION_ID_LIST
                    additionalQuestionList = DIAGNOSTIC_ADDITIONAL_QUESTION_ID_LIST
                }
                def reportID = reports[0]?.id
                if(reportID) {
                    def model = getReportModel(reportID)

                    print "${survey.id},"
                    print "${institution.fice},"
                    print "${institution.name},"
                    print "${survey.term}," //e.g. Fall 2013
                    print "${surveySubject.first_name} ${surveySubject.last_name},"
                    print "${discipline?.code},"
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

                    def term = survey.term.split(" ")[0]
                    TERM_CODE_MAP.each { t ->
                        if (t.key == term) print "${survey.year}${t.value}," //Year_Term_Code
                    }
                    print "," // batch is unused in IDEA-CL

                    //Objectives
                    OBJECTIVE_MAP.each { objective ->
                        def response = "Default-Imp"

                        model.aggregate_data.relevant_results.questions.each { question ->
                            if(question.question_id == objective.key) {
                                response = question.response
                            }
                        }
                        print "${response},"
                    }
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

                    /* As a result of taking this course, I have more positive feelings toward this field of study --- Diagnostic - 40 and Short - 16 */
                    /* Excellent teacher: Overall, I rate this instructor an excellent teacher --- Diagnostic - 41 and Short - 17 */
                    /* Excellent course: Overall, I rate this course as excellent --- Diagnostic - 42 and Short - 18 */


                    questionList.eachWithIndex{ questionID, index ->
                        def reportData = getReportDataByQuestion(reportID, questionID)
                        print "${reportData.results.result.raw.mean},"          //raw_mean
                        print "${reportData.results.result.adjusted.mean},"     //adj_mean
                        print "${reportData.results.result.raw.tscore},"        //raw t-score
                        print "${reportData.results.result.adjusted.tscore},"   //raw adjusted t-score

                        //Except Diagnostic 40 / Short 16 (index = 0), we need compared to discipline / institution
                        if (index != 0){
                            print "${reportData.results.discipline_result.raw.tscore},"         //discipline raw t-score
                            print "${reportData.results.discipline_result.adjusted.tscore},"    //discipline adjusted t-score
                            print "${reportData.results.institution_result.raw.tscore},"        //institution raw t-score
                            print "${reportData.results.institution_result.adjusted.tscore},"   //institution adjusted t-score
                        }

                        //Add Summary before Objective 1 (Diagnostic 21 / Short 1)
                        if (index == 3){
                            print "${model.aggregate_data.summary_evaluation.result.raw.mean},"                     //SumEval_Raw_Mean
                            print "${model.aggregate_data.summary_evaluation.result.adjusted.mean},"                //SumEval_Adj_Mean
                            print "${model.aggregate_data.summary_evaluation.result.raw.tscore},"                   //SumEval_CRaw_IDEA
                            print "${model.aggregate_data.summary_evaluation.result.adjusted.tscore},"              //SumEval_CAdj_IDEA
                            print "${model.aggregate_data.summary_evaluation.discipline_result.raw.tscore},"        //SumEval_CRaw_Disc
                            print "${model.aggregate_data.summary_evaluation.discipline_result.adjusted.tscore},"   //SumEval_CAdj_Disc
                            print "${model.aggregate_data.summary_evaluation.institution_result.raw.tscore},"       //SumEval_CRaw_Inst
                            print "${model.aggregate_data.summary_evaluation.institution_result.adjusted.tscore},"  //SumEval_CAdj_Inst
                        }
                    }

                    //Teaching Method 1-20 (Diagnostic 1-20)
                    DIAGNOSTIC_QUESTION_TEACHING_METHOD_ID_LIST.each { questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID)
                        print (reportData==null? "":"${reportData.results.result.raw.mean},")
                    }

                    //Disgnostic 33-40
                    DIAGNOSTIC_ONLY_QUESTION_ID_LIST.each{ questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID)
                        print (reportData==null? "":"${reportData.results.result.raw.mean},")          //raw_mean

                        //Diagnostic 36, 38 (QUESTION_ID = 486, 488) only requires mean
                        if (questionID != 486 || questionID != 488){
                            print (reportData==null? "":"${reportData.results.result.raw.tscore},")        //raw t-score
                            print (reportData==null? "":"${reportData.results.discipline_result.raw.tscore}," )        //discipline raw t-score
                            print (reportData==null? "":"${reportData.results.institution_result.raw.tscore}," )       //institution raw t-score
                        }
                    }

                    /* Research Questions */
                    RESEARCH_QUESTION_ID_LIST.each{ questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID)
                        print (reportData==null? "":"${reportData.results.result.raw.mean}," )         //raw_mean
                    }

                    //Additional Questions
                    additionalQuestionList.each{ questionID ->
                        def reportData = getReportDataByQuestion(reportID, questionID)
                        print (reportData==null? "":"${reportData.results.result.raw?.mean},")          //raw_mean
                    }
                    println ""
                }
            }
        } else {
            println "No surveys are available."
        }
    }

    static def getDiscipline(disciplineID) {
        def discipline

        if(disciplineID) {
            def client = getRESTClient()
            def response = client.get(
                path: "${basePath}/discipline_code/${disciplineID}",
                requestContentType: ContentType.JSON,
                headers: authHeaders)
            if(response.status == 200) {
                if(verboseOutput) {
                    println "Discipline data: ${response.data}"
                }

                if(response.data) {
                    discipline = response.data
                }
            } else {
                println "An error occured while getting the discipline with ID ${disciplineID}: ${response.status}"
            }
        }

        return discipline
    }

    /**
     * Get the institution information that has the given ID.
     *
     * @param institutionID The ID of the institution.
     * @return The institution with the given ID.
     */
    static def getInstitution(institutionID) {
        def institution

        def client = getRESTClient()
        def response = client.get(
            path: "${basePath}/institutions",
            query: [ id: institutionID ],
            requestContentType: ContentType.JSON,
            headers: authHeaders)
        if(response.status == 200) {
            if(verboseOutput) {
                println "Institution data: ${response.data}"
            }

            if(response.data && response.data.data && response.data.data.size() > 0) {
                // take the first one ... not sure why we would end up with more than 1
                institution = response.data.data.get(0)
            }
        } else {
            println "An error occured while getting the institution with ID ${institutionID}: ${response.status}"
        }

        return institution
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

        return reports
    }

    static def getFormName(id) {
        return FORM_NAMES[id]
    }

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
            println "An error occured while getting the report model with ID ${reportID}: ${response.status}"
        }

        return reportModel
    }

    static def getReportDataByQuestion(reportID, questionID) {
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
        } else {
            //println "An error occured while getting the report data with REPORT_ID = ${reportID}, QUESTION_ID ${questionID}: ${response.status}"
        }

        return reportData
    }

    /**
     * Get all the surveys for the given type (chair, admin, diagnostic, short).
     *
     * @param institutionID The ID of the institution to get the data for.
     * @param types An array of survey types to get data for.
     * @return A list of surveys of the given type; might be empty but never null.
     */
    static def getAllSurveys(institutionID, types) {
        def surveys = []

        def client = getRESTClient()
        def resultsSeen = 0
        def totalResults = Integer.MAX_VALUE
        def currentResults = 0
        def page = 0
        while((totalResults > resultsSeen + currentResults) && (resultsSeen < MAX_SURVEYS)) {
            def response = client.get(
                path: "${basePath}/surveys",
                query: [ max: PAGE_SIZE, page: page/*, institution_id: institutionID, types: types */],
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

        return surveys
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

        return restClient
    }
}