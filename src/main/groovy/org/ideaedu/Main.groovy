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

    /** The maximum number of surveys to get before quitting. */
    private static final def MAX_SURVEYS = 5

    /** The number of surveys to get per page */
    private static final def PAGE_SIZE = 5

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
            println "ID, FICE, Institution, Term_Year, Instructor, Dept_Code_Name, Course_Num, Dept_Name, Dept_Code, Local_Code, Time, Days, Enrolled, Responses, Form, Year_Term"

            surveys.each { survey ->
                def surveySubject = survey.info_form.respondents[0]
                def formName = getFormName(survey.rater_form.id)
                def discipline = [ code:"", name: "", abbrev: "" ] // TODO getDiscipline(survey.info_form.discipline_code)
                def reports = getReports(survey.id, "Short")
                if(!reports) {
                    reports = getReports(survey.id, "Diagnostic")
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

                    OBJECTIVE_MAP.each { objective ->
                        def response = "Default-Imp"

                        model.aggregate_data.relevant_results.questions.each { question ->
                            if(question.question_id == objective.key) {
                                response = question.response
                            }
                        }

                        print "${response},"
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