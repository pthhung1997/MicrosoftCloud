import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as GraphApi from '../../Shared/graph';

let attendanceReport;

type RequestDataType = {
    meetingId: string,
    reportId: string
}

export async function httpTriggerGetMeetingReportDetail(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Request received");
    const userId = process.env.USER_ID;
    context.log('UserId', userId);
    
    const requestData = await request.json() as RequestDataType;
    console.log(requestData);

    attendanceReport = await GraphApi.getListAttendeeReportDetail(userId, requestData.meetingId, requestData.reportId);
    console.log(JSON.stringify(attendanceReport));
    // const body = JSON.stringify(teamsMeetingLink);
    // const meeting = JSON.parse(body);
    // context.log("meeting:", meeting);
    
    return {
        status: 200, /* Defaults to 200 */
        body: JSON.stringify(attendanceReport.attendanceRecords)
    }    
};

app.http('httpTriggerGetMeetingReportDetail', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: httpTriggerGetMeetingReportDetail
});

export default httpTriggerGetMeetingReportDetail;