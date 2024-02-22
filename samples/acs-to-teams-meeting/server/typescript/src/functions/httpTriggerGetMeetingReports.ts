import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as GraphApi from '../../Shared/graph';

let attendanceReports;

type RequestDataType = {
    meetingId: string
}

export async function httpTriggerGetMeetingReports(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Request received");
    const userId = process.env.USER_ID;
    context.log('UserId', userId);
    
    const requestData = await request.json() as RequestDataType;
    console.log(requestData);

    attendanceReports = await GraphApi.getListAttendeeReports(userId, requestData.meetingId);
    console.log(attendanceReports);
    // const body = JSON.stringify(teamsMeetingLink);
    // const meeting = JSON.parse(body);
    // context.log("meeting:", meeting);
    
    return {
        status: 200, /* Defaults to 200 */
        body: JSON.stringify(attendanceReports.value)
    }    
};

app.http('httpTriggerGetMeetingReports', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: httpTriggerGetMeetingReports
});

export default httpTriggerGetMeetingReports;