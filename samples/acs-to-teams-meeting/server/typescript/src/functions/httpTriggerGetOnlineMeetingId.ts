import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as GraphApi from '../../Shared/graph';

let attendanceReports;

type RequestDataType = {
    joinUrl: string
}

export async function httpTriggerGetOnlineMeetingId(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Request received");
    const userId = process.env.USER_ID;
    context.log('UserId', userId);
    
    const requestData = await request.json() as RequestDataType;
    console.log(requestData);

    attendanceReports = await GraphApi.getOnlineMeetingId(userId, requestData.joinUrl);
    console.log(attendanceReports);
    // const body = JSON.stringify(teamsMeetingLink);
    // const meeting = JSON.parse(body);
    // context.log("meeting:", meeting);
    
    return {
        status: 200, /* Defaults to 200 */
        body: attendanceReports.value[0].id
    }    
};

app.http('httpTriggerGetOnlineMeetingId', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: httpTriggerGetOnlineMeetingId
});

export default httpTriggerGetOnlineMeetingId;