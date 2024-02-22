import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as GraphApi from '../../Shared/graph';

let records;

type RequestDataType = {
    meetingId: string
}

export async function httpTriggerGetMeetingRecords(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Request received");
    const userId = process.env.USER_ID;
    context.log('UserId', userId);
    
    const requestData = await request.json() as RequestDataType;
    console.log(requestData);

    records = await GraphApi.getListRecords(userId, requestData.meetingId);
    console.log(records);
    
    return {
        status: 200, /* Defaults to 200 */
        body: JSON.stringify(records.value)
    }    
};

app.http('httpTriggerGetMeetingRecords', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: httpTriggerGetMeetingRecords
});

export default httpTriggerGetMeetingRecords;