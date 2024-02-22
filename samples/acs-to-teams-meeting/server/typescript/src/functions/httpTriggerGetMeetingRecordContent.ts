import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as GraphApi from '../../Shared/graph';

let record;

type RequestDataType = {
    meetingId: string,
    recordId: string
}

export async function httpTriggerGetMeetingRecordContent(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Request received");
    const userId = process.env.USER_ID;
    context.log('UserId', userId);
    
    const requestData = await request.json() as RequestDataType;
    console.log(requestData);

    record = await GraphApi.getRecord(userId, requestData.meetingId, requestData.recordId);
    console.log(record);
    
    return {
        status: 200, /* Defaults to 200 */
    }    
};

app.http('httpTriggerGetMeetingRecordContent', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: httpTriggerGetMeetingRecordContent
});

export default httpTriggerGetMeetingRecordContent;