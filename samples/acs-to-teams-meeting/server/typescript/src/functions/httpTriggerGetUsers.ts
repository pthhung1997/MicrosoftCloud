import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import * as GraphApi from '../../Shared/graph';

let users;

export async function httpTriggerGetUsers(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Request received");
    
    users = await GraphApi.getUsers();
    console.log(users);
    const body = JSON.stringify(users.value);
    
    return {
        status: 200, /* Defaults to 200 */
        body: body
    }    
};

app.http('httpTriggerGetUsers', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: httpTriggerGetUsers
});

export default httpTriggerGetUsers;