import { startDateTimeAsync, endDateTimeAsync } from './dateTimeFormat';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import 'isomorphic-fetch';

let clientSecretCredential;
let appGraphClient;

function ensureGraphForAppOnlyAuth() {

  if (!clientSecretCredential) {
    clientSecretCredential = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );
  }

  if (!appGraphClient) {
    const authProvider = new TokenCredentialAuthenticationProvider(
      clientSecretCredential, {
        scopes: [ 'https://graph.microsoft.com/.default' ]
      });

    appGraphClient = Client.initWithMiddleware({
      authProvider: authProvider
    });
  }
}

async function createNewMeetingAsync(userId) {
    ensureGraphForAppOnlyAuth();
    let startTime = await startDateTimeAsync();
    let endTime = await endDateTimeAsync();
    const newMeeting = `/users/${userId}/calendar/events`;
    
    const event = {
      subject: 'Customer Service Meeting',
      start: {
          dateTime: startTime,
          timeZone: 'UTC'
      },
      end: {
          dateTime: endTime,
          timeZone: 'UTC'
      },
      isOnlineMeeting: true
    };
    
    const newEvent = await appGraphClient.api(newMeeting).post(event);    
    return newEvent;     
}

async function getUsers() {
    ensureGraphForAppOnlyAuth();
    const users = await appGraphClient.api(`/users`).get();

    return users;

}

async function getOnlineMeetingId(userId, joinUrl) {
    ensureGraphForAppOnlyAuth();
    const onlineMeetings = await appGraphClient.api(`/users/${userId}/onlineMeetings`).filter(`JoinWebUrl eq '${joinUrl}'`).get();
    return onlineMeetings;

}

async function getListAttendeeReports(userId, meetingId) {
  ensureGraphForAppOnlyAuth();
  console.log(`/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`)
  const attendanceReports = await appGraphClient.api(`/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`).get();
  return attendanceReports;

}



async function getListAttendeeReportDetail(userId, meetingId, reportId) {
  ensureGraphForAppOnlyAuth();
  console.log(`/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${reportId}`)
  const attendanceReport = await appGraphClient.api(`/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${reportId}`).expand('attendanceRecords').get();
  return attendanceReport;

}

async function getListRecords(userId, meetingId) {
  ensureGraphForAppOnlyAuth();
  console.log(`/users/${userId}/onlineMeetings/${meetingId}/recordings`)
  const records = await appGraphClient.api(`/users/${userId}/onlineMeetings/${meetingId}/recordings`).get();
  return records;

}

async function getRecord(userId, meetingId, recordId) {
  ensureGraphForAppOnlyAuth();
  console.log(`/users/${userId}/onlineMeetings/${meetingId}/recordings/${recordId}/content`)
  const record = await appGraphClient.api(`/users/${userId}/onlineMeetings/${meetingId}/recordings/${recordId}/content`).get();
  return record;

}
      
export {createNewMeetingAsync, getOnlineMeetingId, getListAttendeeReports, getListAttendeeReportDetail, getListRecords, getRecord, getUsers}