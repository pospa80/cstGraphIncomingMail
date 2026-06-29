package com.tririga.custom;

import com.sun.istack.ByteArrayDataSource;
import com.tririga.pub.workflow.CustomBusinessConnectTask;
import com.tririga.pub.workflow.CustomParamTaskResultImpl;
import com.tririga.pub.workflow.Record;
import com.tririga.ws.TririgaWS;
import com.tririga.ws.dto.*;
import com.tririga.ws.dto.content.Content;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.core.config.Configurator;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Date;
import java.util.List;

import com.google.gson.Gson;

import org.apache.http.client.methods.HttpPatch;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.client.methods.CloseableHttpResponse;

import org.apache.http.conn.ssl.SSLConnectionSocketFactory;

import javax.activation.DataHandler;
import javax.net.ssl.SSLContext;

public class cstIncomingMail implements CustomBusinessConnectTask {


    static {
        Configurator.initialize(null, String.valueOf(cstIncomingMail.class.getClassLoader().getResource("custom-log4j2.xml")));
    }

    public static int BUFFER_SIZE = 102400;

    private static final org.apache.logging.log4j.Logger log = LogManager.getLogger(cstIncomingMail.class);

    @Override
    public boolean execute(TririgaWS tws, long l, Record[] records) {
        CustomParamTaskResultImpl result = new CustomParamTaskResultImpl();
        String output = "";
        //Gathering of info from Tririga
        com.tririga.ws.dto.Record[] recordHeaders = null;
        try {
            tws.register(l);
            int count = records.length;
            long[] recordIds = new long[count];
            for (int i = 0; i < count; i++) {
                recordIds[i] = records[i].getId();
                log.info("Record " + i + ": " + records[i].getId());
            }
            recordHeaders = tws.getRecordDataHeaders(recordIds);
        } catch (Exception ex) {
            output = ex.getMessage();
            log.error(ex.getMessage());
            return false;
        }

        int recordCount = 0;
        boolean succeeded = true;
        StringBuilder logTimes = new StringBuilder();

        try {
            String accessToken = fetchAccessToken();
            String jsonResponse = fetchUnreadEmails(accessToken);


            List<GraphMessage> emails = parseEmails(jsonResponse);
            populateAttachments(emails, accessToken);
            createEmailMessageRecords(emails, tws, accessToken);

        } catch (Exception e) {
            log.error(e.getMessage());
            throw new RuntimeException(e);
        }


        return false;
    }


    private String fetchAccessToken() throws Exception {
        String tokenUrl = "https://login.microsoftonline.com/" + TENANT_ID + "/oauth2/v2.0/token";

        String body = "client_id=" + CLIENT_ID
                + "&client_secret=" + URLEncoder.encode(CLIENT_SECRET, "UTF-8")
                + "&scope=" + URLEncoder.encode("https://graph.microsoft.com/.default", "UTF-8")
                + "&grant_type=client_credentials";

        URL url = new URL(tokenUrl);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("POST");
        conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
        conn.setRequestProperty("Accept", "application/x-www-form-urlencoded");
        conn.setDoOutput(true);

        // Write body
        try (OutputStream os = conn.getOutputStream()) {
            os.write(body.getBytes("UTF-8"));
        }

        String response = readResponse(conn);
        conn.disconnect();

        return parseJsonValue(response, "access_token");
    }

    private String fetchUnreadEmails(String accessToken) throws Exception {
        String mailUrl = "https://graph.microsoft.com/v1.0/users/" + MAILBOX
                + "/messages"
                + "?$filter=isRead%20eq%20false"
                + "&$select=id,subject,body,from,receivedDateTime,sentDateTime,hasAttachments,toRecipients";

        URL url = new URL(mailUrl);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setRequestProperty("Authorization", "Bearer " + accessToken);
        conn.setRequestProperty("Content-Type", "application/json");
        conn.setRequestProperty("Accept", "application/json");

        String response = readResponse(conn);
        conn.disconnect();
        return response;
    }

    private String fetchAttachments(String accessToken, String messageId) throws Exception {
        log.info("fetching attachments for {}: ", messageId);
        String attachmentsUrl = "https://graph.microsoft.com/v1.0/users/" + MAILBOX
                + "/messages/" + messageId
                + "/attachments";

        URL url = new URL(attachmentsUrl);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setRequestProperty("Authorization", "Bearer " + accessToken);
        conn.setRequestProperty("Content-Type", "application/json");
        conn.setRequestProperty("Accept", "application/json");

        String response = readResponse(conn);
        conn.disconnect();
        return response;
    }

    public List<GraphEmailAttachment> parseAttachments(String jsonResponse) {
        log.info("fff");
        Gson gson = new Gson();
        GraphAttachmentResponse response = gson.fromJson(jsonResponse, GraphAttachmentResponse.class);
        log.info("ggg");
        return response.value;
    }

    /**
     * For every message that has attachments, fetch them from Graph and
     * populate msg.attachments in place.
     */
    private void populateAttachments(List<GraphMessage> messages, String accessToken) {
        log.info("populateAttachments");
        for (GraphMessage msg : messages) {
            log.info("ccc {}", msg.id);
            if (!msg.hasAttachments) {
                log.info("dddd");
                continue;
            }
            try {
                String attachmentJson = fetchAttachments(accessToken, msg.id);
                log.info("eeee");
                List<GraphEmailAttachment> attachments = parseAttachments(attachmentJson);
                msg.attachments = attachments;
                log.info("Fetched " + (attachments != null ? attachments.size() : 0)
                        + " attachment(s) for message " + msg.id);
            } catch (Exception e) {
                log.error("Failed to fetch attachments for message " + msg.id, e);
            }
        }
        log.info("end populateAttachments");
    }

    private void markAsRead(String accessToken, String messageId) throws Exception {
        String patchUrl = "https://graph.microsoft.com/v1.0/users/" + MAILBOX
                + "/messages/" + messageId;


        try {
            SSLContext sslContext = SSLContext.getInstance("TLSv1.2");
            sslContext.init(null, null, null);

            SSLConnectionSocketFactory sslSocketFactory = new SSLConnectionSocketFactory(
                    sslContext,
                    new String[]{"TLSv1.2"},
                    null,
                    SSLConnectionSocketFactory.getDefaultHostnameVerifier()
            );

            CloseableHttpClient client = HttpClients.custom()
                    .setSSLSocketFactory(sslSocketFactory)
                    .build();

            HttpPatch patch = new HttpPatch(patchUrl);
            patch.setHeader("Authorization", "Bearer " + accessToken);
            patch.setHeader("Content-Type", "application/json");
            patch.setHeader("Accept", "application/json");
            patch.setEntity(new StringEntity("{\"isRead\": true}", "UTF-8"));

            try (CloseableHttpResponse response = client.execute(patch)) {
                int statusCode = response.getStatusLine().getStatusCode();
                if (statusCode != 200) {
                    throw new Exception("Failed to mark email as read. Status: " + statusCode);
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage());
            throw new RuntimeException(e);
        }
    }

    private String readResponse(HttpURLConnection conn) throws Exception {
        int responseCode = conn.getResponseCode();
        InputStream is = (responseCode >= 200 && responseCode < 300)
                ? conn.getInputStream()
                : conn.getErrorStream();

        BufferedReader reader = new BufferedReader(
                new InputStreamReader(is, "UTF-8"));
        StringBuilder sb = new StringBuilder();
        String line;
        while ((line = reader.readLine()) != null) {
            sb.append(line);
        }
        reader.close();
        return sb.toString();
    }

    private String parseJsonValue(String json, String key) {
        String searchKey = "\"" + key + "\":\"";
        int start = json.indexOf(searchKey);
        if (start == -1) return null;
        start += searchKey.length();
        int end = json.indexOf("\"", start);
        return json.substring(start, end);
    }

    private void allowPatch() {
        try {
            java.lang.reflect.Field methodsField = HttpURLConnection.class.getDeclaredField("methods");
            methodsField.setAccessible(true);

            java.lang.reflect.Field modifiersField = java.lang.reflect.Field.class.getDeclaredField("modifiers");
            modifiersField.setAccessible(true);
            modifiersField.setInt(methodsField, methodsField.getModifiers() & ~java.lang.reflect.Modifier.FINAL);

            String[] methods = {
                    "GET", "POST", "HEAD", "OPTIONS", "PUT", "DELETE", "TRACE", "PATCH"
            };
            methodsField.set(null, methods);
        } catch (Exception e) {
            throw new RuntimeException("Failed to enable PATCH method", e);
        }
    }

    public List<GraphMessage> parseEmails(String jsonResponse) {
        Gson gson = new Gson();
        GraphMessageResponse response = gson.fromJson(jsonResponse, GraphMessageResponse.class);
        return response.value;
    }

   private void createEmailMessageRecords(List<GraphMessage> messages, TririgaWS tws, String accessToken) {
        for (GraphMessage msg : messages) {
            try {
                IntegrationSection section = new IntegrationSection();
                section.setName("General Info");

                populateFieldsFromGraphMessage(msg, section);



                IntegrationRecord newRecord = new IntegrationRecord();
                newRecord.setActionName("CREATE");
                newRecord.setSections(new IntegrationSection[]{section});
                newRecord.setId(-1);
                newRecord.setGuiId(10014746);
                newRecord.setObjectTypeId(10009438);
                newRecord.setObjectTypeName("EmailMessage");
                newRecord.setModuleId(17);

                log.info("Saving EmailMessage record for subject: " + msg.subject);

                ResponseHelperHeader rhh = tws.saveRecord(new IntegrationRecord[]{newRecord});

                long newRecordId = extractRecordId(rhh); // TODO: confirm how to pull ID off rhh
                //if (newRecordId <= 0) {
                //    log.error("saveRecord did not return a valid record ID for subject: " + msg.getSubject());
                //    continue;
                //}


                if (msg.hasAttachments && msg.attachments != null) {
                    for (GraphEmailAttachment attachment : msg.attachments) {
                        log.info("calling upload Attachment for : {}", attachment.name);
                        uploadAttachment(attachment, newRecordId, tws);

                    }
                }

                for (GraphPerson recipient : msg.toRecipients) {
                    uploadToRecipient(recipient, newRecordId, tws);
                }

                if(msg.from != null) {
                    uploadFrom(msg.from, newRecordId, tws);
                }

                markAsRead(accessToken, msg.id);
            } catch (Exception e) {
                log.error("Failed to create EmailMessage record for subject: {}", msg.subject, e);
            }
        }
    }

    private void uploadToRecipient(GraphPerson recipient, long recordId, TririgaWS tws) {
        try {
            IntegrationSection section = new IntegrationSection();
            section.setName("EmailAddress");

            populateFieldsFromGraphAddress(recipient, section);

            IntegrationRecord newRecord = new IntegrationRecord();
            newRecord.setActionName("CREATE");
            newRecord.setSections(new IntegrationSection[]{section});
            newRecord.setId(-1);
            newRecord.setGuiId(10014744);
            newRecord.setObjectTypeId(10009441);
            newRecord.setObjectTypeName("EmailAddress");
            newRecord.setModuleId(17);

            log.info("Saving EmailAddress record for file: {}", recipient.emailAddress);

            ResponseHelperHeader rhh = tws.saveRecord(new IntegrationRecord[]{newRecord});

            long newRecordId = extractRecordId(rhh); // TODO: confirm how to pull ID off rhh

            Association association = new Association();
            association.setAssociatedRecordId(newRecordId);
            association.setRecordId(recordId);
            association.setAssociationName("Email To Address");

            tws.associateRecords(new Association[] {association});

            log.info("");

        } catch (Exception e) {
            log.error("Failed to create RecipientToAddress record for: {}", recipient.emailAddress, e);
        }
    }

    private void uploadFrom(GraphPerson from, long recordId, TririgaWS tws) {
        try {
            IntegrationSection section = new IntegrationSection();
            section.setName("EmailAddress");

            populateFieldsFromGraphAddress(from, section);

            IntegrationRecord newRecord = new IntegrationRecord();
            newRecord.setActionName("CREATE");
            newRecord.setSections(new IntegrationSection[]{section});
            newRecord.setId(-1);
            newRecord.setGuiId(10014744);
            newRecord.setObjectTypeId(10009441);
            newRecord.setObjectTypeName("EmailAddress");
            newRecord.setModuleId(17);

            log.info("Saving from EmailAddress record for file: {}", from.emailAddress);

            ResponseHelperHeader rhh = tws.saveRecord(new IntegrationRecord[]{newRecord});

            long newRecordId = extractRecordId(rhh); // TODO: confirm how to pull ID off rhh

            Association association = new Association();
            association.setAssociatedRecordId(newRecordId);
            association.setRecordId(recordId);
            association.setAssociationName("Email From Address");

            tws.associateRecords(new Association[] {association});

            log.info("");

        } catch (Exception e) {
            log.error("Failed to create RecipientToAddress record for: {}", from.emailAddress, e);
        }
    }

    private long extractRecordId(ResponseHelperHeader rhh) {
        if (rhh == null) {
            log.error("saveRecord returned a null response header");
            return -1;
        }

        if (rhh.isAnyFailed()) {
            log.error("saveRecord reported failure: failed=" + rhh.getFailed() + " total=" + rhh.getTotal());
            return -1;
        }

        ResponseHelper[] helpers = rhh.getResponseHelpers();
        if (helpers == null || helpers.length == 0) {
            log.error("saveRecord response contained no responseHelpers entries");
            return -1;
        }

        ResponseHelper helper = helpers[0];

        if (!"Successful".equals(helper.getStatus())) {
            log.error("saveRecord helper status not Successful: " + helper.getStatus());
            return -1;
        }

        return helper.getRecordId();
    }

    private void uploadAttachment(GraphEmailAttachment attachment, long recordId, TririgaWS tws) {
        try {
            IntegrationSection section = new IntegrationSection();
            section.setName("EmailAttachment"); // TODO: confirm exact section name

            populateFieldsFromGraphAttachment(attachment, section);

            IntegrationRecord newRecord = new IntegrationRecord();
            newRecord.setActionName("CREATE"); // TODO: confirm
            newRecord.setSections(new IntegrationSection[]{section});
            newRecord.setId(-1);
            newRecord.setGuiId(10014745);
            newRecord.setObjectTypeId(10009440);
            newRecord.setObjectTypeName("EmailAttachment");
            newRecord.setModuleId(17);

            log.info("Saving EmailAttachment record for file: " + attachment.name);

            ResponseHelperHeader rhh = tws.saveRecord(new IntegrationRecord[]{newRecord});

            long newRecordId = extractRecordId(rhh); // TODO: confirm how to pull ID off rhh

            byte[] bytes = Base64.getDecoder().decode(attachment.contentBytes);

            DataHandler dh = new DataHandler(new ByteArrayDataSource(bytes, attachment.contentType));
            Content content = new Content();
            content.setContent(dh);
            content.setRecordId(newRecordId);
            content.setFileName(attachment.name);
            content.setFieldName("Content");


            tws.upload(content);

            log.info("creating association between {} and {}", recordId, newRecordId);
            Association association = new Association();

            association.setAssociatedRecordId(newRecordId);
            association.setRecordId(recordId);
            association.setAssociationName("Email Attachment");
            association.setReverseAssociationName("Email Attachment");

            tws.associateRecords(new Association[]{association});

            log.info("Uploaded attachment '" + attachment.name + "' (" + bytes.length + " bytes) to record " + recordId);

        } catch (Exception e) {
            log.error("Failed to create EmailAttachment record for: " + attachment.name, e);
        }

    }

    private void populateFieldsFromGraphMessage(GraphMessage msg, IntegrationSection section) {
        List<IntegrationField> fields = new ArrayList<>();

        fields.add(buildField("Subject", msg.subject));
        fields.add(buildField("Body", msg.body.content));
        fields.add(buildField("sentDate", formatDate(msg.sentDateTime)));
        fields.add(buildField("ReceivedDate", formatDate(msg.receivedDateTime)));

        log.info("setting sentDate to {}", msg.sentDateTime);
        log.info("setting receivedDate to {}", msg.receivedDateTime);
        section.setFields(fields.toArray(new IntegrationField[0]));
    }

    private void populateFieldsFromGraphAttachment(GraphEmailAttachment attachment, IntegrationSection section) {
        List<IntegrationField> fields = new ArrayList<>();

        fields.add(buildField("FileName", attachment.name));
        section.setFields(fields.toArray(new IntegrationField[0]));
    }

    private void populateFieldsFromGraphAddress(GraphPerson recipient, IntegrationSection section) {
        List<IntegrationField> fields = new ArrayList<>();

        fields.add(buildField("PersonalName", recipient.emailAddress.name));
        fields.add(buildField("Address", recipient.emailAddress.address));
        section.setFields(fields.toArray(new IntegrationField[0]));
    }

    private IntegrationField buildField(String name, String value) {
        IntegrationField f = new IntegrationField();
        f.setName(name);
        f.setValue(value != null ? value : "");
        return f;
    }

    private String formatDate(Object dateValue) {
        // TODO: confirm exact TRIRIGA date string format expected, e.g. "MM/dd/yyyy HH:mm:ss"
        if (dateValue == null) {
            return "";
        }
        try {
            Date parsed;
            if (dateValue instanceof Date) {
                parsed = (Date) dateValue;
            } else {
                // Graph returns ISO 8601 strings like "2026-06-22T14:30:00Z"
                parsed = javax.xml.datatype.DatatypeFactory.newInstance()
                        .newXMLGregorianCalendar(String.valueOf(dateValue))
                        .toGregorianCalendar()
                        .getTime();
            }
            return new SimpleDateFormat(TRIRIGA_DATE_FORMAT).format(parsed);
        } catch (Exception e) {
            log.warn("Failed to format date value: " + dateValue, e);
            return "";
        }
    }

}


