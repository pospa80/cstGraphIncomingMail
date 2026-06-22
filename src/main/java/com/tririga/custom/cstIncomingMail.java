package com.tririga.custom;

import com.tririga.pub.workflow.CustomBusinessConnectTask;
import com.tririga.pub.workflow.CustomParamTaskResultImpl;
import com.tririga.pub.workflow.Record;
import com.tririga.ws.TririgaWS;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.core.config.Configurator;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.List;

import com.google.gson.Gson;

import org.apache.http.client.methods.HttpPatch;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.client.methods.CloseableHttpResponse;

import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import javax.net.ssl.SSLContext;

public class cstIncomingMail implements CustomBusinessConnectTask {
    private static final String TENANT_ID     = "aaaa";
    private static final String CLIENT_ID     = "bbbb";
    private static final String CLIENT_SECRET = "cccc";
    private static final String MAILBOX       = "dddd";

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

            for (GraphMessage email : emails) {
                String emailId = email.id;
                String subject = email.subject;
                String senderName = email.from.emailAddress.name;
                String senderAddress = email.from.emailAddress.address;
                String receivedDate = email.receivedDateTime;
                String bodyContent = email.body.content;
                String bodyType = email.body.contentType;

                System.out.println(bodyContent);
                markAsRead(accessToken, emailId);
                log.info("Marking email {} as read", email.id);
            }

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
                + "&$select=id,subject,body,from,receivedDateTime"
                + "&$top=25";

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

//    private void markAsRead(String accessToken, String messageId) throws Exception {
//        String patchUrl = "https://graph.microsoft.com/v1.0/users/" + MAILBOX
//                + "/messages/" + messageId;
//
//        String body = "{\"isRead\": true}";
//
//        allowPatch();
//        URL url = new URL(patchUrl);
//        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
//        conn.setRequestMethod("PATCH");
//        conn.setRequestProperty("Authorization", "Bearer " + accessToken);
//        conn.setRequestProperty("Content-Type", "application/json");
//        conn.setRequestProperty("Accept", "application/json");
//        conn.setDoOutput(true);
//
//        try (OutputStream os = conn.getOutputStream()) {
//            os.write(body.getBytes("UTF-8"));
//        }
//
//        int responseCode = conn.getResponseCode();
//        if (responseCode != 200) {
//            throw new Exception("Failed to mark email as read. Response code: " + responseCode);
//        }
//        conn.disconnect();
//    }

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

}


